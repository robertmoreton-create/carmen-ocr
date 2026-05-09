from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, List


PRIMARY_DATE_RE = r"(?P<transaction_date_raw>\d{1,2}/\d{1,2}(?:/\d{2,4})?)"
SECONDARY_DATE_RE = r"(?P<post_date_raw>\d{1,2}/\d{1,2}(?:/\d{2,4})?)"
AMOUNT_RE = r"(?P<amount>\(?-?\$?\d[\d,]*\.\d{2}\)?(?:\s?(?:CR|DR))?)"
TRANSACTION_RE = re.compile(
    rf"^\s*{PRIMARY_DATE_RE}(?:\s+{SECONDARY_DATE_RE})?\s+(?P<description>.+?)\s+{AMOUNT_RE}\s*$",
    re.IGNORECASE,
)
STATEMENT_DATE_RE = re.compile(
    r"(statement\s+(?:date|period\s+ending)|closing\s+date)[^\d]{0,15}(\d{1,2}/\d{1,2}/\d{2,4})",
    re.IGNORECASE,
)
LONG_DATE_RE = re.compile(
    r"\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})\b",
    re.IGNORECASE,
)
PAGE_DATE_RE = re.compile(
    r"Page\s+\d+\s+of\s+\d+.{0,300}?(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})",
    re.IGNORECASE | re.DOTALL,
)
MONTH_DATE_RE = re.compile(
    r"^(?P<day>\d{1,2})\s*(?P<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b",
    re.IGNORECASE,
)
EMBEDDED_MONTH_DATE_RE = re.compile(
    r"\((?P<day>\d{1,2})(?P<month>JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)(?P<year>\d{2})\)",
    re.IGNORECASE,
)

IGNORE_LINE_PATTERNS = [
    re.compile(pattern, re.IGNORECASE)
    for pattern in [
        r"^page\s+\d+",
        r"^total\s",
        r"^new\s+balance",
        r"^previous\s+balance",
        r"^credit\s+limit",
        r"^minimum\s+payment",
        r"^payment\s+due",
        r"^account\s+number",
        r"^customer\s+service",
        r"^interest\s+charge",
        r"^fees?\s+charged",
        r"^transactions?$",
        r"^date\s+description",
    ]
]


@dataclass
class Transaction:
    transaction_date: str
    post_date: str
    description: str
    amount: float
    raw_line: str
    source_file: str
    currency: str = "HKD"


@dataclass
class ParseResult:
    transactions: List[Transaction]
    review_lines: List[str]
    statement_date: str
    opening_balances: dict[str, float] = field(default_factory=dict)
    opening_balance_date: str = ""


@dataclass
class _HsbcParsedTransaction:
    transaction: Transaction
    balance: float | None


def parse_statement_text(text: str, source_file: Path) -> ParseResult:
    normalized_lines = _normalize_lines(text)
    statement_date = _find_statement_date(text)
    hsbc_transactions = _parse_hsbc_one_transactions(normalized_lines, statement_date, source_file)
    if hsbc_transactions:
        opening_date, opening_balances = _parse_hsbc_opening_balances(normalized_lines, statement_date)
        return ParseResult(
            transactions=hsbc_transactions,
            review_lines=[],
            statement_date=statement_date,
            opening_balances=opening_balances,
            opening_balance_date=opening_date,
        )

    transactions: List[Transaction] = []
    review_lines: List[str] = []

    for line in normalized_lines:
        if _should_ignore_line(line):
            continue

        match = TRANSACTION_RE.match(line)
        if match:
            groups = match.groupdict()
            transactions.append(
                Transaction(
                    transaction_date=_normalize_date(groups["transaction_date_raw"], statement_date),
                    post_date=_normalize_date(groups.get("post_date_raw") or "", statement_date),
                    description=_clean_description(groups["description"]),
                    amount=_parse_amount(groups["amount"]),
                    raw_line=line,
                    source_file=source_file.name,
                )
            )
            continue

        if _looks_like_transaction_candidate(line):
            review_lines.append(line)

    return ParseResult(
        transactions=transactions,
        review_lines=review_lines,
        statement_date=statement_date,
    )


def _normalize_lines(text: str) -> List[str]:
    lines: List[str] = []
    for raw_line in text.splitlines():
        compact = " ".join(raw_line.replace("\t", " ").split())
        if compact:
            lines.append(compact)
    return lines


def _should_ignore_line(line: str) -> bool:
    return any(pattern.search(line) for pattern in IGNORE_LINE_PATTERNS)


def _looks_like_transaction_candidate(line: str) -> bool:
    has_date = re.search(r"\d{1,2}/\d{1,2}(?:/\d{2,4})?", line) is not None
    has_amount = re.search(AMOUNT_RE, line, re.IGNORECASE) is not None
    return has_date and has_amount


def _clean_description(description: str) -> str:
    description = re.sub(r"\s{2,}", " ", description).strip(" -")
    return description


def _parse_amount(raw_amount: str) -> float:
    value = raw_amount.upper().replace("$", "").replace(",", "").strip()
    is_credit = value.endswith("CR")
    is_debit = value.endswith("DR")
    value = value.removesuffix("CR").removesuffix("DR").strip()

    negative = value.startswith("-") or value.startswith("(") or is_credit
    value = value.strip("()")
    amount = float(value)
    if negative and amount > 0:
        amount *= -1
    if is_debit and amount < 0:
        amount *= -1
    return amount


def _find_statement_date(text: str) -> str:
    match = STATEMENT_DATE_RE.search(text)
    if match:
        return _normalize_date(match.group(2), "")

    match = PAGE_DATE_RE.search(text)
    if match:
        return _normalize_date(match.group(1), "")

    match = LONG_DATE_RE.search(text)
    if match:
        return _normalize_date(match.group(1), "")

    return ""


def _normalize_date(raw_date: str, statement_date: str) -> str:
    raw_date = raw_date.strip()
    if not raw_date:
        return ""

    explicit_year_formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d %B %Y",
        "%d %b %Y",
    ]
    for date_format in explicit_year_formats:
        try:
            parsed = datetime.strptime(raw_date, date_format)
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            continue

    implied_year = _year_from_statement_date(statement_date)
    partial_formats = ["%m/%d", "%d/%m"]
    for date_format in partial_formats:
        try:
            parsed = datetime.strptime(f"{raw_date}/{implied_year}", f"{date_format}/%Y")
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            continue

    return raw_date


def _year_from_statement_date(statement_date: str) -> int:
    if statement_date:
        try:
            return datetime.strptime(statement_date, "%Y-%m-%d").year
        except ValueError:
            pass
    return datetime.today().year


def _parse_hsbc_one_transactions(
    lines: List[str], statement_date: str, source_file: Path
) -> List[Transaction]:
    if not any("HSBC One" in line for line in lines):
        return []

    parsed_transactions: List[_HsbcParsedTransaction] = []
    current_date = ""
    current_parts: List[str] = []
    current_balance: float | None = None

    def flush_current() -> None:
        nonlocal current_parts, current_balance
        if not current_date or not current_parts:
            current_parts = []
            return

        balance = _extract_hsbc_balance(" ".join(current_parts))
        transaction = _build_hsbc_transaction(current_date, current_parts, statement_date, source_file)
        if transaction:
            parsed_transactions.append(_HsbcParsedTransaction(transaction, balance))
        elif current_balance is None:
            current_balance = _extract_hsbc_opening_balance(" ".join(current_parts))
        current_parts = []

    in_hkd_history = False
    for line in lines:
        if not in_hkd_history:
            if line.lower().startswith("date transaction details"):
                in_hkd_history = True
            continue

        if line.lower().startswith("foreign currency savings") or line.startswith("["):
            break
        if _is_hsbc_noise_line(line):
            continue
        if _is_hsbc_header_line(line):
            continue

        date_match = MONTH_DATE_RE.match(line)
        if date_match:
            flush_current()
            current_date = _normalize_hsbc_date(date_match.group(0), statement_date)
            line = line[date_match.end() :].strip()

        if not line:
            continue

        if _starts_hsbc_transaction(line):
            flush_current()
            current_parts = [line]
        elif current_parts:
            current_parts.append(line)

    flush_current()
    transactions = _reconcile_hsbc_transactions(parsed_transactions, current_balance)
    transactions.extend(_parse_hsbc_foreign_currency_transactions(lines, statement_date, source_file))
    return transactions


def _parse_hsbc_opening_balances(lines: List[str], statement_date: str) -> tuple[str, dict[str, float]]:
    opening_date = ""
    balances: dict[str, float] = {}
    in_foreign_section = False

    for line in lines:
        if line.lower().startswith("foreign currency savings"):
            in_foreign_section = True

        currency = "HKD"
        balance_line = line
        currency_match = re.match(r"^(USD|GBP|JPY|CNY)\s+(.+)$", line, re.IGNORECASE)
        if currency_match:
            currency = currency_match.group(1).upper()
            balance_line = currency_match.group(2)
        elif in_foreign_section:
            continue

        if "BALANCE" not in balance_line.upper():
            continue

        date_match = MONTH_DATE_RE.match(balance_line)
        if date_match:
            normalized_date = _normalize_hsbc_date(date_match.group(0), statement_date)
            opening_date = opening_date or normalized_date

        balance = _extract_hsbc_opening_balance(balance_line)
        if balance is None and currency != "HKD":
            foreign_amounts = _extract_hsbc_foreign_amounts(balance_line)
            balance = foreign_amounts[-1] if foreign_amounts else None
        if balance is not None:
            balances[currency] = balance

    return opening_date, balances


def _is_hsbc_header_line(line: str) -> bool:
    lower = line.lower()
    return (
        "deposit" in lower and "withdrawal" in lower and "balance" in lower
    ) or lower.startswith(("ay ", "ani ", "date "))


def _is_hsbc_noise_line(line: str) -> bool:
    lower = line.lower()
    return lower.startswith(
        (
            "the hongkong and shanghai",
            "hsbc one }",
            "ane ",
            "ae ",
            "iess",
            "vess",
            "es",
            "hsbc",
            "number branch page",
            "pog ",
        )
    )


def _starts_hsbc_transaction(line: str) -> bool:
    upper = line.upper()
    starters = (
        "B/F BALANCE",
        "BYF BALANCE",
        "BAF BALANCE",
        "TO PAYME",
        "CASH REBATE",
        "POS MDC",
        "POSMDC",
        "MDC ",
        "CREDIT INTEREST",
    )
    return upper.startswith(starters)


def _build_hsbc_transaction(
    transaction_date: str, parts: List[str], statement_date: str, source_file: Path
) -> Transaction | None:
    raw_line = " ".join(parts)
    upper = raw_line.upper()
    if "BALANCE" in upper and not any(keyword in upper for keyword in ["CASH REBATE", "CREDIT INTEREST"]):
        return None

    amounts = _extract_hsbc_amounts(raw_line)
    if not amounts:
        return None

    amount = amounts[-1] if len(amounts) == 1 else amounts[-2]
    if _is_hsbc_withdrawal(upper):
        amount = -abs(amount)
    else:
        amount = abs(amount)

    return Transaction(
        transaction_date=_hsbc_embedded_transaction_date(raw_line, transaction_date),
        post_date="",
        description=_clean_hsbc_description(raw_line),
        amount=amount,
        raw_line=raw_line,
        source_file=source_file.name,
    )


def _parse_hsbc_foreign_currency_transactions(
    lines: List[str], statement_date: str, source_file: Path
) -> List[Transaction]:
    parsed_transactions: List[_HsbcParsedTransaction] = []
    in_foreign_section = False
    current_currency = ""
    current_date = ""
    opening_balances: dict[str, float] = {}

    for line in lines:
        if not in_foreign_section:
            if line.lower().startswith("foreign currency savings"):
                in_foreign_section = True
            continue

        if line.startswith("[") or line.lower().startswith("your average total relationship"):
            break
        if _is_hsbc_header_line(line):
            continue

        currency_match = re.match(r"^(USD|GBP|JPY|CNY)\s+(.+)$", line, re.IGNORECASE)
        if currency_match:
            current_currency = currency_match.group(1).upper()
            line = currency_match.group(2).strip()

        date_match = MONTH_DATE_RE.match(line)
        if date_match:
            current_date = _normalize_hsbc_date(date_match.group(0), statement_date)
            line = line[date_match.end() :].strip()

        if not current_currency or not current_date or not line:
            continue
        if "BALANCE" in line.upper():
            balance_amounts = _extract_hsbc_foreign_amounts(line)
            if balance_amounts:
                opening_balances[current_currency] = balance_amounts[-1]
            continue
        if not _starts_hsbc_foreign_transaction(line):
            continue

        transaction = _build_hsbc_foreign_transaction(
            current_date,
            current_currency,
            line,
            source_file,
        )
        if transaction:
            parsed_transactions.append(
                _HsbcParsedTransaction(
                    transaction=transaction,
                    balance=_extract_hsbc_foreign_balance(line),
                )
            )

    for currency, opening_balance in opening_balances.items():
        _reconcile_hsbc_transactions(
            [
                parsed
                for parsed in parsed_transactions
                if parsed.transaction.currency == currency
            ],
            opening_balance,
        )
    return [parsed.transaction for parsed in parsed_transactions]


def _starts_hsbc_foreign_transaction(line: str) -> bool:
    upper = line.upper()
    return upper.startswith(("CASH REBATE", "MDC "))


def _build_hsbc_foreign_transaction(
    transaction_date: str, currency: str, raw_line: str, source_file: Path
) -> Transaction | None:
    amounts = _extract_hsbc_foreign_amounts(raw_line)
    if not amounts:
        return None

    upper = raw_line.upper()
    amount = amounts[-1] if upper.startswith("CASH REBATE") else amounts[0]
    if upper.startswith("MDC "):
        amount = -abs(amount)

    return Transaction(
        transaction_date=transaction_date,
        post_date="",
        description=_clean_hsbc_foreign_description(raw_line),
        amount=amount,
        raw_line=raw_line,
        source_file=source_file.name,
        currency=currency,
    )


def _clean_hsbc_foreign_description(raw_line: str) -> str:
    upper = raw_line.upper()
    if upper.startswith("CASH REBATE"):
        return "CASH REBATE"

    description = re.sub(r"^MDC\s+[$S]*\s*", "", raw_line, flags=re.IGNORECASE)
    amount_match = re.search(r"(?<![\d.])\d[\d,]*(?![\d.]|[A-Za-z])", description)
    if amount_match:
        description = description[: amount_match.start()]
    description = re.sub(r"\b(?:AK|BK|BR|EK|FEK|FI|FR|HK|IK|PEK|RK|WK)\b\.?", "", description, flags=re.IGNORECASE)
    description = re.sub(r"\s+", " ", description)
    return description.strip(" -|")


def _extract_hsbc_foreign_amounts(raw_line: str) -> List[float]:
    amounts = [
        float(match.group(0).replace(",", ""))
        for match in re.finditer(r"(?<![\d.])\d[\d,]*(?![\d.]|[A-Za-z])", raw_line)
    ]
    return [amount for amount in amounts if amount != 1450]


def _extract_hsbc_foreign_balance(raw_line: str) -> float | None:
    amounts = _extract_hsbc_foreign_amounts(raw_line)
    if len(amounts) < 2:
        return None
    return amounts[-1]


def _reconcile_hsbc_transactions(
    parsed_transactions: List[_HsbcParsedTransaction], opening_balance: float | None
) -> List[Transaction]:
    if opening_balance is None:
        return [parsed.transaction for parsed in parsed_transactions]

    reconciled: List[Transaction] = []
    pending: List[_HsbcParsedTransaction] = []
    previous_balance = opening_balance

    for parsed in parsed_transactions:
        pending.append(parsed)
        if parsed.balance is None:
            continue

        _reconcile_hsbc_pending_group(pending, previous_balance, parsed.balance)
        reconciled.extend(item.transaction for item in pending)
        previous_balance = parsed.balance
        pending = []

    reconciled.extend(item.transaction for item in pending)
    return reconciled


def _reconcile_hsbc_pending_group(
    pending: List[_HsbcParsedTransaction], previous_balance: float, printed_balance: float
) -> None:
    parsed_total = round(sum(item.transaction.amount for item in pending), 2)
    expected_total = round(printed_balance - previous_balance, 2)
    if abs(parsed_total - expected_total) < 0.01:
        return

    for item in pending:
        corrected_amount = _correct_hsbc_amount_by_balance(
            item.transaction.amount,
            round(expected_total - (parsed_total - item.transaction.amount), 2),
        )
        if corrected_amount is None:
            continue

        item.transaction.amount = corrected_amount
        item.transaction.raw_line = (
            f"{item.transaction.raw_line} [balance-corrected amount: {corrected_amount:.2f}]"
        )
        return


def _correct_hsbc_amount_by_balance(parsed_amount: float, expected_amount: float) -> float | None:
    if parsed_amount == expected_amount:
        return parsed_amount
    if abs(parsed_amount - expected_amount) < 1:
        return None
    if parsed_amount * expected_amount < 0:
        return None

    parsed_digits = _amount_digits(parsed_amount)
    expected_digits = _amount_digits(expected_amount)
    if len(parsed_digits) != len(expected_digits):
        return None

    differences = sum(1 for left, right in zip(parsed_digits, expected_digits) if left != right)
    if differences == 1:
        return round(expected_amount, 2)
    return None


def _amount_digits(amount: float) -> str:
    return f"{abs(amount):.2f}".replace(".", "")


def _extract_hsbc_amounts(raw_line: str) -> List[float]:
    amount_pattern = re.compile(r"(?<![\d/])(?:[$S]|[^\w\s])?\d[\d,]*(?:[.,]\d{2})(?!\d)")
    amounts: List[float] = []
    for match in amount_pattern.finditer(raw_line):
        raw_amount = match.group(0)
        if raw_amount.startswith("$"):
            raw_amount = raw_amount[1:]
        elif raw_amount.startswith("S") or (raw_amount and not raw_amount[0].isdigit()):
            raw_amount = f"5{raw_amount[1:]}"
        if "." not in raw_amount and "," in raw_amount:
            head, tail = raw_amount.rsplit(",", 1)
            if len(tail) == 2:
                raw_amount = f"{head}.{tail}"
        amounts.append(float(raw_amount.replace(",", "")))
    return amounts


def _extract_hsbc_balance(raw_line: str) -> float | None:
    amounts = _extract_hsbc_amounts(raw_line)
    if len(amounts) < 2:
        return None
    return amounts[-1]


def _extract_hsbc_opening_balance(raw_line: str) -> float | None:
    if "BALANCE" not in raw_line.upper():
        return None
    amounts = _extract_hsbc_amounts(raw_line)
    if not amounts:
        return None
    return amounts[-1]


def _is_hsbc_withdrawal(upper_line: str) -> bool:
    return any(keyword in upper_line for keyword in ["TO PAYME", "POS MDC", "POSMDC"]) or upper_line.startswith("MDC ")


def _clean_hsbc_description(raw_line: str) -> str:
    description = re.sub(r"(?:[$S]|[^\w\s])?\d[\d,]*(?:[.,]\d{2})(?!\d)", "", raw_line)
    description = re.sub(r"\s+", " ", description)
    return description.strip(" |")


def _normalize_hsbc_date(raw_date: str, statement_date: str) -> str:
    date_match = MONTH_DATE_RE.match(raw_date.strip())
    if not date_match:
        return raw_date

    implied_year = _year_from_statement_date(statement_date)
    normalized_input = f"{date_match.group('day')} {date_match.group('month')} {implied_year}"
    return _normalize_date(normalized_input, statement_date)


def _hsbc_embedded_transaction_date(raw_line: str, fallback_date: str) -> str:
    upper = raw_line.upper()
    if "POS MDC" not in upper and "POSMDC" not in upper:
        return fallback_date

    match = EMBEDDED_MONTH_DATE_RE.search(raw_line)
    if not match:
        return fallback_date

    raw_date = f"{match.group('day')} {match.group('month')} 20{match.group('year')}"
    normalized = _normalize_date(raw_date, "")
    return normalized if normalized != raw_date else fallback_date
