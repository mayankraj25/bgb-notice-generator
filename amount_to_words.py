"""Convert a numeric rupee amount to Indian number system words."""

ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
]
TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _two_digits(n: int) -> str:
    if n < 20:
        return ONES[n]
    return (TENS[n // 10] + (" " + ONES[n % 10] if n % 10 else "")).strip()


def _three_digits(n: int) -> str:
    if n >= 100:
        rest = _two_digits(n % 100)
        return ONES[n // 100] + " Hundred" + (" " + rest if rest else "")
    return _two_digits(n)


def amount_to_words(amount: int) -> str:
    """
    Convert an integer amount (in rupees) to Indian number system words.
    e.g. 825000 -> 'Eight Lakh Twenty Five Thousand'
    """
    if amount == 0:
        return "Zero"

    parts = []

    crore = amount // 10_000_000
    amount %= 10_000_000

    lakh = amount // 100_000
    amount %= 100_000

    thousand = amount // 1000
    amount %= 1000

    remainder = amount

    if crore:
        parts.append(_three_digits(crore) + " Crore")
    if lakh:
        parts.append(_two_digits(lakh) + " Lakh")
    if thousand:
        parts.append(_two_digits(thousand) + " Thousand")
    if remainder:
        parts.append(_three_digits(remainder))

    return " ".join(parts)


if __name__ == "__main__":
    tests = [825000, 10000000, 15000, 100, 500000, 12345678]
    for t in tests:
        print(f"{t:>12,} -> {amount_to_words(t)}")
