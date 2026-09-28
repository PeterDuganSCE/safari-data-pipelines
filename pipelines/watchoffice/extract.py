import html
import re
from datetime import datetime
from html.parser import HTMLParser
from typing import Any


UNKNOWN = "unknown"


def isolate_incident_narrative(text: str) -> str:
    """
    Isolate the relevant incident narrative from a complete Watch Office email.

    The narrative generally:
      1. Begins with "On <date> at <time> hours"
      2. Ends before the Watch Office signature, footer, CID, or disclaimer

    Returns an empty string if a narrative cannot be identified.
    """
    if not text or not isinstance(text, str):
        return ""

    # Normalize line endings and whitespace first.
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text).strip()

    # Locate the beginning of the incident narrative.
    start_match = re.search(
        r"\bOn\s+"
        r"(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+"
        r"\d{1,2}"
        r"(?:,\s*\d{4})?"
        r"\s+at\s+\d{3,4}\s+hours?\b",
        text,
        flags=re.IGNORECASE,
    )

    if not start_match:
        return ""

    narrative = text[start_match.start():]

        # Remove Watch Office signature / disclaimer section.
    footer_match = re.search(
        r"(?:Edison Watch Office|"
        r"Business Resiliency|"
        r"WatchOffice@SCE\.com|"
        r"\[cid:|"
        r"This email contains FERC Restricted information|"
        r"CONFIDENTIALITY NOTICE|"
        r"CAUTION:\s*External Email)",
        narrative,
        flags=re.IGNORECASE,
    )

    if footer_match:
        narrative = narrative[:footer_match.start()]

    # Convert the remaining narrative into a single clean paragraph.
    narrative = re.sub(r"\s+", " ", narrative).strip()

    return narrative


class _HTMLTextExtractor(HTMLParser):
    """Extract visible text while preserving superscript text."""

    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def get_text(self) -> str:
        return "".join(self.parts)


def _clean_text(raw_text: str) -> str:
    """
    Remove HTML markup, decode HTML entities, and normalize whitespace.
    """
    if not raw_text:
        return ""

    parser = _HTMLTextExtractor()
    parser.feed(html.unescape(raw_text))

    text = parser.get_text()

    # Convert ordinal fragments such as "42nd" into one complete token.
    text = re.sub(r"\s+(st|nd|rd|th)\b", r"\1", text, flags=re.IGNORECASE)

    # Normalize whitespace.
    return re.sub(r"\s+", " ", text).strip()


def _first_match(
    text: str,
    patterns: list[str],
    group: int | str = 1,
    flags: int = re.IGNORECASE,
) -> str:
    """
    Return the first successful regex match or UNKNOWN.
    """
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            value = match.group(group).strip(" ,.;:-")
            if value:
                return value

    return UNKNOWN


def _parse_number(value: str) -> int | str:
    """
    Convert values such as '2,000' to 2000.
    """
    if value == UNKNOWN:
        return UNKNOWN

    try:
        return int(value.replace(",", ""))
    except (TypeError, ValueError):
        return UNKNOWN


def _normalize_time(value: str) -> str:
    """
    Convert a four-digit military time to HH:MM.
    """
    if value == UNKNOWN:
        return UNKNOWN

    digits = re.sub(r"\D", "", value)

    if len(digits) != 4:
        return UNKNOWN

    try:
        return datetime.strptime(digits, "%H%M").strftime("%H:%M")
    except ValueError:
        return UNKNOWN

from datetime import datetime


from datetime import datetime
import re


def _normalize_incident_date(
    date_text: str,
    email_date: datetime | str | None = None
) -> str:
    """
    Convert incident dates to YYYY-MM-DD.

    Rules:
        1. If incident date contains a year, use it.
        2. Otherwise use the email year.
        3. If email was received in January and incident month is December,
           assume prior year.
        4. Fall back to current year if email_date is not provided.

    Examples:

        email_date = 2026-09-14
        September 5 -> 2026-09-05

        email_date = 2026-01-02
        December 31 -> 2025-12-31

        email_date = 2026-01-02
        January 1 -> 2026-01-01
    """

    if date_text == UNKNOWN:
        return UNKNOWN

    try:
        # Explicit year present in incident text
        if re.search(r"\b\d{4}\b", date_text):
            dt = datetime.strptime(date_text, "%B %d, %Y")
            return dt.strftime("%Y-%m-%d")

        # Convert email_date string if necessary
        if isinstance(email_date, str):
            try:
                email_date = datetime.fromisoformat(
                    email_date.replace("Z", "+00:00")
                )
            except Exception:
                email_date = None

        # Fallback if email_date not supplied
        if email_date is None:
            email_date = datetime.now()

        partial_dt = datetime.strptime(date_text, "%B %d")

        incident_month = partial_dt.month
        incident_day = partial_dt.day

        incident_year = email_date.year

        # Handle year crossover:
        # Email received Jan 1-31
        # Incident occurred in December
        if (
            email_date.month == 1
            and incident_month == 12
        ):
            incident_year -= 1

        dt = datetime(
            year=incident_year,
            month=incident_month,
            day=incident_day,
        )

        return dt.strftime("%Y-%m-%d")

    except Exception:
        return UNKNOWN

    
def _extract_hfra_status(text: str) -> str:
    """
    Return Yes, No, or unknown for HFRA status.
    """
    non_hfra_patterns = [
        r"\bis\s+not\s+(?:located\s+)?in\s+(?:a\s+)?high fire risk area\b",
        r"\bnot\s+in\s+(?:a\s+)?high fire risk area\b",
        r"\bnon[- ]HFRA\b",
        r"\boutside\s+(?:of\s+)?(?:an?\s+)?HFRA\b",
    ]

    hfra_patterns = [
        r"\bis\s+(?:located\s+)?in\s+(?:a\s+)?high fire risk area\b",
        r"\bwithin\s+(?:an?\s+)?HFRA\b",
        r"\bin\s+(?:an?\s+)?HFRA\b",
    ]

    if any(re.search(pattern, text, re.IGNORECASE)
           for pattern in non_hfra_patterns):
        return "No"

    if any(re.search(pattern, text, re.IGNORECASE)
           for pattern in hfra_patterns):
        return "Yes"

    return UNKNOWN


def _extract_fire_department(text: str) -> str:
    return _first_match(
        text,
        [
            r"extinguished by (?:the )?"
            r"((?:[A-Za-z'-]+\s+){1,6}Fire Department)"
            r"(?=[,.;]|$)",

            r"((?:[A-Za-z'-]+\s+){1,6}Fire Department)"
            r"\s+(?:responded|extinguished|suppressed)",
        ],
    )


def _extract_customers_disrupted(text: str) -> int | str:
    if re.search(r"\bno\s+(?:SCE\s+)?customers\s+(?:are|were)\s+affected\b", text, re.IGNORECASE):
        return 0
    count = _first_match(
        text,
        [r"\baffecting\s+(?:approximately\s+)?([\d,]+)\s+customers\b"],
    )
    return _parse_number(count)


def extract_incident_data(
        raw_text: str,
        email_date: datetime | str | None = None,
) -> dict[str, Any]:
    """
    Extract structured incident information from either:
      - a complete Watch Office email, or
      - an incident narrative by itself.
    """
    cleaned_email = _clean_text(raw_text)
    narrative = isolate_incident_narrative(cleaned_email)

    # If a recognizable narrative cannot be isolated, attempt to parse
    # the cleaned input. This supports inputs that do not start with "On".
    text = narrative or cleaned_email

    if not text:
        return {
            "incident_date": UNKNOWN,
            "incident_time": UNKNOWN,
            "circuit": UNKNOWN,
            "substation": UNKNOWN,
            "hfra": UNKNOWN,
            "fire_department": UNKNOWN,
            "incident_narrative": UNKNOWN,
        }

    incident_date = _normalize_incident_date(
        _first_match(
            text,
            [
                (
                    r"\bOn\s+"
                    r"((?:January|February|March|April|May|June|July|"
                    r"August|September|October|November|December)"
                    r"\s+\d{1,2}(?:,\s*\d{4})?)"
                    r"\s+at\b"
                ),
            ],
        ),
        email_date=email_date
    )

    incident_time = _first_match(
        text,
        [
            r"\bOn\s+.+?\s+at\s+(\d{3,4})\s*hours\b",
            r"\breported\s+at\s+(\d{3,4})\s*hours\b",
            r"\bat\s+(\d{3,4})\s*hours\b",
        ],
    )

    circuit = _first_match(
        text,
        [
            # "a portion of the Wyoming 12kV Circuit" -> "Wyoming"
            (
                r"\ba portion of the\s+"
                r"([A-Za-z0-9 .'-]+?)\s+"
                r"\d+(?:\.\d+)?\s*kV\s+Circuit\b"
            ),

            # "the Zappa 12kV Circuit" -> "Zappa"
            (
                r"\bthe\s+"
                r"([A-Za-z0-9 .'-]+?)\s+"
                r"\d+(?:\.\d+)?\s*kV\s+Circuit\b"
            ),

            # "Wyoming 12kV Circuit" -> "Wyoming"
            (
                r"\b([A-Za-z][A-Za-z0-9 .'-]*?)\s+"
                r"\d+(?:\.\d+)?\s*kV\s+Circuit\b"
            ),

            # Fallback when no voltage is present:
            # "the Wyoming Circuit out of" -> "Wyoming"
            (
                r"\bthe\s+"
                r"([A-Za-z0-9 .'-]+?)\s+Circuit\s+out of\b"
            ),
        ],
    )

    substation = _first_match(
        text,
        [
            # "out of Carolina Substation" -> "Carolina"
            r"\bout of\s+([A-Za-z0-9 .'-]+?)\s+Substation\b",

            # "from Quartz Hill Substation" -> "Quartz Hill"
            r"\bfrom\s+([A-Za-z0-9 .'-]+?)\s+Substation\b",
        ],
    )

    return {
        "incident_date": incident_date,
        "incident_time": _normalize_time(incident_time),
        "customers_disrupted": _extract_customers_disrupted(text),
        "circuit": circuit,
        "voltage_kv": _first_match(text, [r"\b(\d+(?:\.\d+)?)\s*kV\s+Circuit\b"]),
        "substation": substation,
        "hfra": _extract_hfra_status(text),
        "fire": "Yes" if re.search(r"\b(?:vegetation|spot|pole)\s+fire\b", text, re.IGNORECASE) else UNKNOWN,
        "fire_department": _extract_fire_department(text),
        "incident_narrative": text,
    }