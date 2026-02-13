#!/usr/bin/env python3
"""
Automatisk vakansberäkning

Processes sick leave reports and payslips to calculate:
- Karens (waiting period) deductions
- OB (unsocial hours) classifications
- Vacant shift compensation

Key improvements:
- Dynamic PDF page detection
- Robust error handling
- Configurable settings
- Detailed logging
- Progress tracking
- Tracks 4320 (sjuklön dag -14) date ranges to correctly
  identify paid sick days that follow the karens day
"""

APP_VERSION = "2026-02-08 18:15 CET"

import re
import os
import logging
from pathlib import Path
from datetime import datetime, timedelta, time, date
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass

import yaml
import pandas as pd
import pdfplumber

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@dataclass
class Config:
    """Configuration for the karens calculator"""
    holidays: set[date]      # Regular holidays → Helg OB
    storhelg: set[date]      # Storhelg holidays → Storhelg OB (Påsk, Midsommar, Jul, Nyår)
    sick_list_header_pattern: str = r"Sjuklista\s+(\w+)\s+(\d{4})"
    sick_row_pattern: str = r"^\s*(\d{1,2})\s+(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\s+(\d+,\d+)\s+(.*)$"
    payslip_anst_pattern: str = r"Anställningsnr\s*:\s*(\d+)"
    karens_codes: List[str] = None
    sick_day_pattern: str = r"432(?:0)?"  # Code for sjuklön dag -14 (days 2-14): matches 432 and 4320
    gt14_pattern: str = r"dag\s*15--"
    
    def __post_init__(self):
        if self.karens_codes is None:
            self.karens_codes = ["43100", "43101"]


CONFIG_PATH = Path(__file__).parent / "config.yaml"


def _parse_date_list(raw: List) -> List[date]:
    """Convert a list of date strings/date objects to date objects."""
    return [d if isinstance(d, date) else date.fromisoformat(str(d)) for d in raw]


def load_holidays_from_yaml(config_path: Path = CONFIG_PATH) -> Optional[Tuple[List[date], List[date]]]:
    """Load holidays and storhelg from config.yaml.

    Returns (holidays, storhelg) tuple, or None if file doesn't exist.
    """
    try:
        if not config_path.exists():
            return None
        with open(config_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        if not data or "holidays" not in data:
            return None
        holidays = _parse_date_list(data["holidays"])
        storhelg = _parse_date_list(data.get("storhelg", []))
        return holidays, storhelg
    except Exception as e:
        logger.warning(f"Could not load holidays from {config_path}: {e}")
        return None


def save_holidays_to_yaml(holidays: List[date], storhelg: Optional[List[date]] = None, config_path: Path = CONFIG_PATH):
    """Save holidays (and optionally storhelg) to config.yaml, preserving other config sections"""
    data = {}
    try:
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
    except Exception as e:
        logger.warning(f"Could not read existing config: {e}")

    data["holidays"] = sorted([d.isoformat() for d in holidays])
    if storhelg is not None:
        data["storhelg"] = sorted([d.isoformat() for d in storhelg])

    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
    logger.info(f"Saved {len(holidays)} holidays and {len(storhelg or [])} storhelg to {config_path}")


def load_berakningsar_rates(year: str, config_path: Path = CONFIG_PATH) -> Optional[Dict]:
    """Load cost rates for a given beräkningsår from config.yaml"""
    try:
        if not config_path.exists():
            return None
        with open(config_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        if not data or "berakningsar" not in data:
            return None
        rates = data["berakningsar"].get(str(year))
        if rates:
            logger.info(f"Loaded beräkningsår rates for {year}")
        else:
            logger.warning(f"No beräkningsår rates found for {year}")
        return rates
    except Exception as e:
        logger.warning(f"Could not load beräkningsår rates: {e}")
        return None


# Mapping from OB class names (used in code) to config.yaml rate keys
OB_RATE_KEYS = {
    "Sjuk jourers helg": "helg_jour",
    "Sjuk jourers vardag": "vardag_jour",
    "Storhelg": "helg_stor",
    "Helg": "helg",
    "Natt": "vardag_natt",
    "Kväll": "vardag_kvall",
}


def load_config(holidays: Optional[List[date]] = None, storhelg: Optional[List[date]] = None, config_path: Path = CONFIG_PATH) -> Config:
    """Load configuration, reading holidays and storhelg from config.yaml if not provided"""
    if holidays is None:
        result = load_holidays_from_yaml(config_path)
        if result is not None:
            holidays, storhelg_from_yaml = result
            if storhelg is None:
                storhelg = storhelg_from_yaml
    if holidays is None:
        # Fallback hardcoded defaults
        holidays = [
            date(2025, 5, 1), date(2025, 5, 29),  # May Day & Ascension
            date(2025, 6, 6),  # National Day
        ]
    if storhelg is None:
        storhelg = [
            date(2025, 12, 24), date(2025, 12, 25), date(2025, 12, 26),  # Christmas
            date(2025, 12, 31), date(2026, 1, 1),  # New Year
        ]
    return Config(holidays=set(holidays), storhelg=set(storhelg))


class SwedishDateHelper:
    """Helper for Swedish date operations"""
    
    MONTHS = {
        "januari": 1, "jan": 1,
        "februari": 2, "feb": 2,
        "mars": 3, "mar": 3,
        "april": 4, "apr": 4,
        "maj": 5,
        "juni": 6, "jun": 6,
        "juli": 7, "jul": 7,
        "augusti": 8, "aug": 8,
        "september": 9, "sep": 9,
        "oktober": 10, "okt": 10,
        "november": 11, "nov": 11,
        "december": 12, "dec": 12
    }
    
    @classmethod
    def parse_month_name(cls, name: str) -> int:
        """Convert Swedish month name to number"""
        return cls.MONTHS.get(name.lower(), datetime.now().month)
    
    @classmethod
    def is_holiday(cls, d: date, holidays: set[date]) -> bool:
        """Check if date is a Swedish holiday"""
        return d in holidays


class OBClassifier:
    """Classify time periods into OB (unsocial hours) categories"""

    def __init__(self, holidays: set[date], storhelg: Optional[set[date]] = None):
        self.holidays = holidays        # Regular holidays → Helg OB
        self.storhelg = storhelg or set()  # Storhelg holidays → Storhelg OB

    def classify(self, dt: datetime) -> str:
        """
        Classify a datetime into OB category

        Categories:
        - Storhelg: Storhelg dates (Påsk, Midsommar, Jul, Nyår) 00:00-24:00,
                    day-before-storhelg 19:00-24:00, day-after-storhelg 00:00-07:00
        - Helg: Saturday/Sunday 00:00-24:00, regular holidays 00:00-24:00,
                Friday 19:00-24:00, Monday 00:00-07:00,
                day-before-regular-holiday 19:00-24:00,
                day-after-regular-holiday 00:00-07:00
        - Natt: Weekday 22:00-06:00
        - Kväll: Weekday 19:00-22:00
        - Dag: Weekday 06:00-19:00 (no OB)
        """
        d = dt.date()
        t = dt.time()

        # Storhelg holidays — full day
        if d in self.storhelg:
            return "Storhelg"

        # Regular holidays — full day → Helg OB (not Storhelg)
        if d in self.holidays:
            return "Helg"

        # Regular weekends (Sat/Sun) — full day
        if d.weekday() >= 5:
            return "Helg"

        # Evening/night transitions (19:00+) into next day
        if t >= time(19, 0):
            next_day = d + timedelta(days=1)
            if next_day in self.storhelg:
                return "Storhelg"
            if next_day in self.holidays or d.weekday() == 4:  # Friday or day-before-holiday
                return "Helg"

        # Morning transitions (00:00-07:00) trailing from previous day
        if t < time(7, 0):
            prev_day = d - timedelta(days=1)
            if prev_day in self.storhelg:
                return "Storhelg"
            if prev_day in self.holidays or d.weekday() == 0:  # Monday or day-after-holiday
                return "Helg"

        # Night (22:00-06:00) — already handled Friday/Monday above
        if t >= time(22, 0) or t < time(6, 0):
            return "Natt"

        # Evening (19:00-22:00) — already handled Friday/day-before-holiday above
        if t >= time(19, 0):
            return "Kväll"

        # Daytime (06:00-19:00)
        return "Dag"


class PersonnummerParser:
    """Parse and normalize Swedish personnummer"""
    
    @staticmethod
    def parse_float_sv(s: str) -> float:
        """Parse Swedish float notation (comma as decimal)"""
        return float(s.replace(" ", "").replace("\xa0", "").replace(",", "."))
    
    @staticmethod
    def from_filename(path: str) -> Optional[str]:
        """Extract 12-digit personnummer from filename: YYMMDD-XXXX"""
        base = os.path.basename(path)
        m = re.search(r"(\d{6})-(\d{4})", base)
        if not m:
            return None
        yymmdd, ext = m.group(1), m.group(2)
        yy = int(yymmdd[:2])
        century = 1900 if yy >= 50 else 2000
        return f"{century+yy:04d}{yymmdd[2:]}{ext}"
    
    @staticmethod
    def normalize(pnr: str) -> str:
        """Normalize 10-digit to 12-digit personnummer"""
        if len(pnr) == 10:
            yy = int(pnr[:2])
            century = 1900 if yy >= 50 else 2000
            return f"{century+yy:04d}{pnr[2:]}"
        return pnr


class PayslipParser:
    """Parse payslip PDFs for employment data, karens, and GT14 periods"""
    
    def __init__(self, config: Config):
        self.config = config
    
    # Matches actual timlön salary lines like:
    #   "11 Timlön direkt sem.ersättning [5001EL] 139,5 tim 156,00 21 762,00"
    #   "114 Timlön direkt sem.ersättning, KOM [ZS] 6,00 tim 150,00"
    #   "Timlön exkl. sem.ersättning [013GA] 120,00 tim 193,00 23 160,00"
    # Requires either a code prefix (11, 114, …) immediately before "Timlön",
    # or "Timlön" at the start of the line. This avoids matching semester
    # settlement lines (2295/2296/2297) where "timlön" appears mid-description.
    TIMLON_PATTERN = re.compile(
        r"(?:\b11\d{0,2}\b\s+|^)Timlön.*?(\d+[,\.]\d+)\s*tim\s+(\d+[,\.]\d+)",
        re.IGNORECASE | re.MULTILINE,
    )

    # Fallback: extract timlön from sjuklön dag -14 line (4320 or 432).
    # The rate on that line is 80% of the real timlön.
    SJUKLON_TIMLON_PATTERN = re.compile(
        r"\b432(?:0)?\b.*?(\d+[,\.]\d+)\s*tim\s+(\d+[,\.]\d+)",
    )

    def parse_multiple(self, payslip_paths: List[str]) -> Tuple[Dict, Dict, Dict, Dict, Dict]:
        """
        Parse multiple payslip PDFs

        Returns:
            anst_map: pnr -> employment number
            karens_seconds: (pnr, date_str) -> seconds of karens
            gt14_ranges: pnr -> [(start_date, end_date), ...]
            sick_day_ranges: pnr -> [(start_date, end_date), ...]
            timlon_map: pnr -> hourly rate (float)
        """
        anst_map = {}
        karens_seconds = {}
        gt14_ranges = {}
        sick_day_ranges = {}  # Track 4320 (sjuklön dag -14) ranges
        timlon_map = {}  # pnr -> hourly rate
        
        for path in payslip_paths:
            if not os.path.exists(path):
                logger.warning(f"Payslip not found: {path}")
                continue
            
            try:
                pnr12 = PersonnummerParser.from_filename(path)
                if not pnr12:
                    logger.warning(f"Could not extract personnummer from: {path}")
                    continue
                
                logger.info(f"Processing payslip: {os.path.basename(path)}")
                
                with pdfplumber.open(path) as pdf:
                    text = "\n".join((p.extract_text() or "") for p in pdf.pages)
                
                # Extract employment number
                m_an = re.search(self.config.payslip_anst_pattern, text)
                if m_an:
                    anst_map[pnr12] = m_an.group(1)
                    logger.debug(f"  Employment nr: {m_an.group(1)}")
                
                # Extract karens periods (43100/43101)
                karens_count = 0
                for code in self.config.karens_codes:
                    pattern = rf"{code}[^\n]*?(\d+[,\.]\d+)\s*tim.*?\n(\d{{4}}-\d{{2}}-\d{{2}})\s*-\s*(\d{{4}}-\d{{2}}-\d{{2}})"
                    for m in re.finditer(pattern, text):
                        hrs = PersonnummerParser.parse_float_sv(m.group(1))
                        sec = hrs * 3600.0
                        d1 = datetime.fromisoformat(m.group(2)).date()
                        d2 = datetime.fromisoformat(m.group(3)).date()
                        if d1 == d2:
                            karens_seconds[(pnr12, d1.isoformat())] = sec
                            karens_count += 1
                
                if karens_count > 0:
                    logger.debug(f"  Found {karens_count} karens entries")
                
                # Extract sick day ranges (4320 - sjuklön dag -14)
                sick_day_count = 0
                sick_day_pattern = rf"{self.config.sick_day_pattern}[^\n]*\n(\d{{4}}-\d{{2}}-\d{{2}})\s*-\s*(\d{{4}}-\d{{2}}-\d{{2}})"
                for m in re.finditer(sick_day_pattern, text):
                    d1 = datetime.fromisoformat(m.group(1)).date()
                    d2 = datetime.fromisoformat(m.group(2)).date()
                    sick_day_ranges.setdefault(pnr12, []).append((d1, d2))
                    sick_day_count += 1
                
                if sick_day_count > 0:
                    logger.debug(f"  Found {sick_day_count} sick day ranges (4320)")
                
                # Extract GT14 periods (sick >14 days)
                gt14_pattern = self.config.gt14_pattern + r"[^\n]*(?:\n|\s)(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})"
                for m in re.finditer(gt14_pattern, text):
                    d1 = datetime.fromisoformat(m.group(1)).date()
                    d2 = datetime.fromisoformat(m.group(2)).date()
                    gt14_ranges.setdefault(pnr12, []).append((d1, d2))
                    logger.debug(f"  Found GT14 period: {d1} to {d2}")

                # Extract timlön (hourly rate) from 11* codes
                rates_found = set()
                for m_tim in self.TIMLON_PATTERN.finditer(text):
                    rate = PersonnummerParser.parse_float_sv(m_tim.group(2))
                    rates_found.add(rate)

                # Fallback: derive timlön from sjuklön 4320 line (rate is 80%)
                if not rates_found:
                    for m_sjk in self.SJUKLON_TIMLON_PATTERN.finditer(text):
                        sjk_rate = PersonnummerParser.parse_float_sv(m_sjk.group(2))
                        derived = round(sjk_rate / 0.8, 2)
                        rates_found.add(derived)
                        logger.debug(f"  Timlön fallback from 4320: {sjk_rate} kr (80%) -> {derived} kr (100%)")

                if rates_found:
                    primary_rate = max(rates_found)  # use highest as primary
                    multi = len(rates_found) > 1
                    # Keep highest rate across payslips; flag if any payslip had multiple
                    existing = timlon_map.get(pnr12)
                    if existing:
                        all_multi = existing["multi"] or multi or existing["rate"] != primary_rate
                        timlon_map[pnr12] = {
                            "rate": max(existing["rate"], primary_rate),
                            "multi": all_multi,
                        }
                    else:
                        timlon_map[pnr12] = {"rate": primary_rate, "multi": multi}
                    logger.debug(f"  Timlön: {rates_found} kr (primary={primary_rate}, multi={multi})")

            except Exception as e:
                logger.error(f"Error processing payslip {path}: {e}")

        logger.info(f"Processed {len(anst_map)} payslips successfully")
        return anst_map, karens_seconds, gt14_ranges, sick_day_ranges, timlon_map


class SickListParser:
    """Parse sick list PDFs"""

    # Regular row: "06  20:00 - 22:30  2,50  ..."
    REGULAR_PATTERN = re.compile(
        r"^\s*(\d{1,2})\s+(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\s+(\d+,\d+)\s+(.*)$"
    )
    # Jour row: "06                  22:30 - 00:00  1,50  ..."
    # Distinguished by large whitespace gap between day and time (>8 spaces)
    JOUR_PATTERN = re.compile(
        r"^\s*(\d{1,2})\s{8,}(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\s+(\d+,\d+)\s+(.*)$"
    )

    # Regex to extract start-end times from table cells like "22:30" "- 00:00"
    _TIME_RE = re.compile(r"(\d{2}:\d{2})")

    def __init__(self, config: Config):
        self.config = config

    # Pattern matching hours like "1,50" or "5,00"
    _HOURS_RE = re.compile(r"^\d+,\d+$")

    def _extract_jour_set(self, page) -> set:
        """
        Use table extraction on a sick list page to detect jour rows.

        The PDF table has separate columns for regular hours and jour
        hours on the left (Sjukskriven) side.  Jour rows are identified
        by: col 0 has only a day number (no time), followed by empty
        cells, then a time pair and hours further right.

        The column layout varies between pages (17 vs 21 cols), so
        detection is position-agnostic: we scan left-side columns for
        the structural jour pattern.

        Returns a set of (day: int, start: str, end: str) tuples that
        are jour rows, used to override the text-based detection which
        cannot distinguish jour from regular on summary pages.
        """
        jour_keys: set = set()
        try:
            tables = page.extract_tables(
                {"vertical_strategy": "text", "horizontal_strategy": "text"}
            )
            if not tables:
                return jour_keys

            for row in tables[0]:
                if len(row) < 8:
                    continue

                # Col 0 must have only a day number (no time merged in)
                cell0 = (row[0] or "").strip()
                day_m = re.match(r"^(\d{1,2})$", cell0)
                if not day_m:
                    continue
                day = int(day_m.group(1))

                # Cols 1-2 must be empty (regular rows have "-" and end time)
                if (row[1] and row[1].strip()) or (row[2] and row[2].strip()):
                    continue

                # Scan remaining left-side cols for times and hours.
                # Look in the first ~10 columns (left/Sjukskriven side).
                scan_limit = min(len(row), 10)
                times = []
                for ci in range(3, scan_limit):
                    cell = (row[ci] or "").strip()
                    for tm in self._TIME_RE.finditer(cell):
                        times.append(tm.group(1))

                if len(times) >= 2:
                    jour_keys.add((day, times[0], times[1]))

        except Exception as e:
            logger.debug(f"Table extraction for jour detection failed: {e}")

        return jour_keys

    def detect_sicklist_pages(self, pdf_path: str) -> List[int]:
        """Dynamically detect which pages contain sick list data"""
        pages = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    if re.search(self.config.sick_list_header_pattern, text):
                        pages.append(i)
                        logger.debug(f"Found sick list on page {i+1}")
        except Exception as e:
            logger.error(f"Error detecting sick list pages: {e}")

        return pages

    def _parse_row(self, line: str) -> Optional[Dict]:
        """
        Parse a single sick list line, detecting regular vs jour rows.
        Returns dict with parsed data and is_jour flag, or None if not a match.
        """
        # Try regular pattern first (day + short gap + time)
        m0 = self.REGULAR_PATTERN.match(line)
        is_jour = False

        if not m0:
            # Try jour pattern (day + long gap + time)
            m0 = self.JOUR_PATTERN.match(line)
            if not m0:
                return None
            is_jour = True

        day = int(m0.group(1))
        sick_start, sick_end = m0.group(2), m0.group(3)
        sick_hrs = PersonnummerParser.parse_float_sv(m0.group(4))
        rest = m0.group(5)

        # Extract personnummer or anställningsnr from Sjukskriven side
        # Try full personnummer first (10-12 digits), then short anställningsnr (3+ digits)
        mp = re.search(r"(\d{10,12})", rest)
        if not mp:
            mp = re.search(r"(\d{3,9})", rest)
            if not mp:
                return None

        raw_id = mp.group(1)
        sick_pnr = PersonnummerParser.normalize(raw_id) if len(raw_id) >= 10 else raw_id
        sick_name = rest[:mp.start()].strip()
        sick_name = re.sub(r"\s{2,}", " ", sick_name).strip()

        # Check if replacement is vacant
        tail = rest[mp.end():].strip()
        m1 = re.search(r"(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\s+(\d+,\d+)\s+(.*)$", tail)
        repl_is_vacant = False
        if m1:
            repl_rest = m1.group(4).strip()
            repl_is_vacant = "[vakant]" in repl_rest

        return {
            "day": day,
            "start": sick_start,
            "end": sick_end,
            "hours": sick_hrs,
            "pnr": sick_pnr,
            "name": sick_name,
            "is_vacant": repl_is_vacant,
            "is_jour": is_jour,
        }

    def parse_sick_rows(self, pdf_path: str, pages: Optional[List[int]] = None) -> pd.DataFrame:
        """
        Parse sick list rows from PDF

        Returns DataFrame with columns:
        - Personnummer, Namn, Datum, Start, Slut
        - Sjuk_timmar_rapport, Ersättare_vakant, Is_jour
        """
        if pages is None:
            pages = self.detect_sicklist_pages(pdf_path)
            if not pages:
                logger.warning("No sick list pages detected, trying all pages")
                with pdfplumber.open(pdf_path) as pdf:
                    pages = list(range(len(pdf.pages)))

        rows = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for pidx in pages:
                    if pidx >= len(pdf.pages):
                        logger.warning(f"Page {pidx} out of range")
                        continue

                    page_obj = pdf.pages[pidx]
                    text = page_obj.extract_text() or ""

                    # Extract month and year from header
                    mh = re.search(self.config.sick_list_header_pattern, text)
                    year = int(mh.group(2)) if mh else datetime.now().year
                    month_name = (mh.group(1).lower() if mh else "")
                    month = SwedishDateHelper.parse_month_name(month_name)

                    logger.info(f"Parsing page {pidx+1}: {month_name.capitalize()} {year}")

                    # Build jour lookup from table extraction (left side
                    # of PDF has separate columns for regular vs jour hours)
                    jour_set = self._extract_jour_set(page_obj)
                    if jour_set:
                        logger.info(f"  Detected {len(jour_set)} jour rows via table extraction")

                    # Parse each sick row
                    for line in text.splitlines():
                        # Debug: log lines starting with a day number that fail to parse
                        parsed = self._parse_row(line)
                        if not parsed:
                            stripped = line.strip()
                            if stripped and stripped[0].isdigit() and re.match(r"^\s*\d{1,2}\s", line):
                                logger.debug(f"  UNPARSED LINE: {line!r}")
                            continue

                        # Override jour flag from table extraction
                        key = (parsed["day"], parsed["start"], parsed["end"])
                        if key in jour_set:
                            parsed["is_jour"] = True

                        try:
                            dt = date(year, month, parsed["day"])
                            rows.append({
                                "Personnummer": parsed["pnr"],
                                "Namn": parsed["name"],
                                "Datum": dt,
                                "Start": parsed["start"],
                                "Slut": parsed["end"],
                                "Sjuk_timmar_rapport": parsed["hours"],
                                "Ersättare_vakant": parsed["is_vacant"],
                                "Is_jour": parsed["is_jour"],
                            })
                        except ValueError as e:
                            logger.warning(f"Invalid date: {year}-{month}-{parsed['day']}: {e}")

        except Exception as e:
            logger.error(f"Error parsing sick list: {e}")

        jour_count = sum(1 for r in rows if r.get("Is_jour"))
        logger.info(f"Parsed {len(rows)} sick leave entries ({jour_count} jour)")
        return pd.DataFrame(rows)


class SjuklonekostnaderParser:
    """Parse Sjuklönekostnader (sick leave cost) PDFs for karens and sick day data"""

    # Pattern to detect personnummer lines: "Name    YYYYMMDD-XXXX"
    PNR_PATTERN = re.compile(r"(\d{8})-(\d{4})")
    # Pattern to extract date range and hours: "2025-09-01 - 2025-09-02  9,03 tim"
    # The "tim" suffix is optional — some PDFs omit it
    RANGE_PATTERN = re.compile(
        r"(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})\s+(\d+[,\.]\d+)\s*(?:tim)?"
    )
    # Pattern for single-date lines: "2025-09-01  2,30 tim"
    SINGLE_DATE_PATTERN = re.compile(
        r"(\d{4}-\d{2}-\d{2})\s+(\d+[,\.]\d+)\s*(?:tim)?"
    )

    def __init__(self, config: Config):
        self.config = config

    @staticmethod
    def _classify_ob_from_description(desc_lower: str) -> Tuple[Optional[str], bool]:
        """
        Map Sjuklönekostnader line description to OB class.

        Returns (ob_class, is_supplement):
          - OB supplement lines (e.g. "Sjuk-OB helg dag -14") return ("Helg", True)
            These are NOT additional hours — they classify a subset of base hours.
          - Base lines (e.g. "Sjuklön timans. dag -14") return ("Dag", False)
            These ARE the actual hours worked.
          - Unrecognised lines return (None, False).
        """
        # Jour supplements (e.g. "Sjuk Jourtidsers. helgdag dag -14")
        if "jourtidsers" in desc_lower:
            if "helg" in desc_lower:
                return "Sjuk jourers helg", True
            return "Sjuk jourers vardag", True
        if "ob natt" in desc_lower:
            return "Natt", True
        if "ob kväll" in desc_lower or "ob kv" in desc_lower:
            return "Kväll", True
        if "ob storhelg" in desc_lower:
            return "Storhelg", True
        if "ob helg" in desc_lower:
            return "Helg", True
        # Base sjuklön line (no OB) — real hours
        if "sjuklön" in desc_lower and "ob" not in desc_lower:
            return "Dag", False
        return None, False

    # Pattern for "Summa <amount>" (NOT "Summa att betala")
    SUMMA_PATTERN = re.compile(r"^Summa\s+([\d\s]+[,\.]\d+)\s*$")

    # Pattern for semesterersättning lines with monetary amount:
    #   "Semesterersättning sjuklön 2025-04-02 - 2025-04-30 2 405,70"
    #   "Semesterersättning Sjuklön 2025-04-08 49,04"
    # The amount follows a date (YYYY-MM-DD) separated by spaces.
    # We match: date, then 1+ spaces, then the amount at end-of-line.
    SEM_ERS_AMOUNT_PATTERN = re.compile(
        r"\d{4}-\d{2}-\d{2}\s+([\d\s]+[,\.]\d+)\s*$"
    )

    def parse(self, pdf_path: str) -> Tuple[Dict, Dict, Dict, Dict, Dict, Dict, Dict]:
        """
        Parse a Sjuklönekostnader PDF.

        Returns:
            karens_seconds: (pnr, date_str) -> seconds of karens
            sick_day_ranges: pnr -> [(start_date, end_date), ...]
            total_hours_by_ob: pnr -> {ob_class: hours} (paid sjuklön only)
            karens_hours_by_pnr: pnr -> total karens hours
            base_hours_by_pnr: pnr -> total base paid hours (excl. supplements)
            summa_by_pnr: pnr -> total "Summa" amount from PDF (per employee)
            sem_ers_by_pnr: pnr -> total semesterersättning kr
        """
        karens_seconds: Dict[Tuple[str, str], float] = {}
        sick_day_ranges: Dict[str, List[Tuple[date, date]]] = {}
        total_hours_by_ob: Dict[str, Dict[str, float]] = {}
        karens_hours_by_pnr: Dict[str, float] = {}
        base_hours_by_pnr: Dict[str, float] = {}
        summa_by_pnr: Dict[str, float] = {}
        sem_ers_by_pnr: Dict[str, float] = {}

        if not os.path.exists(pdf_path):
            logger.warning(f"Sjuklönekostnader file not found: {pdf_path}")
            return karens_seconds, sick_day_ranges, total_hours_by_ob, karens_hours_by_pnr, base_hours_by_pnr, summa_by_pnr, sem_ers_by_pnr

        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        except Exception as e:
            logger.error(f"Error reading Sjuklönekostnader PDF: {e}")
            return karens_seconds, sick_day_ranges, total_hours_by_ob, karens_hours_by_pnr, base_hours_by_pnr, summa_by_pnr, sem_ers_by_pnr

        current_pnr = None
        brukare_pnr = None  # The brukare PNR from the page header (skip on continuation pages)
        expect_brukare_pnr = False  # Next PNR line is the brukare, not an employee
        karens_count = 0
        sick_range_count = 0

        for line in text.splitlines():
            # Detect brukare header line (marks next PNR as the brukare, not an employee)
            line_stripped = line.strip().lower()
            if line_stripped.startswith("brukare"):
                expect_brukare_pnr = True
                continue

            # Check for personnummer line
            m_pnr = self.PNR_PATTERN.search(line)
            if m_pnr:
                pnr_candidate = m_pnr.group(1) + m_pnr.group(2)
                if not re.match(r"^\s*\d{4}-\d{2}-\d{2}", line):
                    if expect_brukare_pnr:
                        # This is the brukare PNR from the header — skip it
                        brukare_pnr = pnr_candidate
                        expect_brukare_pnr = False
                        logger.debug(f"  Sjuklönekostnader: brukare PNR {brukare_pnr} (skipped)")
                        continue
                    if pnr_candidate == brukare_pnr:
                        # Brukare PNR reappearing on a continuation page — skip
                        logger.debug(f"  Sjuklönekostnader: brukare PNR {pnr_candidate} (continuation, skipped)")
                        continue
                    current_pnr = pnr_candidate
                    logger.debug(f"  Sjuklönekostnader: person {current_pnr}")
                    continue

            if not current_pnr:
                continue

            # Check for per-employee "Summa" line (NOT "Summa att betala")
            m_summa = self.SUMMA_PATTERN.match(line.strip())
            if m_summa:
                amount_str = m_summa.group(1).replace(" ", "")
                amount = PersonnummerParser.parse_float_sv(amount_str)
                summa_by_pnr[current_pnr] = amount
                logger.debug(f"    Summa for {current_pnr}: {amount}")
                continue

            line_lower = line.lower()

            # Semesterersättning lines — extract monetary amount per person
            if "semesterersättning" in line_lower:
                m_sem = self.SEM_ERS_AMOUNT_PATTERN.search(line)
                if m_sem and current_pnr:
                    sem_kr = PersonnummerParser.parse_float_sv(m_sem.group(1).replace(" ", ""))
                    sem_ers_by_pnr[current_pnr] = sem_ers_by_pnr.get(current_pnr, 0.0) + sem_kr
                continue

            # Try date range pattern first, then single date
            m_range = self.RANGE_PATTERN.search(line)
            m_single = None
            if m_range:
                d1 = datetime.fromisoformat(m_range.group(1)).date()
                d2 = datetime.fromisoformat(m_range.group(2)).date()
                hrs = PersonnummerParser.parse_float_sv(m_range.group(3))
            else:
                m_single = self.SINGLE_DATE_PATTERN.search(line)
                if not m_single:
                    continue
                d1 = datetime.fromisoformat(m_single.group(1)).date()
                d2 = d1
                hrs = PersonnummerParser.parse_float_sv(m_single.group(2))

            # GT14 lines: "dag 15--" (sick day 15 onwards, paid by Försäkringskassan)
            # Route to karens_hours_by_pnr for "Semesterers sjuklön (Karens och >14)"
            if "dag" in line_lower and "15--" in line_lower:
                karens_hours_by_pnr[current_pnr] = karens_hours_by_pnr.get(current_pnr, 0.0) + hrs
                logger.debug(f"    GT14 (dag 15--): {d1}-{d2} ({hrs}h)")
                continue

            # Karens lines: contain "karens" (e.g. "karenstid", "karens dir sem")
            # Must be checked BEFORE OB classification to separate karens hours
            if "karens" in line_lower:
                sec = hrs * 3600.0
                key = (current_pnr, d1.isoformat())
                if key not in karens_seconds:
                    karens_seconds[key] = sec
                    karens_count += 1
                    logger.debug(f"    Karens: {d1} ({hrs}h)")
                karens_hours_by_pnr[current_pnr] = karens_hours_by_pnr.get(current_pnr, 0.0) + hrs
                if d1 != d2:
                    next_day = d1 + timedelta(days=1)
                    sick_day_ranges.setdefault(current_pnr, []).append((next_day, d2))
                continue

            # Classify OB from line description and accumulate hours (paid sjuklön only)
            ob_class, is_supplement = self._classify_ob_from_description(line_lower)
            if ob_class and hrs > 0:
                ob_dict = total_hours_by_ob.setdefault(current_pnr, {})
                if is_supplement:
                    # OB supplement: reclassify hours from Dag to the specific OB class
                    ob_dict[ob_class] = ob_dict.get(ob_class, 0.0) + hrs
                    ob_dict["Dag"] = ob_dict.get("Dag", 0.0) - hrs
                else:
                    # Base line (Dag): these are the actual worked hours
                    ob_dict[ob_class] = ob_dict.get(ob_class, 0.0) + hrs
                    base_hours_by_pnr[current_pnr] = base_hours_by_pnr.get(current_pnr, 0.0) + hrs

            # Sick day -14 lines: contain "dag -14" or "dag-14" (days 2-14 of absence)
            if "dag" in line_lower and "-14" in line_lower:
                sick_day_ranges.setdefault(current_pnr, []).append((d1, d2))
                sick_range_count += 1
                logger.debug(f"    Sick day range: {d1} to {d2}")
                continue

        # Clamp Dag to 0 if OB supplements fully reclassified all base hours
        for pnr, ob_dict in total_hours_by_ob.items():
            if "Dag" in ob_dict and ob_dict["Dag"] < 0.01:
                ob_dict["Dag"] = max(0.0, ob_dict["Dag"])

        logger.info(
            f"Sjuklönekostnader: {karens_count} karens entries, "
            f"{sick_range_count} sick day ranges, "
            f"{len(total_hours_by_ob)} persons with OB hour data"
        )
        return karens_seconds, sick_day_ranges, total_hours_by_ob, karens_hours_by_pnr, base_hours_by_pnr, summa_by_pnr, sem_ers_by_pnr


class KarensCalculator:
    """Calculate karens (waiting period) and OB for vacant shifts"""
    
    def __init__(self, config: Config):
        self.config = config
        self.ob_classifier = OBClassifier(config.holidays, config.storhelg)
    
    def in_gt14(self, gt14_ranges: Dict, pnr: str, d: date) -> bool:
        """Check if date falls within a GT14 (>14 days sick) period"""
        for start, end in gt14_ranges.get(pnr, []):
            if start <= d <= end:
                return True
        return False
    
    def in_sick_day_range(self, sick_day_ranges: Dict, pnr: str, d: date) -> Tuple[bool, Optional[date]]:
        """
        Check if date falls within a sick day range (4320)
        Returns: (is_in_range, first_day_of_range)
        """
        for start, end in sick_day_ranges.get(pnr, []):
            if start <= d <= end:
                return True, start
        return False, None
    
    def parse_time(self, time_str: str) -> time:
        """Parse HH:MM time string"""
        h, m = map(int, time_str.split(":"))
        return time(h, m)
    
    def calculate_segments(
        self, 
        sick_df: pd.DataFrame,
        anst_map: Dict,
        karens_seconds: Dict,
        gt14_ranges: Dict,
        sick_day_ranges: Dict
    ) -> pd.DataFrame:
        """
        Calculate detailed segments with karens and OB classification
        
        Key logic:
        - Karens is consumed across ALL sick time on a date (vacant or not)
        - Output only shows VACANT segments
        - Each segment is split by OB class boundaries and karens cutoff
        - Uses sick_day_ranges to determine that days after karens day
          are "Betald" (not "underlag saknas")
        """
        detail_segments = []

        # Track karens remaining per person across dates so multi-day karens
        # carries over correctly (e.g. 8h karens spanning March 30-31).
        karens_remaining_by_pnr: Dict[str, Optional[float]] = {}

        for (pnr, d), grp in sick_df.groupby(["Personnummer", "Datum"]):
            pnr = str(pnr)
            d = pd.to_datetime(d).date()

            # Collect all intervals for this person-date
            intervals = []
            for _, r in grp.iterrows():
                start_dt = datetime.combine(d, self.parse_time(r["Start"]))
                end_dt = datetime.combine(d, self.parse_time(r["Slut"]))
                if end_dt <= start_dt:
                    end_dt += timedelta(days=1)
                is_jour = bool(r.get("Is_jour", False))
                intervals.append((start_dt, end_dt, bool(r["Ersättare_vakant"]), is_jour))

            intervals.sort(key=lambda x: x[0])

            # Determine karens status
            gt14 = self.in_gt14(gt14_ranges, pnr, d)
            ksec_total = karens_seconds.get((pnr, d.isoformat()), None)

            if ksec_total is not None and ksec_total > 0:
                # This date has a karens entry — start (or restart) the balance
                karens_remaining = ksec_total
                karens_remaining_by_pnr[pnr] = ksec_total
            elif pnr in karens_remaining_by_pnr and karens_remaining_by_pnr[pnr] is not None and karens_remaining_by_pnr[pnr] > 0:
                # Carry over remaining karens from a previous date
                karens_remaining = karens_remaining_by_pnr[pnr]
            else:
                karens_remaining = None
                # Check if date is in a sick day range
                # This means karens was on day 1 of the range, and this date is a continuation
                in_sick_range, range_start = self.in_sick_day_range(sick_day_ranges, pnr, d)
                if in_sick_range and range_start:
                    if (pnr, range_start.isoformat()) in karens_seconds:
                        # Karens was consumed on earlier day(s), so today is fully paid
                        karens_remaining = 0.0
                        logger.debug(f"  {pnr} on {d}: In sick range from {range_start}, karens already consumed")

            # Determine if this is the actual karens day (day 1 of sick period)
            is_karens_day = ksec_total is not None and ksec_total > 0

            # Process each interval and apply karens consumption
            interval_cuts = []
            for start_dt, end_dt, is_vacant, is_jour in intervals:
                if gt14:
                    interval_cuts.append((start_dt, end_dt, is_vacant, is_jour, 0.0, "GT14"))
                    continue

                if karens_remaining is None:
                    interval_cuts.append((start_dt, end_dt, is_vacant, is_jour, 0.0, "UNKNOWN"))
                    continue

                dur = (end_dt - start_dt).total_seconds()
                if karens_remaining <= 0:
                    mode = "PAID_DAY1" if is_karens_day else "PAID"
                    interval_cuts.append((start_dt, end_dt, is_vacant, is_jour, 0.0, mode))
                elif karens_remaining >= dur:
                    interval_cuts.append((start_dt, end_dt, is_vacant, is_jour, dur, "KARENS_FULL"))
                    karens_remaining -= dur
                else:
                    interval_cuts.append((start_dt, end_dt, is_vacant, is_jour, karens_remaining, "KARENS_PART"))
                    karens_remaining = 0.0

            # Save remaining karens for next date
            if karens_remaining is not None:
                karens_remaining_by_pnr[pnr] = karens_remaining

            # Create segments only for VACANT intervals
            for start_dt, end_dt, is_vacant, is_jour, karens_in_interval, mode in interval_cuts:
                if not is_vacant:
                    continue

                if is_jour:
                    # Jour segments: split at the helg/vardag boundary (06:00)
                    # so e.g. Mon 00:00-06:00 is helg (trailing Sunday) while
                    # Mon 06:00-08:00 is vardag.  Jour uses 06:00 boundary, not
                    # the regular OB 07:00 boundary.
                    jour_segs = self._split_jour_by_helg(
                        start_dt, end_dt, karens_in_interval, mode
                    )
                    for seg_start, seg_end, jour_ob, seg_status in jour_segs:
                        dur_sec = (seg_end - seg_start).total_seconds()
                        detail_segments.append({
                            "Anställningsnr": anst_map.get(pnr),
                            "Personnummer": pnr,
                            "Namn": grp["Namn"].iloc[0],
                            "Datum": seg_start.date().isoformat(),
                            "Start": seg_start.strftime("%H:%M"),
                            "Slut": seg_end.strftime("%H:%M"),
                            "Timmar": round(dur_sec / 3600.0, 4),
                            "OB-klass": jour_ob,
                            "Status": seg_status
                        })
                else:
                    # Regular segments: split by OB boundaries
                    segments = self._split_by_boundaries(
                        start_dt, end_dt, karens_in_interval, mode
                    )
                    for seg in segments:
                        detail_segments.append({
                            "Anställningsnr": anst_map.get(pnr),
                            "Personnummer": pnr,
                            "Namn": grp["Namn"].iloc[0],
                            "Datum": seg["start"].date().isoformat(),
                            "Start": seg["start"].strftime("%H:%M"),
                            "Slut": seg["end"].strftime("%H:%M"),
                            "Timmar": round(seg["hours"], 4),
                            "OB-klass": seg["ob_class"],
                            "Status": seg["status"]
                        })
        
        return pd.DataFrame(detail_segments)
    
    def _is_jour_helg(self, dt: datetime) -> bool:
        """Check if a datetime falls in jour-helg territory.

        Jour uses a 06:00 boundary (not 07:00 like regular OB):
        - Weekends, holidays, storhelg → helg all day
        - Mon 00:00-06:00 (trailing Sunday) → helg
        - Day-after-holiday 00:00-06:00 → helg
        - Fri 19:00-24:00 (leading into Saturday) → helg
        - Day-before-holiday 19:00-24:00 → helg
        """
        d = dt.date()
        t = dt.time()

        if d in self.config.storhelg or d in self.config.holidays:
            return True
        if d.weekday() >= 5:
            return True
        # Evening before helg day
        if t >= time(19, 0):
            nxt = d + timedelta(days=1)
            if nxt in self.config.storhelg or nxt in self.config.holidays or d.weekday() == 4:
                return True
        # Morning after helg day (06:00 boundary for jour)
        if t < time(6, 0):
            prev = d - timedelta(days=1)
            if prev in self.config.storhelg or prev in self.config.holidays or d.weekday() == 0:
                return True
        return False

    def _split_jour_by_helg(
        self,
        start_dt: datetime,
        end_dt: datetime,
        karens_in_interval: float,
        mode: str,
    ) -> List[Tuple]:
        """Split a jour interval at helg/vardag boundaries (06:00, 19:00, midnight).

        Returns list of (seg_start, seg_end, jour_ob_class, status) tuples.
        """
        result = []
        cur = start_dt
        while cur < end_dt:
            dcur = cur.date()
            # Potential transition points for jour helg/vardag
            boundaries = [
                datetime.combine(dcur, time(6, 0)),
                datetime.combine(dcur, time(19, 0)),
                datetime.combine(dcur + timedelta(days=1), time(0, 0)),
            ]
            nb = min([b for b in boundaries if b > cur] + [end_dt])

            is_helg = self._is_jour_helg(cur)
            jour_ob = "Sjuk jourers helg" if is_helg else "Sjuk jourers vardag"
            offset = (cur - start_dt).total_seconds()
            status = self._status_for_offset(mode, karens_in_interval, offset)
            result.append((cur, nb, jour_ob, status))
            cur = nb
        return result

    def _split_by_boundaries(
        self,
        start_dt: datetime,
        end_dt: datetime,
        karens_in_interval: float,
        mode: str
    ) -> List[Dict]:
        """Split interval by OB boundaries and karens cutoff"""
        segments = []
        cur = start_dt
        
        while cur < end_dt:
            # Calculate boundaries: day changes, OB hour changes, karens cutoff
            dcur = cur.date()
            boundaries = [
                datetime.combine(dcur + timedelta(days=1), time(0, 0)),
                datetime.combine(dcur, time(6, 0)),
                datetime.combine(dcur, time(7, 0)),
                datetime.combine(dcur, time(19, 0)),
                datetime.combine(dcur, time(22, 0)),
            ]
            
            # Add karens cutoff boundary
            if mode not in ("GT14", "UNKNOWN", "PAID", "PAID_DAY1") and karens_in_interval > 0:
                cutoff = start_dt + timedelta(seconds=karens_in_interval)
                if cur < cutoff < end_dt:
                    boundaries.append(cutoff)
            
            # Find next boundary
            nb = min([b for b in boundaries if b > cur] + [end_dt])
            
            # Determine status and OB class
            offset = (cur - start_dt).total_seconds()
            status = self._status_for_offset(mode, karens_in_interval, offset)
            ob_class = self.ob_classifier.classify(cur)
            
            dur_sec = (nb - cur).total_seconds()
            segments.append({
                "start": cur,
                "end": nb,
                "hours": dur_sec / 3600.0,
                "status": status,
                "ob_class": ob_class
            })
            
            cur = nb
        
        return segments
    
    def _status_for_offset(self, mode: str, karens_in_interval: float, offset_sec: float) -> str:
        """Determine payment status for a segment"""
        if mode == "GT14":
            return "Karens och >14"
        if mode == "UNKNOWN":
            return "Sjuklön dag 2-14"
        if mode == "PAID_DAY1":
            return "Sjuklön dag 1 - utanför karens"
        if mode == "PAID":
            return "Sjuklön dag 2-14"
        # KARENS_FULL or KARENS_PART
        if offset_sec < karens_in_interval:
            return "Karens"
        return "Sjuklön dag 1 - utanför karens"


class ReportGenerator:
    """Generate Excel reports from calculated segments"""
    
    @staticmethod
    def merge_adjacent_segments(detail: pd.DataFrame) -> pd.DataFrame:
        """Merge adjacent segments with same person/date/OB/status"""
        if detail.empty:
            return detail
        
        detail = detail.sort_values(
            ["Personnummer", "Datum", "Start", "OB-klass", "Status"]
        ).reset_index(drop=True)
        
        merged = []
        for _, r in detail.iterrows():
            if not merged:
                merged.append(r.to_dict())
                continue
            
            last = merged[-1]
            if (last["Personnummer"] == r["Personnummer"] and
                last["Datum"] == r["Datum"] and
                last["OB-klass"] == r["OB-klass"] and
                last["Status"] == r["Status"] and
                last["Slut"] == r["Start"]):
                # Merge with previous segment
                last["Slut"] = r["Slut"]
                last["Timmar"] = round(last["Timmar"] + r["Timmar"], 4)
            else:
                merged.append(r.to_dict())
        
        return pd.DataFrame(merged)
    
    @staticmethod
    def add_paid_hours_column(detail: pd.DataFrame) -> pd.DataFrame:
        """Add column for paid hours (vacant shifts)"""
        paid_statuses = {"Sjuklön dag 2-14", "Sjuklön dag 1 - utanför karens"}
        detail["Betalda timmar (vakant)"] = detail.apply(
            lambda r: r["Timmar"] if r["Status"] in paid_statuses else 0.0,
            axis=1
        )
        return detail
    
    @staticmethod
    def sort_chronologically(detail: pd.DataFrame) -> pd.DataFrame:
        """Sort segments chronologically"""
        detail["Datum_dt"] = pd.to_datetime(detail["Datum"])
        detail["Start_dt"] = pd.to_datetime(detail["Datum"] + " " + detail["Start"])
        detail = detail.sort_values(["Personnummer", "Datum_dt", "Start_dt"])
        return detail.drop(columns=["Datum_dt", "Start_dt"])
    
    @staticmethod
    def create_summary(df: pd.DataFrame, statuses: List[str]) -> pd.DataFrame:
        """Create summary by OB class for specific statuses"""
        dff = df[df["Status"].isin(statuses)].copy()
        if dff.empty:
            return pd.DataFrame(columns=[
                "Anställningsnr", "Personnummer", "Namn", "OB-klass", "Timmar"
            ])
        
        out = dff.groupby([
            "Anställningsnr", "Personnummer", "Namn", "OB-klass"
        ])["Timmar"].sum().reset_index()
        out["Timmar"] = out["Timmar"].round(2)
        return out.sort_values(["Personnummer", "OB-klass"])
    
    # All OB classes (including Dag) — used for totals
    OB_ROW_ORDER = [
        "Sjuk jourers helg",
        "Sjuk jourers vardag",
        "Storhelg",
        "Helg",
        "Natt",
        "Kväll",
        "Dag",
    ]

    # OB classes shown as individual rows (Dag excluded — folded into summary)
    OB_SECTION_ORDER = [
        "Sjuk jourers helg",
        "Sjuk jourers vardag",
        "Storhelg",
        "Helg",
        "Natt",
        "Kväll",
    ]

    # Display names for the employee sheet
    OB_DISPLAY_NAMES = {
        "Sjuk jourers helg": "Sjuk jourers helg",
        "Sjuk jourers vardag": "Sjuk jourers vardag",
        "Storhelg": "Storhelg",
        "Helg": "Helg",
        "Natt": "Natt: Vardagar från 22-06",
        "Kväll": "Kväll: Vardagar från 19-22",
        "Dag": "Dag",
    }

    # Statuses that represent paid sjuklön (used for Justering column)
    PAID_STATUSES = {"Sjuklön dag 1 - utanför karens", "Sjuklön dag 2-14"}

    # All day 1-14 statuses including karens (for "Sjuklön (timlön)" row)
    SJUKLON_STATUSES = {"Karens", "Sjuklön dag 1 - utanför karens", "Sjuklön dag 2-14"}

    # Status columns in the pivot (kept for reference)
    STATUS_COLUMNS = [
        ("Karens", "Vakant: Karens"),
        ("Sjuklön dag 1 - utanför karens", "Vakant: Sjuklön dag 1"),
        ("Sjuklön dag 2-14", "Vakant: Sjuklön dag 2-14"),
        ("Karens och >14", "Vakant: >14"),
    ]

    @staticmethod
    def create_employee_sheet_data(
        emp_detail: pd.DataFrame,
        sjk_hours: Dict[str, float],
        karens_hours: float = 0.0,
        base_hours: float = 0.0,
        timlon_rate: Optional[float] = None,
        rates: Optional[Dict] = None,
        under_23: bool = False,
        is_pensioner: bool = False,
    ) -> List[Dict]:
        """
        Build data for the per-employee sheet.

        Returns a list of dicts with keys:
          ob_class, display_name, sjk_timmar, sjk_kronor,
          justering_timmar, justering_kronor, netto_timmar, netto_kronor

        sjk_hours: paid sjuklön hours by OB class (karens excluded)
        karens_hours: total karens hours from sjuklönekostnader parser
        base_hours: total base paid hours from sjuklönekostnader (excl supplements)
        timlon_rate: 100% hourly rate for the employee
        rates: beräkningsår rates dict from config
        under_23: whether the employee is under 23

        Special ob_class values:
          "_summary"    — "Sjuklön (timlön)" totals (paid statuses only)
          "_sem_sjk"    — "Semesterersättning sjuklön" (semester_ersattning % of sjuklön)
          "_gt14"       — "Semesterers sjuklön (Karens och >14)"
          "_sjk_exkl"   — "Sjuklön exkl sem ers" subtotal
          "_sem_ers"    — "Semesterersättning" subtotal
          "_sjuklon"    — "Sjuklön" subtotal
          "_forsakring" — "Försäkringar"
          "_soc_avg"    — "Sociala avgifter"
          "_summa"      — "Summa Sjuklönekostnader"
        """
        sjklon_procent = rates.get("sjuklon_procent", 0.80) if rates else 0.80
        timlon_80 = round(timlon_rate * sjklon_procent, 2) if timlon_rate else 0.0
        rows = []

        # Gather total paid vacancy hours from sick list detail
        actual_total_just = round(
            emp_detail.loc[
                emp_detail["Status"].isin(ReportGenerator.PAID_STATUSES), "Timmar"
            ].sum(),
            2,
        )

        # Gather vacancy hours per OB class from sick list
        vacancy_by_ob: Dict[str, float] = {}
        for ob in ReportGenerator.OB_ROW_ORDER:
            mask = (
                (emp_detail["OB-klass"] == ob)
                & (emp_detail["Status"].isin(ReportGenerator.PAID_STATUSES))
            )
            vacancy_by_ob[ob] = round(emp_detail.loc[mask, "Timmar"].sum(), 2)

        # Distribute vacancy hours across OB rows.
        # Each OB class is capped at its sjuklönekostnader allocation.
        # Jour is now correctly detected via table extraction, so no
        # spill between OB classes is needed.
        just_by_ob: Dict[str, float] = {}
        for ob in ReportGenerator.OB_ROW_ORDER:
            sjk = round(sjk_hours.get(ob, 0.0), 2)
            vac = vacancy_by_ob.get(ob, 0.0)
            just_by_ob[ob] = round(min(sjk, vac), 2)

        # Helper to get OB rate from config (at sjuklön %)
        def ob_rate(ob_class: str) -> float:
            if not rates:
                return 0.0
            key = OB_RATE_KEYS.get(ob_class)
            return round(rates.get(key, 0.0) * sjklon_procent, 2) if key else 0.0

        # Build individual OB rows (without Dag)
        for ob in ReportGenerator.OB_SECTION_ORDER:
            sjk = round(sjk_hours.get(ob, 0.0), 2)
            just = just_by_ob.get(ob, 0.0)
            netto = round(max(0.0, sjk - just), 2)
            rate = ob_rate(ob)
            rows.append({
                "ob_class": ob,
                "display_name": ReportGenerator.OB_DISPLAY_NAMES[ob],
                "sjk_timmar": sjk,
                "sjk_kronor": round(sjk * rate, 2),
                "justering_timmar": just,
                "justering_kronor": round(just * rate, 2),
                "netto_timmar": netto,
                "netto_kronor": round(netto * rate, 2),
            })

        # "Sjuklön (timlön)" — use base hours (not sum of OB rows which includes supplements)
        # Jour vacancy hours are OB supplements with a separate rate and are
        # already accounted for on their own rows ("Sjuk jourers helg/vardag").
        # They must NOT be deducted again from the timlön summary.
        JOUR_OB_CLASSES = {"Sjuk jourers helg", "Sjuk jourers vardag"}
        jour_just = sum(
            just_by_ob.get(ob, 0.0) for ob in JOUR_OB_CLASSES
        )
        non_jour_just = round(actual_total_just - jour_just, 2)

        total_sjk = round(base_hours, 2)
        total_just = round(min(total_sjk, non_jour_just), 2)
        total_netto = round(max(0.0, total_sjk - total_just), 2)

        # "Sjuklön (timlön)" — paid sjuklön totals (dag 1 utanför karens + dag 2-14)
        rows.append({
            "ob_class": "_summary",
            "display_name": "Sjuklön (timlön)",
            "sjk_timmar": total_sjk,
            "sjk_kronor": round(total_sjk * timlon_80, 2),
            "justering_timmar": total_just,
            "justering_kronor": round(total_just * timlon_80, 2),
            "netto_timmar": total_netto,
            "netto_kronor": round(total_netto * timlon_80, 2),
        })

        # "Semesterersättning sjuklön" (semester_ersattning % of sjuklön, based on 100% timlön)
        sem_pct = rates.get("semester_ersattning", 0.12) if rates else 0.12
        timlon_100 = timlon_rate if timlon_rate else 0.0
        sem_sjk_sjk_kr = round(total_sjk * timlon_100, 2)
        sem_sjk_just_kr = round(total_just * timlon_100, 2)
        sem_sjk_netto_kr = round(total_netto * timlon_100, 2)
        rows.append({
            "ob_class": "_sem_sjk",
            "display_name": "Semesterersättning sjuklön",
            "sjk_timmar": sem_pct,
            "sjk_kronor": round(sem_sjk_sjk_kr * sem_pct, 2),
            "justering_timmar": sem_pct,
            "justering_kronor": round(sem_sjk_just_kr * sem_pct, 2),
            "netto_timmar": sem_pct,
            "netto_kronor": round(sem_sjk_netto_kr * sem_pct, 2),
        })

        # "Semesterers sjuklön (Karens och >14)"
        # sjk = karens hours from sjuklönekostnader (separated by parser)
        # justering = karens + >14 vacancy hours from detail
        # netto = max(0, sjk - justering)
        karens_sjk = round(karens_hours, 2)
        gt14_karens_just = round(
            min(
                karens_sjk,
                emp_detail.loc[
                    emp_detail["Status"].isin({"Karens", "Karens och >14"}), "Timmar"
                ].sum(),
            ),
            2,
        )
        sem_netto = round(max(0.0, karens_sjk - gt14_karens_just), 2)
        rows.append({
            "ob_class": "_gt14",
            "display_name": "Semesterers sjuklön (Karens och >14)",
            "sjk_timmar": karens_sjk,
            "sjk_kronor": round(karens_sjk * timlon_100 * sem_pct, 2),
            "justering_timmar": gt14_karens_just,
            "justering_kronor": round(gt14_karens_just * timlon_100 * sem_pct, 2),
            "netto_timmar": sem_netto,
            "netto_kronor": round(sem_netto * timlon_100 * sem_pct, 2),
        })

        # ── Cost summary rows ──
        # Salary rows = OB rows + _summary (timlön), excludes semesterersättning
        SALARY_CLASSES = {ob for ob in ReportGenerator.OB_SECTION_ORDER} | {"_summary"}
        SEM_CLASSES = {"_sem_sjk", "_gt14"}

        def sum_kr_by(classes, key: str) -> float:
            return round(sum(r[key] for r in rows if r["ob_class"] in classes), 2)

        # "Sjuklön exkl sem ers" = sum of salary parts only
        sjk_exkl_sjk = sum_kr_by(SALARY_CLASSES, "sjk_kronor")
        sjk_exkl_just = sum_kr_by(SALARY_CLASSES, "justering_kronor")
        sjk_exkl_netto = sum_kr_by(SALARY_CLASSES, "netto_kronor")
        rows.append({
            "ob_class": "_sjk_exkl",
            "display_name": "Sjuklön exkl sem ers",
            "sjk_timmar": None,
            "sjk_kronor": sjk_exkl_sjk,
            "justering_timmar": None,
            "justering_kronor": sjk_exkl_just,
            "netto_timmar": None,
            "netto_kronor": sjk_exkl_netto,
        })

        # "Semesterersättning" = sum of semesterersättning rows (_sem_sjk + _gt14)
        sem_ers_sjk = sum_kr_by(SEM_CLASSES, "sjk_kronor")
        sem_ers_just = sum_kr_by(SEM_CLASSES, "justering_kronor")
        sem_ers_netto = sum_kr_by(SEM_CLASSES, "netto_kronor")
        rows.append({
            "ob_class": "_sem_ers",
            "display_name": "Semesterersättning",
            "sjk_timmar": None,
            "sjk_kronor": sem_ers_sjk,
            "justering_timmar": None,
            "justering_kronor": sem_ers_just,
            "netto_timmar": None,
            "netto_kronor": sem_ers_netto,
        })

        # "Sjuklön" subtotal = sjk exkl + semesterersättning
        sjuklon_sjk = round(sjk_exkl_sjk + sem_ers_sjk, 2)
        sjuklon_just = round(sjk_exkl_just + sem_ers_just, 2)
        sjuklon_netto = round(sjk_exkl_netto + sem_ers_netto, 2)
        rows.append({
            "ob_class": "_sjuklon",
            "display_name": "Sjuklön",
            "sjk_timmar": None,
            "sjk_kronor": sjuklon_sjk,
            "justering_timmar": None,
            "justering_kronor": sjuklon_just,
            "netto_timmar": None,
            "netto_kronor": sjuklon_netto,
        })

        # "Försäkringar"
        forsakring_pct = 0.0
        if rates:
            if under_23:
                forsakring_pct = rates.get("forsakring_procent_under25", 0.0)
            else:
                forsakring_pct = rates.get("forsakring_procent", 0.0)
        rows.append({
            "ob_class": "_forsakring",
            "display_name": "Försäkringar",
            "sjk_timmar": forsakring_pct,
            "sjk_kronor": round(sjuklon_sjk * forsakring_pct, 2),
            "justering_timmar": forsakring_pct,
            "justering_kronor": round(sjuklon_just * forsakring_pct, 2),
            "netto_timmar": forsakring_pct,
            "netto_kronor": round(sjuklon_netto * forsakring_pct, 2),
        })

        # "Sociala avgifter" (reduced rate for pensioners 67+)
        if rates and is_pensioner:
            soc_avg_pct = rates.get("sociala_avgifter_pens", 0.1021)
        elif rates:
            soc_avg_pct = rates.get("sociala_avgifter", 0.3142)
        else:
            soc_avg_pct = 0.3142
        rows.append({
            "ob_class": "_soc_avg",
            "display_name": "Sociala avgifter",
            "sjk_timmar": soc_avg_pct,
            "sjk_kronor": round(sjuklon_sjk * soc_avg_pct, 2),
            "justering_timmar": soc_avg_pct,
            "justering_kronor": round(sjuklon_just * soc_avg_pct, 2),
            "netto_timmar": soc_avg_pct,
            "netto_kronor": round(sjuklon_netto * soc_avg_pct, 2),
        })

        # "Summa Sjuklönekostnader"
        forsakring_row = rows[-2]
        soc_avg_row = rows[-1]
        sjuklon_row = [r for r in rows if r["ob_class"] == "_sjuklon"][0]
        summa_sjk = round(sjuklon_row["sjk_kronor"] + forsakring_row["sjk_kronor"] + soc_avg_row["sjk_kronor"], 2)
        summa_just = round(sjuklon_row["justering_kronor"] + forsakring_row["justering_kronor"] + soc_avg_row["justering_kronor"], 2)
        summa_netto = round(sjuklon_row["netto_kronor"] + forsakring_row["netto_kronor"] + soc_avg_row["netto_kronor"], 2)
        rows.append({
            "ob_class": "_summa",
            "display_name": "Summa Sjuklönekostnader",
            "sjk_timmar": None,
            "sjk_kronor": summa_sjk,
            "justering_timmar": None,
            "justering_kronor": summa_just,
            "netto_timmar": None,
            "netto_kronor": summa_netto,
        })

        return rows

    # Row groups for Excel layout
    # OB rows: displayed with Timmar + Kronor
    # Summary section 1: _summary, _sem_sjk, _gt14 (Timmar + Kronor)
    # Blank row
    # Summary section 2: _sjk_exkl, _sem_ers, _sjuklon (Kronor only)
    # Blank row
    # Fees: _forsakring, _soc_avg (% + Kronor)
    # Blank row
    # Total: _summa (Kronor only)

    SUMMARY_SECTION_1 = {"_summary", "_sem_sjk", "_gt14"}
    SUMMARY_SECTION_2 = {"_sjk_exkl", "_sem_ers", "_sjuklon"}
    FEES_SECTION = {"_forsakring", "_soc_avg"}
    TOTAL_SECTION = {"_summa"}
    # Rows where Timmar column shows a percentage instead
    PERCENT_ROWS = {"_sem_sjk", "_forsakring", "_soc_avg"}
    # Rows that have no Timmar value at all
    NO_TIMMAR_ROWS = {"_sjk_exkl", "_sem_ers", "_sjuklon", "_summa"}

    @staticmethod
    def save_excel(
        detail: pd.DataFrame,
        output_path: str,
        sjk_total_hours: Optional[Dict[str, Dict[str, float]]] = None,
        sjk_karens_hours: Optional[Dict[str, float]] = None,
        sjk_base_hours: Optional[Dict[str, float]] = None,
        timlon_map: Optional[Dict[str, float]] = None,
        file_code: str = "",
        rates: Optional[Dict] = None,
        berakningsar: str = "",
        sjk_summa_by_pnr: Optional[Dict[str, float]] = None,
    ):
        """Save detailed report + per-employee sheets to Excel"""
        if sjk_total_hours is None:
            sjk_total_hours = {}
        if sjk_karens_hours is None:
            sjk_karens_hours = {}
        if sjk_base_hours is None:
            sjk_base_hours = {}
        if timlon_map is None:
            timlon_map = {}
        if sjk_summa_by_pnr is None:
            sjk_summa_by_pnr = {}

        # Parse file_code into brukare / period
        parts = file_code.split("_", 1)
        brukare = parts[0] if len(parts) >= 1 else ""
        period = parts[1] if len(parts) >= 2 else ""
        year = berakningsar or (period[:4] if len(period) >= 4 else "")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # Sheet 1: Full detail
            detail.to_excel(writer, sheet_name="Detalj", index=False)
            # Set Detalj column widths (25 for columns with content)
            ws_det = writer.sheets["Detalj"]
            for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
                ws_det.column_dimensions[col_letter].width = 25

            # Per-employee sheets — collect netto totals for summary
            used_names = set()
            summary_rows = []  # (nyckel, netto_kronor)
            for pnr in detail["Personnummer"].unique():
                emp = detail[detail["Personnummer"] == pnr]
                anst = emp["Anställningsnr"].iloc[0]
                sjk_hrs = sjk_total_hours.get(pnr, {})
                timlon_info = timlon_map.get(pnr)
                timlon_rate = timlon_info["rate"] if timlon_info else None

                emp_karens_hrs = sjk_karens_hours.get(pnr, 0.0)
                emp_base_hrs = sjk_base_hours.get(pnr, 0.0)

                # Determine under_23 and pensioner status from personnummer
                under_23 = False
                under_23_str = "nej"
                is_pensioner = False
                pensionar_str = "nej"
                if len(pnr) >= 8 and period:
                    try:
                        birth = date(int(pnr[:4]), int(pnr[4:6]), int(pnr[6:8]))
                        p_year = int(period[:4])
                        p_month = int(period[4:6])
                        if p_month == 12:
                            end_of_month = date(p_year + 1, 1, 1) - timedelta(days=1)
                        else:
                            end_of_month = date(p_year, p_month + 1, 1) - timedelta(days=1)
                        try:
                            birthday_23 = date(birth.year + 23, birth.month, birth.day)
                        except ValueError:
                            birthday_23 = date(birth.year + 23, 3, 1)
                        if birthday_23 > end_of_month:
                            under_23 = True
                            under_23_str = "ja"

                        # Pensioner: pension_age+ at start of beräkningsår (Jan 1)
                        pension_age = rates.get("pension_age", 67) if rates else 67
                        ber_year = int(year) if year and year.isdigit() else p_year
                        age_at_year_start = ber_year - birth.year
                        if age_at_year_start >= pension_age:
                            is_pensioner = True
                            pensionar_str = "ja"
                    except (ValueError, IndexError):
                        pass

                sheet_data = ReportGenerator.create_employee_sheet_data(
                    emp, sjk_hrs, emp_karens_hrs, emp_base_hrs,
                    timlon_rate=timlon_rate, rates=rates, under_23=under_23,
                    is_pensioner=is_pensioner,
                )

                # Capture justering total for summary sheet (vacancy cost)
                summa_row = next((r for r in sheet_data if r["ob_class"] == "_summa"), None)
                nyckel = f"{brukare}_{period}_{anst}" if anst else pnr
                just_total = summa_row["justering_kronor"] if summa_row else 0.0
                summary_rows.append((nyckel, just_total))

                # Build sheet name (max 31 chars for Excel)
                sheet_name = str(anst)[:31] if anst else pnr[:31]
                base_name = sheet_name
                counter = 2
                while sheet_name in used_names:
                    sheet_name = f"{base_name[:28]}_{counter}"
                    counter += 1
                used_names.add(sheet_name)

                ws = writer.book.create_sheet(sheet_name)

                sjklon_procent = rates.get("sjuklon_procent", 0.80) if rates else 0.80

                # ── Metadata (rows 1-11) ──
                ws["A1"] = "Brukare"
                ws["B1"] = brukare

                ws["A2"] = "Period"
                ws["B2"] = period

                ws["A3"] = "Anställd"
                ws["B3"] = anst or ""

                ws["A4"] = "Nyckel"
                ws["B4"] = f"{brukare}_{period}_{anst}" if anst else ""

                ws["A5"] = "Under 23"
                ws["B5"] = under_23_str

                pension_age_label = rates.get("pension_age", 67) if rates else 67
                ws["A6"] = f"Pensionär ({pension_age_label}+)"
                ws["B6"] = pensionar_str

                ws["A7"] = "Timlön (80%)"
                ws["B7"] = round(timlon_rate * sjklon_procent, 2) if timlon_rate else ""

                ws["A8"] = "Sjuklönprocent"
                if timlon_rate:
                    ws["B8"] = sjklon_procent
                    ws["B8"].number_format = '0.00%'

                ws["A9"] = "Timlön (100%)"
                ws["B9"] = timlon_rate if timlon_rate else ""

                ws["A10"] = "Beräkningsår"
                ws["B10"] = year if year else ""

                ws["A11"] = "Beräknare"
                ws["B11"] = "APP"

                # ── Table (rows 12+) ──
                # Column layout: A=label, B=sjk_timmar, C=sjk_kronor,
                #   D=just_timmar, E=just_kronor, F=netto_timmar, G=netto_kronor

                # Group headers (row 12)
                ws.cell(row=12, column=2, value="Enligt sjuklönekostnader")
                ws.cell(row=12, column=4, value="Justering för vakanser")
                ws.cell(row=12, column=6, value="Netto")

                # Sub-headers (row 14)
                ws.cell(row=14, column=2, value="Timmar")
                ws.cell(row=14, column=3, value="Kronor")
                ws.cell(row=14, column=4, value="Timmar")
                ws.cell(row=14, column=5, value="Kronor")
                ws.cell(row=14, column=6, value="Timmar")
                ws.cell(row=14, column=7, value="Kronor")

                # ── Write data rows ──
                row_num = 15

                # Helper to write one data row
                def write_row(item, rn):
                    oc = item["ob_class"]
                    ws.cell(row=rn, column=1, value=item["display_name"])

                    if oc in ReportGenerator.NO_TIMMAR_ROWS:
                        # Kronor only — no Timmar columns
                        ws.cell(row=rn, column=3, value=item.get("sjk_kronor", 0.0))
                        ws.cell(row=rn, column=5, value=item.get("justering_kronor", 0.0))
                        ws.cell(row=rn, column=7, value=item.get("netto_kronor", 0.0))
                    elif oc in ReportGenerator.PERCENT_ROWS:
                        # Timmar column shows percentage, plus Kronor
                        pct_val = item.get("sjk_timmar", 0.0)
                        ws.cell(row=rn, column=2, value=pct_val)
                        ws.cell(row=rn, column=2).number_format = '0.00%'
                        ws.cell(row=rn, column=3, value=item.get("sjk_kronor", 0.0))
                        ws.cell(row=rn, column=4, value=pct_val)
                        ws.cell(row=rn, column=4).number_format = '0.00%'
                        ws.cell(row=rn, column=5, value=item.get("justering_kronor", 0.0))
                        ws.cell(row=rn, column=6, value=pct_val)
                        ws.cell(row=rn, column=6).number_format = '0.00%'
                        ws.cell(row=rn, column=7, value=item.get("netto_kronor", 0.0))
                    else:
                        # Normal: Timmar + Kronor
                        ws.cell(row=rn, column=2, value=item.get("sjk_timmar", 0.0))
                        ws.cell(row=rn, column=3, value=item.get("sjk_kronor", 0.0))
                        ws.cell(row=rn, column=4, value=item.get("justering_timmar", 0.0))
                        ws.cell(row=rn, column=5, value=item.get("justering_kronor", 0.0))
                        ws.cell(row=rn, column=6, value=item.get("netto_timmar", 0.0))
                        ws.cell(row=rn, column=7, value=item.get("netto_kronor", 0.0))

                # OB rows (14-19)
                for item in sheet_data:
                    if not item["ob_class"].startswith("_"):
                        write_row(item, row_num)
                        row_num += 1

                # Blank separator
                row_num += 1

                # Summary section 1: Sjuklön (timlön), Semesterersättning sjuklön, Karens/GT14
                for item in sheet_data:
                    if item["ob_class"] in ReportGenerator.SUMMARY_SECTION_1:
                        write_row(item, row_num)
                        row_num += 1

                # Blank separator
                row_num += 1

                # Summary section 2: Sjuklön exkl, Semesterersättning, Sjuklön subtotal
                for item in sheet_data:
                    if item["ob_class"] in ReportGenerator.SUMMARY_SECTION_2:
                        write_row(item, row_num)
                        row_num += 1

                # Fees: Försäkringar, Sociala avgifter
                for item in sheet_data:
                    if item["ob_class"] in ReportGenerator.FEES_SECTION:
                        write_row(item, row_num)
                        row_num += 1

                # Blank separator
                row_num += 1

                # Total: Summa Sjuklönekostnader
                for item in sheet_data:
                    if item["ob_class"] in ReportGenerator.TOTAL_SECTION:
                        write_row(item, row_num)
                        row_num += 1

                # Validation: compare our Sjk Kronor total vs PDF "Summa"
                pdf_summa = sjk_summa_by_pnr.get(pnr)
                if pdf_summa is not None and summa_row is not None:
                    row_num += 1
                    our_total = round(summa_row["sjk_kronor"])
                    pdf_total = round(pdf_summa)
                    if our_total == pdf_total:
                        flag = "OK"
                    else:
                        flag = f"DIFF ({our_total} vs {pdf_total})"
                    ws.cell(row=row_num, column=1, value="Kontroll mot Sjuklönekostnader")
                    ws.cell(row=row_num, column=2, value=our_total)
                    ws.cell(row=row_num, column=3, value=pdf_total)
                    ws.cell(row=row_num, column=4, value=flag)

                # Set employee sheet column widths
                for col_letter in ["A", "B", "C", "D", "E", "F", "G"]:
                    ws.column_dimensions[col_letter].width = 25

            # ── Summary sheet: Vakanssammanfattning ──
            ws_sum = writer.book.create_sheet("Vakanssammanfattning", 0)  # insert first
            ws_sum["A1"] = "Vakanskostnader"
            row_num = 2
            grand_total = 0.0
            for nyckel, netto_kr in summary_rows:
                ws_sum.cell(row=row_num, column=1, value=nyckel)
                ws_sum.cell(row=row_num, column=2, value=netto_kr)
                grand_total += netto_kr
                row_num += 1
            row_num += 1
            ws_sum.cell(row=row_num, column=1, value="Totala vakanskostnader")
            ws_sum.cell(row=row_num, column=2, value=round(grand_total, 2))

            # Set Vakanssammanfattning column widths
            ws_sum.column_dimensions["A"].width = 25
            ws_sum.column_dimensions["B"].width = 25

        logger.info(f"Excel report saved: {output_path} ({len(used_names)} employee sheets)")


def process_karens_calculation(
    sick_pdf: str,
    payslip_paths: List[str],
    output_xlsx: str,
    config: Optional[Config] = None,
    sjuklonekostnader_path: Optional[str] = None,
    berakningsar_override: Optional[str] = None,
):
    """Main processing function"""
    if config is None:
        config = load_config()

    # Parse payslips
    payslip_parser = PayslipParser(config)
    anst_map, karens_seconds, gt14_ranges, sick_day_ranges, timlon_map = payslip_parser.parse_multiple(payslip_paths)

    # Parse Sjuklönekostnader (supplementary data, fills gaps from payslips)
    sjk_total_hours = {}
    sjk_karens_hours = {}
    sjk_base_hours = {}
    sjk_summa_by_pnr = {}
    sjk_sem_ers_by_pnr = {}
    if sjuklonekostnader_path:
        sjk_parser = SjuklonekostnaderParser(config)
        sjk_karens, sjk_sick_ranges, sjk_total_hours, sjk_karens_hours, sjk_base_hours, sjk_summa_by_pnr, sjk_sem_ers_by_pnr = sjk_parser.parse(sjuklonekostnader_path)

        # Merge: payslip data takes priority, sjuklönekostnader fills gaps
        for key, val in sjk_karens.items():
            if key not in karens_seconds:
                karens_seconds[key] = val
        for pnr, ranges in sjk_sick_ranges.items():
            sick_day_ranges.setdefault(pnr, []).extend(ranges)

        # Derive timlön from semesterersättning for employees without payslip rate
        # (e.g. long-term sick dag 15+ where employer pays no salary)
        # Formula: sem_ers_kr = timlön_100 × total_hours × semester_pct(12%)
        for pnr, sem_kr in sjk_sem_ers_by_pnr.items():
            if pnr not in timlon_map and sem_kr > 0:
                # Total hours = karens + base (all sick hour types for this person)
                total_hrs = sjk_karens_hours.get(pnr, 0.0) + sjk_base_hours.get(pnr, 0.0)
                if total_hrs > 0:
                    derived_rate = round(sem_kr / (total_hrs * 0.12), 2)
                    timlon_map[pnr] = {"rate": derived_rate, "multi": False}
                    logger.info(f"Derived timlön for {pnr} from semesterersättning: {sem_kr} / ({total_hrs} x 0.12) = {derived_rate} kr")

    # Parse sick list
    sicklist_parser = SickListParser(config)
    sick_df = sicklist_parser.parse_sick_rows(sick_pdf)

    if sick_df.empty:
        raise ValueError("Inga sjuklistrader hittades i PDF:en.")

    # Resolve anställningsnr to personnummer in sick list
    # anst_map is pnr -> anst_nr, build reverse: anst_nr -> pnr
    anst_to_pnr = {str(v): k for k, v in anst_map.items()}
    def resolve_pnr(val):
        val = str(val)
        if len(val) >= 10:
            return val
        return anst_to_pnr.get(val, val)
    sick_df["Personnummer"] = sick_df["Personnummer"].apply(resolve_pnr)

    # Calculate segments
    calculator = KarensCalculator(config)
    detail = calculator.calculate_segments(
        sick_df, anst_map, karens_seconds, gt14_ranges, sick_day_ranges
    )
    
    if detail.empty:
        raise ValueError("Inga vakanta segment hittades efter beräkning.")
    
    # Post-process
    detail = ReportGenerator.merge_adjacent_segments(detail)
    detail = ReportGenerator.add_paid_hours_column(detail)
    detail = ReportGenerator.sort_chronologically(detail)

    # Extract code from sick list filename (e.g. "Sjuklista_013_202405.pdf" -> "013_202405")
    sick_stem = Path(sick_pdf).stem  # e.g. "Sjuklista_013_202405"
    file_code = sick_stem.replace("Sjuklista", "", 1).lstrip("_") or ""
    detail.insert(0, "Kod", file_code)

    # Load beräkningsår rates for cost calculation
    if berakningsar_override:
        berakningsar = berakningsar_override
    else:
        parts = file_code.split("_", 1)
        period = parts[1] if len(parts) >= 2 else ""
        berakningsar = period[:4] if len(period) >= 4 else ""
    rates = load_berakningsar_rates(berakningsar) if berakningsar else None

    # Save
    ReportGenerator.save_excel(
        detail, output_xlsx,
        sjk_total_hours=sjk_total_hours,
        sjk_karens_hours=sjk_karens_hours,
        sjk_base_hours=sjk_base_hours,
        timlon_map=timlon_map,
        file_code=file_code,
        rates=rates,
        berakningsar=berakningsar,
        sjk_summa_by_pnr=sjk_summa_by_pnr,
    )


if __name__ == "__main__":
    import sys
    from pathlib import Path as _Path
    import glob as _glob

    # Separate --sjk and --year flags from positional args
    sjk_path = None
    year_override = None
    positional = []
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == "--sjk" and i + 1 < len(sys.argv):
            sjk_path = sys.argv[i + 1]
            i += 2
        elif sys.argv[i] == "--year" and i + 1 < len(sys.argv):
            year_override = sys.argv[i + 1]
            i += 2
        else:
            positional.append(sys.argv[i])
            i += 1

    if len(positional) < 2:
        print("Usage: python vakant_karens_app.py <sick_list.pdf> <payslip1.pdf> [payslip2.pdf ...] [--sjk <sjuklonekostnader.pdf>] [--year <YYYY>]")
        print("Output: vakansrapport.xlsx")
        print()
        print("The sjuklönekostnader PDF is auto-detected from the same directory if not specified.")
        print("The --year flag overrides the beräkningsår (default: derived from period).")
        sys.exit(1)

    sick_pdf = positional[0]
    payslips = positional[1:]

    # Auto-detect sjuklönekostnader PDF if not explicitly provided
    if not sjk_path:
        sick_dir = _Path(sick_pdf).parent
        # Look for files matching "Sjuklönekostnader*" or "Sjuklonekostnader*" in same dir
        for pattern in ["Sjuklönekostnader*.pdf", "Sjuklonekostnader*.pdf"]:
            candidates = list(sick_dir.glob(pattern))
            if candidates:
                sjk_path = str(candidates[0])
                logger.info(f"Auto-detected sjuklönekostnader: {sjk_path}")
                break

    # Derive output name from sick list filename
    stem = _Path(sick_pdf).stem
    suffix = stem.replace("Sjuklista", "", 1).lstrip("_")
    output_name = f"Vakansrapport_{suffix}.xlsx" if suffix else "Vakansrapport.xlsx"

    process_karens_calculation(sick_pdf, payslips, output_name, sjuklonekostnader_path=sjk_path, berakningsar_override=year_override)
