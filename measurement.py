"""
measurement.py — Measurement-tietoluokka ja siihen liittyvät apufunktiot.

Sisältää yhden mittausrivin tietoluokan (Measurement) sekä funktiot eri
laitteiden mittausten yhdistämiseksi ja kirjoittamiseksi Excel-välilehdelle.
"""

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Optional


# ═══════════════════════════════════════════════════════════════════════════════
# MEASUREMENT — yhden mittausrivin tietoluokka
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class Measurement:
    """Yhden mittauskerran tiedot kaikilta kolmelta laitteelta.

    Yksi Measurement-olio = yksi rivi Excel-taulukossa.
    Sarakejärjestys on määritelty COLUMNS-luokkamuuttujassa.
    """
    timestamp: Optional[datetime] = None
    # Contour Care 7951H
    glucose_mmol: Optional[float] = None
    # Beurer BF 720 — paino
    weight:       Optional[float] = None
    # Omron M7
    systolic:     Optional[int]   = None
    diastolic:    Optional[int]   = None
    pulse:        Optional[int]   = None
    # Beurer BF 720 — muut terveystiedot
    fat:          Optional[float] = None
    water:        Optional[float] = None
    muscle:       Optional[float] = None
    bone:         Optional[float] = None
    bmr:          Optional[int]   = None
    amr:          Optional[int]   = None

    # Sarakeotsikot ja leveydet Excel-taulukkoa varten
    COLUMNS = [
        ("Päiväys ja kellonaika", 22),
        ("Verensokeri (mmol/L)",  22),
        ("Paino (kg)",            12),
        ("Systolinen (mmHg)",     18),
        ("Diastolinen (mmHg)",    18),
        ("Pulssi (/min)",         14),
        ("Rasva (%)",             12),
        ("Vesi (%)",              12),
        ("Lihas (%)",             12),
        ("Luumassa (kg)",         14),
        ("BMR (kcal)",            12),
        ("AMR (kcal)",            12),
    ]

    def to_row(self) -> list:
        """Palauttaa rivin arvoina samassa järjestyksessä kuin COLUMNS."""
        ts = self.timestamp.strftime("%d.%m.%Y %H:%M") if self.timestamp else ""
        return [
            ts,
            self.glucose_mmol,
            self.weight,
            self.systolic,
            self.diastolic,
            self.pulse,
            self.fat,
            self.water,
            self.muscle,
            self.bone,
            self.bmr,
            self.amr,
        ]

    def write_row(self, ws, row: int) -> None:
        """Kirjoittaa olion tiedot Excel-taulukon riville `row`."""
        from openpyxl.styles import Alignment
        for col, value in enumerate(self.to_row(), 1):
            cell = ws.cell(row=row, column=col, value=value)
            if col > 1:
                cell.alignment = Alignment(horizontal="center")
                if isinstance(value, float):
                    cell.number_format = "0.0"


# ═══════════════════════════════════════════════════════════════════════════════
# YHDISTÄMINEN — eri laitteiden mittaukset yhteen aikaikkunaan
# ═══════════════════════════════════════════════════════════════════════════════

def combine_measurements(glucose: list, bp: list, scale: list,
                         window_minutes: int = 60) -> list:
    """Yhdistää eri laitteiden mittaukset Measurement-olioiksi.

    Samaan aikaikkunaan (oletus 60 min) osuvat eri laitteiden mittaukset
    yhdistetään samaan Measurement-olioon. Mittaukset joille ei löydy paria
    saavat oman rivin.
    """
    measurements: list[Measurement] = []

    def find_or_create(ts: datetime) -> Measurement:
        if ts is None:
            m = Measurement(timestamp=None)
            measurements.append(m)
            return m
        for m in measurements:
            if m.timestamp and abs(m.timestamp - ts) <= timedelta(minutes=window_minutes):
                return m
        m = Measurement(timestamp=ts)
        measurements.append(m)
        return m

    for g in glucose:
        m = find_or_create(g.get("timestamp"))
        m.glucose_mmol = g.get("glucose_mmol")
        if m.timestamp is None and g.get("timestamp"):
            m.timestamp = g["timestamp"]

    for b in bp:
        m = find_or_create(b.get("timestamp"))
        m.systolic  = b.get("systolic")
        m.diastolic = b.get("diastolic")
        m.pulse     = b.get("pulse") or None
        if m.timestamp is None and b.get("timestamp"):
            m.timestamp = b["timestamp"]

    for s in scale:
        m = find_or_create(s.get("timestamp"))
        m.weight = s.get("weight")
        m.fat    = s.get("fat")
        m.water  = s.get("water")
        m.muscle = s.get("muscle")
        m.bone   = s.get("bone")
        m.bmr    = s.get("bmr")
        m.amr    = s.get("amr")
        if m.timestamp is None and s.get("timestamp"):
            m.timestamp = s["timestamp"]

    measurements.sort(key=lambda x: x.timestamp or datetime.min)
    return measurements


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL — "Mittaukset"-välilehden kirjoitus
# ═══════════════════════════════════════════════════════════════════════════════

def save_measurements_sheet(wb, measurements: list) -> None:
    """Kirjoittaa kaikki Measurement-oliot omalle 'Mittaukset'-välilehdelle."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("Mittaukset", 0)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, (header, width) in enumerate(Measurement.COLUMNS, 1):
        c = ws.cell(row=1, column=col, value=header)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="34495E")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    for ri, m in enumerate(measurements, 2):
        m.write_row(ws, ri)
        for col in range(1, len(Measurement.COLUMNS) + 1):
            ws.cell(ri, col).border = border

    if not measurements:
        ws.cell(2, 1, "Ei mittauksia.")


def append_measurement(filename: str, m: Measurement) -> None:
    """Lisää yksi Measurement-rivi olemassa olevaan 'Mittaukset'-välilehteen.

    Rivi kirjoitetaan ensimmäiselle vapaalle riville otsikkorivin jälkeen.
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side

    wb = load_workbook(filename)
    ws = wb["Mittaukset"]

    # Lasketaan, montako datariviä otsikon jälkeen on jo täytetty,
    # ja kirjoitetaan uusi rivi niiden jatkoksi.
    data_rows = 0
    for row in range(2, ws.max_row + 1):
        if any(ws.cell(row, col).value is not None
               for col in range(1, len(Measurement.COLUMNS) + 1)):
            data_rows += 1
    next_row = 2 + data_rows
    m.write_row(ws, next_row)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, len(Measurement.COLUMNS) + 1):
        ws.cell(next_row, col).border = border

    wb.save(filename)