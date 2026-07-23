# ============================================================
#  formatter.py
#  CSIR-NPL LCR Automation — Value Formatting Engine
#  Matches Agilent E4980A front-panel display EXACTLY
#
#  Verified sig figs (from real meter photo comparison):
#    SHORT → 6 significant figures
#    MED   → 7 significant figures
#    LONG  → 7 significant figures
#
#  SYNCED 2026-07-16 to match the inline copy in lcr_guii.py, which
#  is the one that actually runs. This file was dead code (never
#  imported) and had drifted to a stale, less-correct version of
#  the same logic. Kept in sync now purely so it isn't misleading —
#  the GUI does not import from here.
#
#  Core display rule (identical to E4980A firmware):
#    1. Scale raw value to best prefix (so display is 1.0–999.9)
#    2. Round to sig_figs significant figures
#    3. Decimal places = sig_figs - digits_before_decimal
# ============================================================

import math
from decimal import Decimal, ROUND_HALF_UP


def _round_half_up(x: float, ndigits: int) -> float:
    """Round like the E4980A firmware does (half-away-from-zero),
    not like Python's built-in round() (half-to-even/"banker's
    rounding"). The two only disagree on exact .5 boundaries, but
    when they do, this was producing a genuine off-by-one-digit
    mismatch against the panel even for an identical raw value."""
    q = Decimal(1).scaleb(-ndigits) if ndigits > 0 else Decimal(1)
    return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP))


# ============================================================
#  1. APERTURE → SIGNIFICANT FIGURES
#     Verified against real E4980A display:
#     LONG  showed  100.0081 pF  = 7 sig figs  ✓
# ============================================================
def get_sig_figs(aperture: str) -> int:
    return {
        "SHORT": 6,
        "MED":   7,
        "LONG":  7,
    }.get(aperture.strip().upper(), 7)


# ============================================================
#  2. INTERNAL: scale + round + format
# ============================================================
def _apply_prefix(value: float, sig_figs: int,
                  prefixes: list) -> str:
    """
    Pick the best prefix so the displayed number is >= 1.0 and < 1000.
    Then round to sig_figs and calculate exact decimal places.

    prefixes = list of (divisor, unit_string) in DESCENDING order
    e.g. [(1e-3,"mF"), (1e-6,"µF"), (1e-9,"nF"), (1e-12,"pF")]

    UNIT ROLLOVER FIX: choosing the prefix from the *unrounded* value
    and rounding afterwards breaks near every unit boundary — e.g.
    999.99962 nF picks "nF" (999.99962/1e-9 >= 1) but rounds to
    "1000.00 nF", which the instrument would instead roll over to
    "1.00000 µF". Fix: round first in the initially-chosen unit,
    then re-check whether the rounded value spilled into the next
    prefix up, and re-render there if so.
    """
    abs_val = abs(value)

    # Default to smallest prefix
    chosen_div, chosen_unit = prefixes[-1]
    chosen_idx = len(prefixes) - 1

    for idx, (div, unit) in enumerate(prefixes):
        scaled_check = abs_val / div
        if scaled_check >= 1.0:
            chosen_div  = div
            chosen_unit = unit
            chosen_idx  = idx
            break

    def _render(div, sig_figs_local):
        scaled = value / div
        if scaled == 0.0:
            return 0.0, f"0.{'0' * (sig_figs_local - 1)}"
        mag = math.floor(math.log10(abs(scaled)))
        decimal_places = max(0, sig_figs_local - 1 - int(mag))
        factor = 10 ** (sig_figs_local - 1 - int(mag))
        rounded = _round_half_up(scaled * factor, 0) / factor
        if rounded != 0.0:
            new_mag = math.floor(math.log10(abs(rounded)))
            if new_mag != mag:
                decimal_places = max(0, sig_figs_local - 1 - int(new_mag))
        return rounded, f"{rounded:.{decimal_places}f}"

    rounded, text = _render(chosen_div, sig_figs)

    if chosen_idx > 0 and abs(rounded) >= 1000.0 - 1e-9:
        next_div, next_unit = prefixes[chosen_idx - 1]
        if chosen_div != 0 and round(next_div / chosen_div) == 1000:
            chosen_div, chosen_unit = next_div, next_unit
            rounded, text = _render(chosen_div, sig_figs)

    return f"{text} {chosen_unit}"


# ============================================================
#  3. DIMENSIONLESS (D, Q) — no prefix, just decimal
# ============================================================
def _format_dimensionless(value: float, sig_figs: int,
                          param_type: str = "Q") -> str:
    """
    Format a dimensionless number (D or Q) to match E4980A display.

    D (Dissipation Factor):
      E4980A always shows D with FIXED 6 decimal places.
      e.g.  -0.000015  /  0.050000  /  0.002341
      (Verified from real meter photo)

    Q (Quality Factor):
      Large dimensionless number — use sig_figs significant figures.
      e.g.  120.500  /  5500.00
    """
    if param_type.upper() == "D":
        # D always shown to 6 decimal places on E4980A
        return f"{value:.6f}"

    # Q — use sig_figs
    if value == 0.0:
        return f"0.{'0' * (sig_figs - 1)}"
    mag = math.floor(math.log10(abs(value)))
    decimal_places = max(0, sig_figs - 1 - int(mag))
    factor = 10 ** (sig_figs - 1 - int(mag))
    rounded = _round_half_up(value * factor, 0) / factor
    return f"{rounded:.{decimal_places}f}"


# ============================================================
#  4. MAIN PUBLIC FUNCTION
# ============================================================
def format_lcr_value(value: float, param_type: str,
                     aperture: str) -> str:
    """
    Format a raw FETCh? float to match E4980A front-panel display.

    Parameters
    ----------
    value      : raw float from instrument  e.g. 1.000081e-10
    param_type : "C", "L", "R", "Z", "X", "RS", "RP",
                 "Y", "G", "B", "D", "Q", "THETA"
    aperture   : "SHORT", "MED", or "LONG"

    Returns
    -------
    str : e.g.  "100.0081 pF"  /  "15.9179 kΩ"  /  "-0.000015"
    """
    sig_figs = get_sig_figs(aperture)
    p = param_type.strip().upper()

    # ── Capacitance ──────────────────────────────────────────
    if p == "C":
        return _apply_prefix(value, sig_figs, [
            (1e-3,  "mF"),
            (1e-6,  "µF"),
            (1e-9,  "nF"),
            (1e-12, "pF"),
        ])

    # ── Inductance ───────────────────────────────────────────
    elif p == "L":
        return _apply_prefix(value, sig_figs, [
            (1.0,   "H"),
            (1e-3,  "mH"),
            (1e-6,  "µH"),
            (1e-9,  "nH"),
        ])

    # ── Resistance / Impedance Z / Reactance X / Rs / Rp ────
    elif p in ("R", "Z", "X", "RS", "RP"):
        return _apply_prefix(value, sig_figs, [
            (1e6,  "MΩ"),
            (1e3,  "kΩ"),
            (1.0,  "Ω"),
            (1e-3, "mΩ"),
            (1e-6, "µΩ"),
            (1e-9, "nΩ"),
        ])

    # ── Admittance Y / Conductance G / Susceptance B ────────
    elif p in ("Y", "G", "B"):
        return _apply_prefix(value, sig_figs, [
            (1.0,  "S"),
            (1e-3, "mS"),
            (1e-6, "µS"),
            (1e-9, "nS"),
        ])

    # ── Dissipation Factor D ─────────────────────────────────
    elif p == "D":
        return _format_dimensionless(value, sig_figs, "D")

    # ── Quality Factor Q ─────────────────────────────────────
    elif p == "Q":
        return _format_dimensionless(value, sig_figs, "Q")

    # ── Phase Angle Theta (degrees) ──────────────────────────
    elif p == "THETA":
        # E4980A always shows phase to 4 decimal places
        return f"{value:.4f} °"

    # ── Fallback ─────────────────────────────────────────────
    else:
        return _format_dimensionless(value, sig_figs)


# ============================================================
#  5. MODE → (PARAM1_TYPE, PARAM2_TYPE)
# ============================================================
MODE_PARAM_TYPES = {
    # Cp series
    "CPD":   ("C", "D"),
    "CPQ":   ("C", "Q"),
    "CPG":   ("C", "G"),
    "CPRP":  ("C", "R"),
    # Cs series
    "CSD":   ("C", "D"),
    "CSQ":   ("C", "Q"),
    "CSRS":  ("C", "RS"),
    # Lp series
    "LPD":   ("L", "D"),
    "LPQ":   ("L", "Q"),
    "LPG":   ("L", "G"),
    "LPRP":  ("L", "R"),
    # Ls series
    "LSD":   ("L", "D"),
    "LSQ":   ("L", "Q"),
    "LSRS":  ("L", "RS"),
    # Z-Theta
    "ZTD":   ("Z", "THETA"),
    "ZTR":   ("Z", "THETA"),
    "ZTHD":  ("Z", "THETA"),
    "ZTHR":  ("Z", "THETA"),
    # Y-Theta
    "YTD":   ("Y", "THETA"),
    "YTR":   ("Y", "THETA"),
    # R-X
    "RX":    ("R", "X"),
    # G-B
    "GB":    ("G", "B"),
}


# ============================================================
#  7. MODE → (P1_LABEL, P2_LABEL)  for graph titles / axis labels
#     These are the human-facing names (Cp, Cs, Lp, Ls, Rp, Rs, ...)
#     distinct from MODE_PARAM_TYPES, which only carries the
#     raw value TYPE ("C","L","R"...) used for unit formatting.
# ============================================================
MODE_LABELS = {
    "CPD":   ("Cp", "D"),
    "CPQ":   ("Cp", "Q"),
    "CPG":   ("Cp", "G"),
    "CPRP":  ("Cp", "Rp"),

    "CSD":   ("Cs", "D"),
    "CSQ":   ("Cs", "Q"),
    "CSRS":  ("Cs", "Rs"),

    "LPD":   ("Lp", "D"),
    "LPQ":   ("Lp", "Q"),
    "LPG":   ("Lp", "G"),
    "LPRP":  ("Lp", "Rp"),

    "LSD":   ("Ls", "D"),
    "LSQ":   ("Ls", "Q"),
    "LSRS":  ("Ls", "Rs"),

    "ZTD":   ("Z", "Theta"),
    "ZTR":   ("Z", "Theta"),
    "ZTHD":  ("Z", "Theta"),
    "ZTHR":  ("Z", "Theta"),

    "YTD":   ("Y", "Theta"),
    "YTR":   ("Y", "Theta"),

    "RX":    ("R", "X"),
    "GB":    ("G", "B"),
}


def get_mode_labels(mode: str):
    """Return (p1_label, p2_label) for a mode code, e.g. 'CSRS' -> ('Cs','Rs')."""
    return MODE_LABELS.get(mode.strip().upper(), ("P1", "P2"))


# ============================================================
#  8. SELF-TEST — python formatter.py
# ============================================================
if __name__ == "__main__":
    tests = [
        # ── VERIFIED from real meter photo ──
        (1.000081e-10, "C", "LONG",  "METER SHOWS: 100.0081 pF"),
        (-1.5e-5,      "D", "LONG",  "METER SHOWS: -0.000015"),

        # ── Capacitance across apertures ──
        (1.000081e-10, "C", "SHORT", "100.01 pF   (5 sig)"),
        (1.000081e-10, "C", "MED",   "100.008 pF  (6 sig)"),
        (3.3e-9,       "C", "LONG",  "3.300000 nF"),
        (4.7e-6,       "C", "LONG",  "4.700000 µF"),
        (820e-12,      "C", "LONG",  "820.0000 pF"),
        (1.5e-12,      "C", "MED",   "1.50000 pF"),

        # ── Inductance ──
        (1.5e-3,       "L", "MED",   "1.50000 mH"),
        (47e-6,        "L", "LONG",  "47.00000 µH"),
        (2.2,          "L", "MED",   "2.20000 H"),

        # ── Resistance / Z ──
        (470.0,        "R", "LONG",  "470.0000 Ω"),
        (15917.89,     "Z", "LONG",  "15.91789 kΩ"),
        (2.2e6,        "R", "MED",   "2.20000 MΩ"),

        # ── D values ──
        (-1.5e-5,      "D", "LONG",  "-0.000015000"),
        (0.002341,     "D", "SHORT", "0.0023410"),
        (0.05,         "D", "MED",   "0.050000"),

        # ── Q ──
        (120.5,        "Q", "MED",   "120.500"),
        (5500.0,       "Q", "MED",   "5500.00"),

        # ── Theta ──
        (-89.9937,     "THETA","LONG", "-89.9937 °"),

        # ── Y / G / B ──
        (0.00628,      "Y", "MED",   "6.28000 mS"),
    ]

    print(f"\n{'RAW VALUE':>22}  {'TYPE':5}  {'APT':5}  {'RESULT':<22}  NOTE")
    print("─" * 80)
    for val, ptype, apt, note in tests:
        result = format_lcr_value(val, ptype, apt)
        marker = "◀ VERIFIED" if "METER SHOWS" in note else ""
        print(f"{val:>22.6e}  {ptype:5}  {apt:5}  {result:<22}  {note} {marker}")
