from pathlib import Path
import re
import pandas as pd
import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter


NUM_BLACKBOXES = 8192

# SR0–SR59 (Athena) FF count per SR
FF_PER_SR = [
    4, 4, 1, 4, 1, 1, 2, 2, 2, 2, 1, 2, 2, 2, 1, 1,
    2, 6, 6, 6, 2, 4, 4, 4, 4, 4, 2, 2, 2, 2, 2, 4,
    6, 6, 1, 5, 5, 3, 4, 4, 4, 4, 1, 2, 3, 2, 2, 4,
    4, 1, 4, 4, 1, 2, 3, 3, 1, 1, 1, 2
]
NUM_SRS = len(FF_PER_SR)


FIT_SCALE = 1.0e15
FIT_FACTOR = 0.001


def safe_float(x, default=0.0) -> float:
    try:
        if x is None:
            return default
        if isinstance(x, str) and x.strip() in {"---", ""}:
            return default
        return float(x)
    except Exception:
        return default


def get_timestamp(s: str) -> dt.datetime:
    s = str(s).strip()
    for fmt in ("%m/%d/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%H:%M:%S", "%m/%d/%y %H:%M"):
        try:
            return dt.datetime.strptime(s, fmt)
        except Exception:
            pass
    return dt.datetime.strptime("1:1:1", "%H:%M:%S")


def compute_duration_seconds_from_seu_csv(csv_path: Path) -> float:
    # Compute irradiation duration from first and last timestamps in SEU CSV
    try:
        df = pd.read_csv(csv_path, usecols=[0])
        if df.empty:
            return 0.0
        start = get_timestamp(df.iloc[0, 0])
        end = get_timestamp(df.iloc[-1, 0])
        dur = (end - start).total_seconds()
        return float(dur) if dur > 0 else 0.0
    except Exception:
        # Fallback: manually read file line by line
        try:
            with open(csv_path, "r") as f:
                f.readline()
                first = None
                last = None
                for line in f:
                    parts = line.split(",")
                    if not parts:
                        continue
                    ts = parts[0].strip()
                    if not ts:
                        continue
                    if first is None:
                        first = ts
                    last = ts
            if first is None or last is None:
                return 0.0
            start = get_timestamp(first)
            end = get_timestamp(last)
            dur = (end - start).total_seconds()
            return float(dur) if dur > 0 else 0.0
        except Exception:
            return 0.0


def infer_fluence_from_particle_and_duration(particle: str, duration_s: float) -> float:
    # Estimate fluence based on particle type encoded in filename
    # Assumes fixed flux rates for known beams
    p = (particle or "").upper()
    if "100U" in p:
        return duration_s * 1_000_000.0
    if "10U" in p:
        return duration_s * 100_000.0
    return 0.0


def parse_seu_filename(stem: str) -> dict:
    # Extract metadata fields from SEU filename
    parts = stem.split("_")

    meta = {
        "angle": "0",
        "temperature": "---",
        "vdd_core": "---",
        "vdd_io": "1.2",
        "die": "---",
        "particle": "---",
        "run": "---",
        "code_used": "LSB",
        "clock_mode": "2.5",
        "frequency": "---",
        "input_bits": "---",
        "fluence": "---",
        "board": "---",
        "actual_freq": "0.0",
    }

    # Expected pattern: <run>_<...>_<vdd>_<input>_<particle>_<freq>_<temp>_<board>_SEU
    if len(parts) >= 9 and parts[-1].upper() == "SEU":
        meta["run"] = parts[0]
        meta["vdd_core"] = parts[2]
        meta["input_bits"] = parts[3]
        meta["particle"] = parts[4]
        meta["frequency"] = parts[5]
        meta["temperature"] = parts[6]
        meta["board"] = parts[7]
        meta["die"] = meta["board"]
        meta["actual_freq"] = parts[5]

    return meta


def extract_run_number_from_path(p: Path) -> int:
    # Extract numeric run number from filename for sorting
    m = re.match(r"^(\d+)_", p.stem)
    return int(m.group(1)) if m else 10**9


def build_run_id_from_seu_stem(stem: str) -> str:
    # Build run identifier matching RUN_LOG.csv format
    parts = stem.split("_")
    if len(parts) >= 9 and parts[-1].upper() == "SEU":
        return "_".join(parts[0:8])
    return ""


def load_fluence_map_from_run_log(run_log_path: Path) -> dict:
    # Load fluence values from RUN_LOG.csv
    # Maps run_id → fluence
    fluences = {}
    if not run_log_path.exists():
        return fluences

    with open(run_log_path, "r") as infile:
        infile.readline()
        for line in infile:
            temp1 = line.strip().split(",")
            if len(temp1) < 2:
                continue

            temp = temp1[1].split("_") if len(temp1) < 4 else temp1
            if len(temp) < 8:
                continue

            run_id = "_".join(temp[:8])
            try:
                fluences[run_id] = float(temp1[-1])
            except Exception:
                continue

    return fluences


def compute_errors_per_sr(csv_path: Path, sr_count: int = NUM_SRS) -> pd.Series:
    # Read SEU CSV and sum errors per shift register
    df = pd.read_csv(csv_path)

    cols = {c.strip(): c for c in df.columns}
    sr_col = None
    err_col = None

    # Detect SR column name
    for candidate in ["SR", "Sr", "Shift Register", "ShiftRegister"]:
        if candidate in cols:
            sr_col = cols[candidate]
            break

    # Detect error-count column name
    for candidate in ["Error Count", "ErrorCount", "Errors", "Error"]:
        if candidate in cols:
            err_col = cols[candidate]
            break

    if sr_col is None or err_col is None:
        raise ValueError(f"{csv_path.name}: missing SR / Error Count columns")

    df[sr_col] = pd.to_numeric(df[sr_col], errors="coerce")
    df[err_col] = pd.to_numeric(df[err_col], errors="coerce").fillna(0)
    df = df.dropna(subset=[sr_col])

    # Convert SR numbering to 0-based if needed
    sr_vals = df[sr_col].astype(int)
    if len(sr_vals) > 0 and sr_vals.min() == 1 and sr_vals.max() == sr_count:
        sr_vals -= 1

    df["_sr0"] = sr_vals
    summed = df.groupby("_sr0")[err_col].sum()

    # Ensure every SR index exists
    full = pd.Series(0, index=range(sr_count), dtype="int64")
    for k, v in summed.items():
        if 0 <= int(k) < sr_count:
            full[int(k)] = int(v)

    return full


def resolve_fluence(meta: dict, fpath: Path, fluence_map: dict) -> float:
    # Resolve fluence from RUN_LOG if possible, else infer from duration+particle
    stem = fpath.stem
    run_id = build_run_id_from_seu_stem(stem)
    if run_id in fluence_map:
        meta["fluence"] = str(fluence_map[run_id])

    fluence = safe_float(meta.get("fluence"), 0.0)
    if fluence <= 0.0:
        duration_s = compute_duration_seconds_from_seu_csv(fpath)
        fluence = infer_fluence_from_particle_and_duration(meta.get("particle", ""), duration_s)
        meta["fluence"] = str(fluence)

    return float(fluence)


def compute_cross_section(err_cnt: int, sr_index: int, fluence: float) -> float:
    """
    CrossSection/FF (cm^2) = errors / (fluence * total_FF)
    total_FF = NUM_BLACKBOXES * FF_PER_SR[sr_index]
    """
    total_ffs = NUM_BLACKBOXES * int(FF_PER_SR[sr_index])
    denom = float(fluence) * float(total_ffs)
    return (float(err_cnt) / denom) if denom > 0 else 0.0


def build_summary_excel(folder: str, out_xlsx: str, sr_count: int = NUM_SRS) -> None:
    # Main routine: parse all SEU CSVs and build Excel summary (red blocks)
    if sr_count != len(FF_PER_SR):
        raise ValueError("SR count mismatch with FF_PER_SR")

    folder_path = Path(folder)
    fluence_map = load_fluence_map_from_run_log(folder_path / "RUN_LOG.csv")

    # Collect and sort SEU CSV files
    seu_files = sorted(
        [p for p in folder_path.iterdir()
         if p.is_file() and p.name.upper().endswith("_SEU.CSV")],
        key=lambda p: (extract_run_number_from_path(p), p.name)
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    pink = PatternFill("solid", fgColor="E7C3C0")
    header_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    # Header metadata labels
    labels = [
        ("Angle", "Degrees", "angle"),
        ("Temperature", "C", "temperature"),
        ("Vdd_core", "V", "vdd_core"),
        ("Vdd_io", "V", "vdd_io"),
        ("Die #", "", "die"),
        ("Particle", "", "particle"),
        ("Run", "", "run"),
        ("Code Used", "", "code_used"),
        ("Clock Mode", "", "clock_mode"),
        ("Frequency", "", "frequency"),
        ("Input", "", "input_bits"),
        ("Fluence", "", "fluence"),
    ]

    # Write left-side metadata labels
    for r, (lab, unit, _) in enumerate(labels, start=1):
        ws.cell(row=r, column=1, value=lab).alignment = left
        ws.cell(row=r, column=2, value=unit).alignment = left
        ws.cell(row=r, column=1).font = header_font

    table_start_row = 15

    # Write SR index and hardware constants
    for i in range(sr_count):
        ws.cell(row=table_start_row + 1 + i, column=1, value=f"SR-{i}")
        ws.cell(row=table_start_row + 1 + i, column=2, value=NUM_BLACKBOXES).alignment = center
        ws.cell(row=table_start_row + 1 + i, column=3, value=FF_PER_SR[i]).alignment = center

    ws.cell(row=table_start_row, column=2, value="# of BlackBoxes").font = header_font
    ws.cell(row=table_start_row, column=3, value="# of FF/BB").font = header_font

    block_start_col = 5
    block_width = 6
    gap = 2

    # Process each SEU run and create a results block
    for idx, fpath in enumerate(seu_files):
        meta = parse_seu_filename(fpath.stem)
        fluence = resolve_fluence(meta, fpath, fluence_map)

        errors = compute_errors_per_sr(fpath, sr_count)
        start_col = block_start_col + idx * (block_width + gap)

        # Column formatting
        ws.column_dimensions[get_column_letter(start_col)].width = 14
        ws.column_dimensions[get_column_letter(start_col + 1)].width = 28
        ws.column_dimensions[get_column_letter(start_col + 2)].width = 14

        # Background shading for block
        for r in list(range(1, len(labels) + 1)) + list(range(table_start_row, table_start_row + 1 + sr_count)):
            for c in range(start_col, start_col + block_width):
                ws.cell(row=r, column=c).fill = pink

        # Write metadata values
        for r, (_, _, key) in enumerate(labels, start=1):
            ws.cell(row=r, column=start_col, value=meta.get(key, "---")).alignment = center
            ws.cell(row=r, column=start_col).font = header_font

        # Table headers
        ws.cell(row=table_start_row, column=start_col, value="# of Errors").font = header_font
        ws.cell(row=table_start_row, column=start_col + 1, value="CrossSection/FF(cm^2)").font = header_font
        ws.cell(row=table_start_row, column=start_col + 2, value="FIT").font = header_font

        # Per-SR calculations
        for i in range(sr_count):
            err_i = int(errors.iloc[i])
            ws.cell(row=table_start_row + 1 + i, column=start_col, value=err_i).alignment = center

            xs_per_ff = compute_cross_section(err_i, i, fluence)
            ws.cell(row=table_start_row + 1 + i, column=start_col + 1, value=xs_per_ff).alignment = center

            # FIT = cross section per FF × 1e15 × scaling factor
            fit = xs_per_ff * FIT_SCALE * FIT_FACTOR
            ws.cell(row=table_start_row + 1 + i, column=start_col + 2, value=fit).alignment = center

    wb.save(out_xlsx)


def build_long_format_excel(folder: str, out_long_xlsx: str, sr_count: int = NUM_SRS) -> None:
    """
    Long format: one row per (run, SR) with cs == CrossSection/FF from the summary blocks.
    """
    folder_path = Path(folder)
    fluence_map = load_fluence_map_from_run_log(folder_path / "RUN_LOG.csv")

    seu_files = sorted(
        [p for p in folder_path.iterdir()
         if p.is_file() and p.name.upper().endswith("_SEU.CSV")],
        key=lambda p: (extract_run_number_from_path(p), p.name)
    )

    rows = []
    for fpath in seu_files:
        meta = parse_seu_filename(fpath.stem)
        fluence = resolve_fluence(meta, fpath, fluence_map)
        errors = compute_errors_per_sr(fpath, sr_count)

        # Stable "id" per run (adjust if you want a different convention)
        parts = fpath.stem.split("_")
        run_id_like = "_".join(parts[:4]) if len(parts) >= 4 else fpath.stem

        vdd = safe_float(meta.get("vdd_core"), 0.0)
        freq = safe_float(meta.get("frequency"), 0.0)
        actual_freq = safe_float(meta.get("actual_freq"), freq)

        for sr in range(sr_count):
            err_cnt = int(errors.iloc[sr])
            ff_per_bb = int(FF_PER_SR[sr])

            cs = compute_cross_section(err_cnt, sr, fluence)
            fit = cs * FIT_SCALE * FIT_FACTOR

            rows.append({
                "id": run_id_like,
                "run": meta.get("run", "---"),
                "vdd": vdd,
                "input": meta.get("input_bits", "---"),
                "ion": meta.get("particle", "---"),
                "temp": meta.get("temperature", "---"),
                "freq": freq,
                "brd": meta.get("board", "---"),
                "actual_freq": actual_freq,
                "fluence": float(fluence),
                "die": meta.get("die", "---"),
                "# of blackboxes": NUM_BLACKBOXES,
                "# of FF/BB": ff_per_bb,
                "SR_NUM": sr,
                "err_cnt": err_cnt,
                "cs": cs,
                "upper": 0.0,
                "lower": 0.0,
                "FIT": fit,
                "source_file": fpath.name,
            })

    df = pd.DataFrame(rows)

    col_order = [
        "id", "run", "vdd", "input", "ion", "temp", "freq", "brd", "actual_freq",
        "fluence", "die", "# of blackboxes", "# of FF/BB", "SR_NUM", "err_cnt",
        "cs", "upper", "lower", "FIT", "source_file"
    ]
    df = df[[c for c in col_order if c in df.columns]]

    with pd.ExcelWriter(out_long_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="LongFormat", index=False)

        # Basic column width cleanup
        ws = writer.book["LongFormat"]
        for j, col in enumerate(df.columns, start=1):
            width = max(10, min(40, len(col) + 2))
            ws.column_dimensions[get_column_letter(j)].width = width


if __name__ == "__main__":
    # Input folder containing SEU CSVs and RUN_LOG.csv
    folder = "/Users/jennakronenberg/Desktop/N3hf/"

    # Output Excel summary paths
    out_summary_xlsx = "/Users/jennakronenberg/Desktop/N3hf/summary.xlsx"
    out_long_xlsx = "/Users/jennakronenberg/Desktop/N3hf/summary_long.xlsx"

    build_summary_excel(folder, out_summary_xlsx)
    build_long_format_excel(folder, out_long_xlsx)

    print(f"Saved summary: {out_summary_xlsx}")
    print(f"Saved long format: {out_long_xlsx}")
