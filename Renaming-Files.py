import pandas as pd
from pathlib import Path
import shutil
import sys
import re

EXCEL_FILE = r"File Path"   # <— change this
SHEET_NAME = 0
COL_PATH = "Path_Files"
COL_TARGET = "Names_to_be"
DRY_RUN = False
MAKE_BACKUP_CSV = True


TARGET_FOLDER = None

_ILLEGAL = r'<>:"/\\|?*'
_illegal_re = re.compile(f"[{re.escape(_ILLEGAL)}]")

def sanitize_filename(name: str) -> str:
    name = name.strip()
    name = _illegal_re.sub("_", name)
    return name.rstrip(" .")

def ensure_extension(target_stem: str, original_path: Path) -> str:
    tgt = Path(target_stem)
    if tgt.suffix:  # user provided extension
        return str(tgt.name)
    else:
        return tgt.name + original_path.suffix

def unique_path(dest_path: Path) -> Path:
    if not dest_path.exists():
        return dest_path
    base = dest_path.stem
    ext = dest_path.suffix
    parent = dest_path.parent
    i = 1
    while True:
        candidate = parent / f"{base} ({i}){ext}"
        if not candidate.exists():
            return candidate
        i += 1

def main():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"ERROR: Could not read Excel: {e}")
        sys.exit(1)

    for col in (COL_PATH, COL_TARGET):
        if col not in df.columns:
            print(f"ERROR: Column '{col}' not found in Excel.")
            sys.exit(1)

    actions = []  # for optional logging

    for idx, row in df.iterrows():
        raw_path = row[COL_PATH]
        raw_newname = row[COL_TARGET]

        if pd.isna(raw_path) or pd.isna(raw_newname):
            print(f"[SKIP] Row {idx}: missing Path or Names_to_be.")
            continue

        src = Path(str(raw_path)).expanduser()
        if not src.exists():
            print(f"[MISS] Row {idx}: File not found -> {src}")
            continue
        if not src.is_file():
            print(f"[SKIP] Row {idx}: Path is not a file -> {src}")
            continue

        target_stem = sanitize_filename(str(raw_newname))
        final_name = ensure_extension(target_stem, src)

        dest_dir = Path(TARGET_FOLDER).expanduser() if TARGET_FOLDER else src.parent
        dest_dir.mkdir(parents=True, exist_ok=True)

        dest = dest_dir / final_name
        dest = unique_path(dest)  # avoid overwrite

        if DRY_RUN:
            print(f"[DRY] {src}  ->  {dest.name}")
        else:
            try:
                # same-folder rename == move with same parent
                shutil.move(str(src), str(dest))
                print(f"[OK ] {src.name}  ->  {dest.name}")
                actions.append({"old_path": str(src), "new_path": str(dest)})
            except Exception as e:
                print(f"[ERR] Could not rename '{src}' -> '{dest}': {e}")

    if (not DRY_RUN) and MAKE_BACKUP_CSV and actions:
        log_path = Path(EXCEL_FILE).with_name("renamed_log.csv")
        pd.DataFrame(actions).to_csv(log_path, index=False)
        print(f"Log written to: {log_path}")

if __name__ == "__main__":
    main()
