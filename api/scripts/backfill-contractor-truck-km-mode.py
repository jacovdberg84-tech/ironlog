"""Set utilization_mode=km on contractor / truck assets that should use km/L benchmarking."""
import os
import sqlite3

DB = os.path.join(os.environ["APPDATA"], "ironlog-api", "data", "db", "ironlog.db")
if not os.path.exists(DB):
    DB = os.path.join(os.path.dirname(__file__), "..", "db", "ironlog.db")

KM_HINTS = ("truck", "vehicle", "ldv", "pickup", "bakkie", "tipper", "dump", "haul", "spinner")


def should_be_km(code, name, category):
    code_u = (code or "").upper()
    name_l = (name or "").lower()
    cat_l = (category or "").lower()
    if code_u.startswith("PTT") or code_u.startswith("LDV"):
        return True
    if len(code_u) == 5 and code_u[0] in "VT" and code_u.endswith("AM") and code_u[1:3].isdigit():
        return True
    if any(h in cat_l or h in name_l for h in KM_HINTS):
        return True
    if "toyota" in name_l and "hilux" in name_l:
        return True
    return False


def main():
    print("DB:", DB)
    c = sqlite3.connect(DB)
    rows = c.execute(
        "SELECT id, asset_code, asset_name, category, utilization_mode FROM assets WHERE COALESCE(archived,0)=0"
    ).fetchall()
    updated = 0
    for asset_id, code, name, category, mode in rows:
        if str(mode or "").strip().lower() == "km":
            continue
        if not should_be_km(code, name, category):
            continue
        c.execute("UPDATE assets SET utilization_mode = 'km' WHERE id = ?", (asset_id,))
        print(f"  km mode: {code} — {name}")
        updated += 1
    c.commit()
    c.close()
    print(f"Updated {updated} asset(s).")


if __name__ == "__main__":
    main()
