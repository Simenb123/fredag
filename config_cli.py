from __future__ import annotations
import argparse
from .config_io import export_config, import_config

def main():
    ap = argparse.ArgumentParser(prog="fredag.config", description="Eksporter/Importer Fredag-konfig (ZIP).")
    sp = ap.add_subparsers(dest="cmd", required=True)

    p_exp = sp.add_parser("export", help="Eksporter konfig til ZIP")
    p_exp.add_argument("zip_path", help="Sti til ZIP som skal skrives")

    p_imp = sp.add_parser("import", help="Importer konfig fra ZIP")
    p_imp.add_argument("zip_path", help="Sti til ZIP som skal leses")
    p_imp.add_argument("--no-backup", action="store_true", help="Ikke ta backup av eksisterende filer")

    args = ap.parse_args()
    if args.cmd == "export":
        ok, msg = export_config(args.zip_path)
        print(("OK: " if ok else "FEIL: ") + msg)
    else:
        ok, msg = import_config(args.zip_path, backup_current=not args.no_backup)
        print(("OK: " if ok else "FEIL: ") + msg)

if __name__ == "__main__":
    main()
