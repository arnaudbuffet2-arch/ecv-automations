import csv
import sys
from pathlib import Path


def fix_csv(input_path: str):
    src = Path(input_path)
    if not src.exists():
        print(f"Fichier introuvable : {input_path}")
        sys.exit(1)

    dst = src.with_stem(src.stem + "_propre")

    with open(src, newline="", encoding="utf-8-sig") as f_in:
        raw_rows = list(csv.reader(f_in))

    if not raw_rows:
        print("Fichier vide.")
        sys.exit(1)

    with open(dst, "w", newline="", encoding="utf-8-sig") as f_out:
        writer = csv.writer(f_out, delimiter=";")
        for row in raw_rows:
            writer.writerow(row)

    col_count = len(raw_rows[0][0].split(",")) if len(raw_rows[0]) == 1 else len(raw_rows[0])
    print(f"Fichier créé : {dst}")
    print(f"Colonnes détectées : {col_count}")
    print(f"Lignes traitées : {len(raw_rows)}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('Usage : python scripts/fix_csv.py "C:\\chemin\\vers\\fichier.csv"')
        sys.exit(1)
    fix_csv(sys.argv[1])
