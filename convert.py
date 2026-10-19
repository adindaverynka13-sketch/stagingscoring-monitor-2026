"""Ubah file Excel menjadi data.json (tanpa nama/email/NIP) untuk dashboard.
Pemakaian:  python convert.py "Data_Leader_Score_2026.xlsx"
Butuh: pip install pandas openpyxl
Catatan: nilai kolom 'Score Total' dibaca dari hasil hitung Excel, jadi file
harus pernah dibuka & disimpan di Excel agar nilai formulanya tersimpan."""
import sys, json, pandas as pd

src = sys.argv[1] if len(sys.argv) > 1 else "Data_Leader_Score_2026.xlsx"
df = pd.read_excel(src, engine="openpyxl")

rows = []
for _, r in df.iterrows():
    s = r["Score Total"]
    region = (str(r["Departemen"]).replace("REGION ", "").replace(" OPERATION", "")
              .strip().title().replace("Diy","DIY"))
    rows.append({
        "region": region,
        "kategori": str(r["Kategori Staging"]).strip().title(),
        "score": None if pd.isna(s) else round(float(s) * 100, 2),
    })

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False)
print(f"OK: {len(rows)} baris -> data.json")
