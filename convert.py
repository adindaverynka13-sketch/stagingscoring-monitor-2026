"""Ubah file Excel menjadi:
  1) data.js                      -> untuk dashboard PUBLIK (tanpa nama/NIP), boleh di-upload ke GitHub
  2) individu-lokal/data-individu.js -> untuk laporan individu (ADA nama & NIP), JANGAN di-upload ke repo Public
Pemakaian:  python convert.py "Data_Leader_Score_2026.xlsx"
Butuh: pip install pandas openpyxl. File Excel harus pernah disimpan di Excel agar nilai formulanya tersimpan."""
import sys, os, json, pandas as pd

src = sys.argv[1] if len(sys.argv) > 1 else "Data_Leader_Score_2026.xlsx"
df = pd.read_excel(src, engine="openpyxl")

def num(x, nd=2): return None if pd.isna(x) else round(float(x), nd)
def txt(x): return "" if pd.isna(x) else str(x).strip()
def short(d): return txt(d).replace("REGION ", "").replace(" OPERATION", "").title().replace("Diy", "DIY")

pub, priv = [], []
for _, r in df.iterrows():
    s = num(r["Score Total"], 4)
    base = {"region": short(r["Departemen"]), "kategori": txt(r["Kategori Staging"]).title()}
    comp = {"fm": num(r["First Mile Score"]), "lm": num(r["Last Mile Score"]), "ph": num(r["Product Handling Score"])}
    score = None if s is None else round(s * 100, 2)
    pub.append({**base, "score": score, "ipk": num(r["IPK Total"]), **comp})
    priv.append({**base, "nip": txt(r["NIP"]), "nama": txt(r["Nama Pegawai"]), "jabatan": txt(r["Jabatan Terakhir"]).title(),
                 "kode": txt(r["Kode Lokasi"]), "lokasi": txt(r["Lokasi"]), "batch": txt(r.get("Batch")).title(), "pic": txt(r.get("PIC")).title(),
                 "score": score, "ipk": num(r["IPK Total"]), **comp,
                 "lf": num(r["IPK First Mile"], 0), "ll": num(r["IPK Last Mile"], 0), "lp": num(r["IPK Product Handling"], 0)})

with open("data.js", "w", encoding="utf-8") as f:
    f.write("const DATA = " + json.dumps(pub, ensure_ascii=False) + ";")
os.makedirs("individu-lokal", exist_ok=True)
with open("individu-lokal/data-individu.js", "w", encoding="utf-8") as f:
    f.write("const DATA = " + json.dumps(priv, ensure_ascii=False) + ";")
print(f"OK: {len(pub)} baris -> data.js dan individu-lokal/data-individu.js")
