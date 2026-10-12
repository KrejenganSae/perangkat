#!/usr/bin/env python3
"""Validasi data/perangkat.xlsx, cek foto, olah data, lalu tulis dist/ untuk GitHub Pages.

Yang masuk ke dist/: index.html, data.json, foto/ (sudah diperkecil).
Yang TIDAK masuk: data/perangkat.xlsx, tanggal lahir, telepon.

Kode keluar 1 bila ada ERROR (deploy dibatalkan). WARNING hanya dilaporkan.
Jalankan lokal:  python build.py
"""
import collections
import html
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

try:
    from PIL import Image, ImageOps
except ImportError:  # foto disalin apa adanya bila Pillow tidak ada
    Image = None

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
XLSX = DATA / "perangkat.xlsx"
FOTO = DATA / "foto"
DIST = ROOT / "dist"
DEFAULT = {
    "instansi": "Pemerintah Kabupaten Probolinggo", "kecamatan": "Krejengan",
    "ignore_sheets": ["CATATAN", "Siltap"],
    "usia_pensiun": 60, "bulan_peringatan": 12, "usia_min": 17, "usia_maks": 80, "lebar_foto": 480,
    "siltap": {"kades": 2426640, "sekdes": 2224420, "perangkat": 2022200, "bulan": 12, "bulan_taliasih": 2},
}
CFG = {**DEFAULT, **json.loads((ROOT / "config.json").read_text(encoding="utf-8"))}  # kunci yang hilang memakai bawaan
CFG["siltap"] = {**DEFAULT["siltap"], **CFG["siltap"]}
CFG["ignore_sheets"] = sorted(set(CFG["ignore_sheets"]) | {"CATATAN", "Siltap"})

KOLOM = [
    "NO", "NAMA", "TANGGAL LAHIR", "JENIS KELAMIN", "JABATAN",
    "SK PENGANGKATAN", "PINDAIAN SK PENGANGKATAN", "TMT PENGANGKATAN",
    "SK JABATAN TERAKHIR", "PINDAIAN SK JABATAN TERAKHIR",
    "PENDIDIKAN", "AGAMA", "STATUS", "KINERJA/PENGHARGAAN", "RIWAYAT PEMBINAAN",
    "TELEPON", "PAS FOTO",
]
EKSTENSI_FOTO = {".jpg", ".jpeg", ".png", ".webp"}
WIB = timezone(timedelta(hours=7))
HARI_INI = datetime.now(WIB).date()

temuan = []  # (level, lokasi, pesan)


def error(lokasi, pesan):
    temuan.append(("ERROR", lokasi, pesan))


def warning(lokasi, pesan):
    temuan.append(("WARNING", lokasi, pesan))


def info(lokasi, pesan):
    temuan.append(("INFO", lokasi, pesan))


# ---------- util ----------
def slug(teks):
    t = unicodedata.normalize("NFKD", str(teks)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-") or "x"


def kunci(teks):
    """Kunci pencocokan nama/folder: huruf kecil, tanpa spasi dan tanda baca."""
    return re.sub(r"[^a-z0-9]", "", slug(teks))


def norm_kolom(t):
    """Samakan judul kolom: huruf besar, spasi rapat, tanpa spasi di sekitar '/'."""
    return re.sub(r"\s*/\s*", "/", re.sub(r"\s+", " ", str(t).upper().strip())) if t is not None else ""


def daftar(v):
    """Isian banyak dalam satu sel: dipisah baris baru (Alt+Enter) atau titik koma."""
    if v is None:
        return []
    return [t for t in (re.sub(r"\s+", " ", x).strip() for x in re.split(r"[\n;]+", str(v))) if t]


def status_scan(v, lok, nama, kolom):
    """'ada' -> True, 'tidak ada'/'belum' -> False, kosong -> None."""
    t = re.sub(r"\s+", " ", str(v)).strip().lower() if v is not None else ""
    if not t:
        return None
    if t == "ada":
        return True
    if t.startswith("tidak") or t.startswith("belum"):
        return False
    warning(lok, f"{nama}, kolom {kolom}: '{v}' tidak dikenal (isi 'ada' atau 'tidak ada')")
    return None


def bersih(v):
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def baca_tanggal(v):
    """Terima date/datetime, YYYY-MM-DD, atau DD-MM-YYYY / DD/MM/YYYY."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", s)
        if not m:
            raise ValueError(f"format tanggal tidak dikenal: {s!r}")
        d, mo, y = map(int, m.groups())
    return date(y, mo, d)  # ValueError bila tanggal mustahil


def hitung_usia(lahir):
    return HARI_INI.year - lahir.year - ((HARI_INI.month, HARI_INI.day) < (lahir.month, lahir.day))


def tanggal_pensiun(lahir):
    """Pensiun tepat di awal bulan berikutnya setelah berusia usia_pensiun tahun."""
    y, m = lahir.year + CFG["usia_pensiun"], lahir.month + 1
    if m > 12:
        y, m = y + 1, 1
    return date(y, m, 1)


def status_pensiun(tgl):
    """'lewat' bila sudah melewati tanggal pensiun, 'mendekati' bila dalam jendela peringatan."""
    if tgl <= HARI_INI:
        return "lewat"
    bulan = (tgl.year - HARI_INI.year) * 12 + tgl.month - HARI_INI.month
    return "mendekati" if bulan <= CFG["bulan_peringatan"] else None


def pendidikan_baku(v):
    s = bersih(v)
    if not s:
        return None
    u = s.upper()
    if re.match(r"S[\s.\-]?3", u): return "S3"
    if re.match(r"S[\s.\-]?2", u): return "S2"
    if re.match(r"S[\s.\-]?1", u): return "S1"
    if re.match(r"D[\s.\-]?(IV|4)", u): return "D4"
    if re.match(r"D[\s.\-]?(III|3)", u): return "D3"
    if re.match(r"D[\s.\-]?(II|2)", u): return "D2"
    if re.match(r"D[\s.\-]?(I|1)\b", u): return "D1"
    if re.match(r"(SLTA|SMA|SMK|MA\b)", u): return "SLTA"
    if re.match(r"(SLTP|SMP|MTS)", u): return "SLTP"
    if re.match(r"(SD|MI\b)", u): return "SD"
    if "PAKET C" in u: return "Paket C"
    if "PAKET B" in u: return "Paket B"
    return s


def kelompok(jabatan):
    """Kelompok struktur. Jabatan umum/tak dikenal (mis. 'PERANGKAT DESA') masuk 'Perangkat Desa' sampai dirinci."""
    j = (jabatan or "").upper().strip()
    if re.search(r"\b(KEPALA DESA|KADES)\b", j) and "DUSUN" not in j: return "Kepala Desa"
    if "SEKRETARIS" in j or j.startswith("SEKDES"): return "Sekretaris Desa"
    if "DUSUN" in j: return "Kepala Dusun"
    if re.match(r"(KASI|KEPALA SEKSI)\b", j): return "Kasi"
    if re.match(r"(KAUR|KEPALA URUSAN)\b", j): return "Kaur"
    if j.startswith("STAF"): return "Staf"
    return "Perangkat Desa"


# ---------- foto ----------
def indeks_foto():
    """{kunci_desa: {kunci_nama: Path}}; juga catat folder/berkas yang tak terpakai."""
    peta = {}
    if not FOTO.is_dir():
        return peta
    for folder in sorted(p for p in FOTO.iterdir() if p.is_dir()):
        d = peta.setdefault(kunci(folder.name), {"_folder": folder.name})
        for f in sorted(folder.iterdir()):
            if f.suffix.lower() in EKSTENSI_FOTO:
                k = kunci(f.stem)
                if k in d:
                    warning(f"data/foto/{folder.name}", f"Dua foto untuk nama yang sama: {f.name}")
                d[k] = f
            elif f.name != ".gitkeep":
                warning(f"data/foto/{folder.name}", f"Berkas bukan foto diabaikan: {f.name}")
    return peta


def tulis_foto(sumber, tujuan):
    tujuan.parent.mkdir(parents=True, exist_ok=True)
    if Image is None:
        shutil.copyfile(sumber, tujuan)
        return
    with Image.open(sumber) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((CFG["lebar_foto"], CFG["lebar_foto"] * 4 // 3))
        im.save(tujuan, "JPEG", quality=82, optimize=True)


# ---------- baca dan validasi ----------
def proses():
    if not XLSX.exists():
        error("data/perangkat.xlsx", "Berkas tidak ditemukan di folder data")
        return None
    wb = load_workbook(XLSX, read_only=True, data_only=True)
    pagu = baca_pagu(wb)
    pagu_terpakai = set()
    foto = indeks_foto()
    foto_terpakai = set()
    hasil = []
    DIST_FOTO = DIST / "foto"

    for ws in wb.worksheets:
        desa = ws.title.strip()
        if desa in CFG["ignore_sheets"]:
            continue
        baris = ws.iter_rows(values_only=True)
        header = [norm_kolom(c) for c in next(baris, [])]

        def cari(k):  # judul boleh diberi keterangan dalam kurung, mis. "SK JABATAN TERAKHIR (MUTASI/PENGANGKATAN)"
            for i, h in enumerate(header):
                if h == k or h.startswith(k + " ("):
                    return i
            return None

        idx = {k: cari(k) for k in KOLOM}
        # judul polos "PINDAIAN SK" di sebelah kanan kolom SK-nya juga diterima
        for k, dasar in (("PINDAIAN SK PENGANGKATAN", "SK PENGANGKATAN"), ("PINDAIAN SK JABATAN TERAKHIR", "SK JABATAN TERAKHIR")):
            if idx[k] is None and idx[dasar] is not None:
                j = idx[dasar] + 1
                if j < len(header) and header[j] == "PINDAIAN SK":
                    idx[k] = j
        hilang = [k for k, i in idx.items() if i is None]
        if hilang:
            error(desa, "Kolom hilang atau salah ketik: " + ", ".join(hilang))
            continue

        orang, nama_terlihat = [], {}
        dk = kunci(desa)
        foto_desa = foto.get(dk, {})
        slug_desa = slug(desa)
        if slug_desa in RESERVED:
            error(desa, f"Nama desa menghasilkan alamat '{slug_desa}' yang bentrok dengan berkas situs")

        for n, r in enumerate(baris, start=2):
            if not any(c not in (None, "") for c in r):
                continue
            g = lambda k: r[idx[k]] if idx[k] < len(r) else None
            nama, jabatan = bersih(g("NAMA")), bersih(g("JABATAN"))
            lok = f"{desa} baris {n}"

            if not nama:
                if jabatan:
                    orang.append({
                        "nama": None, "jabatan": jabatan, "kelompok": kelompok(jabatan),
                        "lowong": True, "foto": None,
                    })
                    info(lok, f"{jabatan}: nama kosong, ditampilkan sebagai lowong")
                continue
            if not jabatan:  # data belum lengkap: jangan menghentikan build
                warning(lok, f"{nama}: jabatan kosong, ditampilkan sebagai PERANGKAT DESA sampai diisi")
                jabatan = "PERANGKAT DESA"

            if nama.upper() in nama_terlihat:
                warning(lok, f"Nama ganda di desa yang sama ({nama}); foto bisa tertukar")
            nama_terlihat[nama.upper()] = n

            # tanggal
            tgl = {}
            for k in ("TANGGAL LAHIR", "TMT PENGANGKATAN"):
                try:
                    tgl[k] = baca_tanggal(g(k))
                except ValueError as e:
                    error(lok, f"{nama}, kolom {k}: {e}")
                    tgl[k] = None
            lahir = tgl["TANGGAL LAHIR"]
            usia = None
            if lahir is None:
                if not bersih(g("TANGGAL LAHIR")):
                    warning(lok, f"{nama}: tanggal lahir kosong, usia tidak bisa dihitung")
            else:
                usia = hitung_usia(lahir)
                if not CFG["usia_min"] <= usia <= CFG["usia_maks"]:
                    error(lok, f"{nama}: usia {usia} tahun tidak wajar (lahir {lahir})")

            jk = bersih(g("JENIS KELAMIN"))
            jk = jk.upper() if jk else None
            if jk not in ("LAKI-LAKI", "PEREMPUAN"):
                warning(lok, f"{nama}: jenis kelamin kosong/tidak dikenal ({jk})")

            # foto vs Excel
            status_foto = bersih(g("PAS FOTO"))
            f_asal = foto_desa.get(kunci(nama))
            path_foto = None
            if f_asal:
                foto_terpakai.add((dk, kunci(nama)))
                nama_file = f"{slug(nama)}.jpg"
                tulis_foto(f_asal, DIST_FOTO / slug_desa / nama_file)
                path_foto = f"foto/{slug_desa}/{nama_file}"
                if status_foto and "belum" in status_foto.lower():
                    info(lok, f"{nama}: foto ada, tetapi kolom PAS FOTO masih '{status_foto}'")
            elif status_foto and "belum" not in status_foto.lower():
                warning(lok, f"{nama}: kolom PAS FOTO '{status_foto}', tetapi berkas foto tidak ada")

            kinerja = daftar(g("KINERJA/PENGHARGAAN"))
            pembinaan = daftar(g("RIWAYAT PEMBINAAN"))
            pensiun = sp = None
            if lahir is not None and kelompok(jabatan) != "Kepala Desa":  # aturan usia berlaku bagi perangkat, bukan kades
                sp = status_pensiun(tanggal_pensiun(lahir))
                if sp:  # tanggal hanya diterbitkan bila perlu lencana
                    pensiun = tanggal_pensiun(lahir).isoformat()
            if sp == "lewat" and int(pensiun[:4]) < HARI_INI.year:
                info(lok, f"{nama}: pensiun sejak {pensiun}, tidak dihitung lagi. Biarkan barisnya sampai akhir periode agar tercatat")
            ultah = lahir is not None and lahir.month == HARI_INI.month and sp != "lewat"
            telp = g("TELEPON")
            if isinstance(telp, (int, float)):
                telp = str(int(telp))
            telp = bersih(telp)
            if telp and re.fullmatch(r"8\d{7,12}", telp):
                telp = "0" + telp  # Excel membuang angka 0 di depan
            orang.append({
                "nama": nama,
                "jenis_kelamin": jk,
                "jabatan": jabatan,
                "kelompok": kelompok(jabatan),
                "usia": usia,
                "pendidikan": pendidikan_baku(g("PENDIDIKAN")),
                "agama": bersih(g("AGAMA")),
                "status": bersih(g("STATUS")),
                "sk_pengangkatan": bersih(g("SK PENGANGKATAN")),
                "scan_pengangkatan": status_scan(g("PINDAIAN SK PENGANGKATAN"), lok, nama, "PINDAIAN SK PENGANGKATAN"),
                "tmt": tgl["TMT PENGANGKATAN"].isoformat() if tgl["TMT PENGANGKATAN"] else None,
                "sk_terakhir": bersih(g("SK JABATAN TERAKHIR")),
                "scan_terakhir": status_scan(g("PINDAIAN SK JABATAN TERAKHIR"), lok, nama, "PINDAIAN SK JABATAN TERAKHIR"),
                "pensiun": pensiun,
                "pensiun_status": sp,
                "ultah_hari": lahir.day if ultah else None,
                "ultah_usia": HARI_INI.year - lahir.year if ultah else None,
                "telepon": telp,
                "kinerja": kinerja or None,          # daftar isian; rinciannya ditampilkan
                "pembinaan": len(pembinaan) or None,  # hanya jumlah; rincian tidak diterbitkan
                "foto": path_foto,
                "lowong": False,
            })

        # foto yatim di desa ini
        for k, f in foto_desa.items():
            if k != "_folder" and (dk, k) not in foto_terpakai:
                warning(f"data/foto/{foto_desa['_folder']}", f"{f.name} tidak cocok dengan nama mana pun di Excel")

        orang.extend(lengkapi_struktur(orang, desa))
        cnt = collections.Counter(p["kelompok"] for p in orang)
        if cnt["Perangkat Desa"]:
            info(desa, f"{cnt['Perangkat Desa']} orang berjabatan umum (mis. 'PERANGKAT DESA'), ditempatkan di kelompok Perangkat Desa sampai jabatannya dirinci")
        if cnt["Perangkat Desa"] < 3 and cnt["Sekretaris Desa"] < 1:
            info(desa, "Sekretaris Desa belum ada di Excel")
        p_desa = pagu.get(dk)
        if p_desa:
            pagu_terpakai.add(dk)
        elif pagu:
            warning(desa, "Pagu siltap desa ini tidak ada di sheet Siltap")
        hasil.append({"nama": desa, "slug": slug_desa, "perangkat": orang,
                      "siltap": hitung_siltap(orang, p_desa[1] if p_desa else None)})

    for k, (nm, _) in pagu.items():
        if k not in pagu_terpakai:
            warning("Siltap", f"Baris '{nm}' tidak cocok dengan nama sheet desa mana pun")
    # folder foto tanpa sheet
    sheet_kunci = {kunci(d["nama"]) for d in hasil}
    for k, d in foto.items():
        if k not in sheet_kunci:
            warning(f"data/foto/{d['_folder']}", "Folder tidak cocok dengan nama sheet/desa mana pun")
    return hasil


def tanggal_data():
    """Tanggal commit terakhir yang mengubah data/perangkat.xlsx (WIB); cadangan: waktu ubah berkas."""
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%cI", "--", "data/perangkat.xlsx"],
            cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip()
        if out:
            return datetime.fromisoformat(out).astimezone(WIB).date()
    except Exception:
        pass
    return datetime.fromtimestamp(XLSX.stat().st_mtime, WIB).date()


def ringkas(desa):
    """Yang sudah melewati tanggal pensiun tidak dihitung. Kades dan staf dipisahkan dari perangkat."""
    semua = [p for d in desa for p in d["perangkat"]]
    aktif = [p for p in semua if not p["lowong"] and p.get("pensiun_status") != "lewat"]
    kades = [p for p in aktif if p["kelompok"] == "Kepala Desa"]
    staf = [p for p in aktif if p["kelompok"] == "Staf"]
    return {
        "desa": len(desa),
        "kades": len(kades),
        "perangkat": len(aktif) - len(kades) - len(staf),
        "staf": len(staf),
        "lowong": sum(1 for p in semua if p["lowong"]),
        "pensiun_tahun_ini": sum(1 for p in semua if (p.get("pensiun") or "").startswith(str(HARI_INI.year))),
        "foto": sum(1 for p in aktif if p["foto"]),
    }


SLOT_KAUR = [("KAUR UMUM", r"UMUM|TATA USAHA"), ("KAUR KEUANGAN", r"KEUANGAN"), ("KAUR PERENCANAAN", r"PERENCANAAN")]
SLOT_KASI = [("KASI PEMERINTAHAN", r"PEMERINTAHAN"), ("KASI KESRA", r"KESRA|KESEJAHTERAAN"),
             ("KASI PELAYANAN/PEMBANGUNAN", r"PELAYANAN|PEMBANGUNAN")]


def lengkapi_struktur(orang, desa):
    """Struktur baku: 3 Kaur dan 3 Kasi. Slot yang belum ada barisnya ditambahkan sebagai jabatan kosong."""
    if sum(1 for p in orang if p["kelompok"] == "Perangkat Desa") >= 3:
        return []  # jabatan belum dirinci (mis. Kamalkuning): jangan menebak
    baru, ekstra = [], []
    for kel, slot in (("Kaur", SLOT_KAUR), ("Kasi", SLOT_KASI)):
        terpakai = set()
        for p in (p for p in orang if p["kelompok"] == kel):
            i = next((i for i, (_, pola) in enumerate(slot) if i not in terpakai and re.search(pola, p["jabatan"].upper())), None)
            if i is None:
                ekstra.append(p["jabatan"])
            else:
                terpakai.add(i)
        for i, (nama_slot, _) in enumerate(slot):
            if i not in terpakai:
                baru.append({"nama": None, "jabatan": nama_slot, "kelompok": kel, "lowong": True, "foto": None})
    if baru:
        info(desa, "Struktur baku belum lengkap di Excel, ditampilkan sebagai jabatan kosong: " + ", ".join(b["jabatan"] for b in baru))
    if ekstra:
        info(desa, "Jabatan di luar struktur baku (periksa penulisannya): " + ", ".join(ekstra))
    return baru


def baca_pagu(wb):
    """{kunci_desa: (nama, pagu_rupiah)} dari sheet Siltap (kolom A desa, kolom B pagu)."""
    pagu = {}
    if "Siltap" not in wb.sheetnames:
        warning("Siltap", "Sheet Siltap tidak ada; bagian Pagu Siltap tidak ditampilkan")
        return pagu
    for r in wb["Siltap"].iter_rows(values_only=True):
        if len(r) < 2 or not r[0] or kunci(r[0]) == "jumlah":
            continue
        v = r[1]
        n = int(v) if isinstance(v, (int, float)) else int(re.sub(r"\D", "", str(v or "")) or -1)
        if n >= 0:
            pagu[kunci(r[0])] = (str(r[0]).strip(), n)
    return pagu


def hitung_siltap(orang, pagu):
    """Kebutuhan siltap setahun: kades, sekdes, perangkat (staf tidak dihitung).
    Yang pensiun tetap dihitung sampai 'bulan_taliasih' bulan sejak bulan pensiun."""
    T = CFG["siltap"]
    tahun = HARI_INI.year
    total, kosong, bergeser = 0, 0, 0
    for p in orang:
        kel = p["kelompok"]
        if kel == "Staf":
            continue
        if p["lowong"]:
            kosong += 1
            continue
        tarif = T["kades"] if kel == "Kepala Desa" else T["sekdes"] if kel == "Sekretaris Desa" else T["perangkat"]
        bulan = T["bulan"]
        if p.get("pensiun"):
            tp = date.fromisoformat(p["pensiun"])
            terakhir = tp.year * 12 + tp.month - 1 + T["bulan_taliasih"] - 1  # indeks bulan gaji terakhir
            bulan = max(0, min(T["bulan"], terakhir - tahun * 12 + 1))
            if bulan < T["bulan"]:
                bergeser += 1
        total += tarif * bulan
    return {
        "pagu": pagu,
        "kebutuhan": total,
        "silpa": None if pagu is None else pagu - total,
        "kosong": kosong,
        "pensiun": bergeser,
    }


RESERVED = {"foto", "data", "logo", "index", "404"}


def halaman(tpl, base, judul, deskripsi, gambar):
    """Salinan index.html untuk satu alamat: atur basis jalur, judul, dan pratinjau tautan (WhatsApp dll)."""
    q = lambda t: html.escape(t, quote=True)
    h = tpl.replace('<meta name="base" content="./">', f'<meta name="base" content="{base}">', 1)
    h = re.sub(r"<title>.*?</title>", lambda _: f"<title>{q(judul)}</title>", h, count=1)
    og = (f'<meta property="og:type" content="website">\n<meta property="og:title" content="{q(judul)}">\n'
          f'<meta property="og:description" content="{q(deskripsi)}">\n')
    if gambar:
        og += f'<meta property="og:image" content="{q(gambar)}">\n'
    return h.replace("</head>", og + "</head>", 1)


def alamat_situs():
    """Alamat dasar situs untuk og:image: config 'url', atau diturunkan dari GITHUB_REPOSITORY."""
    if CFG.get("url"):
        return CFG["url"].rstrip("/") + "/"
    repo = os.environ.get("GITHUB_REPOSITORY", "")
    if "/" in repo:
        pemilik, nama = repo.split("/", 1)
        return f"https://{pemilik}.github.io/{nama}/"
    return None


def laporkan():
    gha = os.environ.get("GITHUB_ACTIONS") == "true"
    urut = {"ERROR": 0, "WARNING": 1, "INFO": 2}
    for lv, lok, msg in sorted(temuan, key=lambda t: urut[t[0]]):
        if gha and lv != "INFO":
            print(f"::{lv.lower()} title={lok}::{msg}")
        else:
            print(f"[{lv}] {lok}: {msg}")
    ringkasan = os.environ.get("GITHUB_STEP_SUMMARY")
    if ringkasan:
        with open(ringkasan, "a", encoding="utf-8") as f:
            f.write("## Hasil validasi\n\n| Level | Lokasi | Pesan |\n|---|---|---|\n")
            for lv, lok, msg in sorted(temuan, key=lambda t: urut[t[0]]):
                f.write(f"| {lv} | {lok} | {msg} |\n")
            if not temuan:
                f.write("| - | - | Tidak ada temuan |\n")


def main():
    if DIST.exists():
        shutil.rmtree(DIST)
    DIST.mkdir()
    desa = proses()
    n_err = sum(1 for t in temuan if t[0] == "ERROR")
    laporkan()
    if desa is None or n_err:
        print(f"\nBuild dibatalkan: {n_err} error. Perbaiki data/perangkat.xlsx lalu push lagi.")
        sys.exit(1)

    data = {
        "instansi": CFG["instansi"],
        "kecamatan": CFG["kecamatan"],
        "diperbarui": HARI_INI.isoformat(),
        "data_per": tanggal_data().isoformat(),
        "tahun": HARI_INI.year,
        "siltap_tarif": CFG["siltap"],
        "bulan": HARI_INI.month,
        "ringkasan": ringkas(desa),
        "desa": desa,
    }
    (DIST / "data.json").write_text(
        json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    tpl = (ROOT / "index.html").read_text(encoding="utf-8")
    ada_logo = (ROOT / "logo.png").exists()
    situs = alamat_situs()
    gambar = situs + "logo.png" if (situs and ada_logo) else None
    judul = f"Profil Perangkat Desa Kecamatan {CFG['kecamatan']}"
    (DIST / "index.html").write_text(
        halaman(tpl, "./", judul, f"Kepala desa dan perangkat desa se-Kecamatan {CFG['kecamatan']}, {CFG['instansi']}.", gambar),
        encoding="utf-8")
    for d in desa:  # alamat bersih: /<desa>/
        (DIST / d["slug"]).mkdir(exist_ok=True)
        (DIST / d["slug"] / "index.html").write_text(
            halaman(tpl, "../", f"Desa {d['nama']} - {judul}",
                    f"Kepala desa dan perangkat Desa {d['nama']}, Kecamatan {CFG['kecamatan']}.", gambar),
            encoding="utf-8")
    (DIST / ".nojekyll").write_text("")
    if (ROOT / "logo.png").exists():
        shutil.copyfile(ROOT / "logo.png", DIST / "logo.png")
    else:
        print("INFO: logo.png belum ada, header tampil tanpa lambang")
    r = data["ringkasan"]
    print(f"\nSelesai: {r['desa']} desa, {r['kades']} kades, {r['perangkat']} perangkat, {r['staf']} staf, "
          f"{r['foto']} berfoto, {r['lowong']} lowong, "
          f"{sum(1 for t in temuan if t[0]=='WARNING')} peringatan.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:  # galat program: tampilkan sebabnya jelas di Actions
        import traceback
        traceback.print_exc()
        pesan = f"{type(e).__name__}: {e}"
        print(f"::error title=Build gagal::{pesan}" if os.environ.get("GITHUB_ACTIONS") == "true" else f"Build gagal: {pesan}")
        ringkasan = os.environ.get("GITHUB_STEP_SUMMARY")
        if ringkasan:
            with open(ringkasan, "a", encoding="utf-8") as f:
                f.write(f"## Build gagal\n\n`{pesan}`\n\nRincian ada di log langkah build.\n")
        sys.exit(1)
