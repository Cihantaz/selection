import os
import io
import json
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, session

app = Flask(__name__)
app.secret_key = "exam_secret_key"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

TABLO_BASLIKLARI = [
    ("bolum_adi", "Bölüm Adı"),
    ("puan_turu", "Puan Türü"),
    ("burs_orani", "Burs Oranı"),
    ("taban_siralama", "Taban Sıralama"),
    ("taban_puan", "Taban Puan"),
    ("tavan_puan", "Tavan Puan"),
    ("ucret", "Ücret"),
    ("dil", "Dil"),
    ("etiket", "Etiket"),
    ("riskli_t", "Riskli T"),
    ("z_riskli", "Z Riskli"),
    ("parametre", "Parametre"),
]

def temizle_sayi(s):
    if isinstance(s, str):
        s = s.replace('.', '').replace(',', '')
    try:
        return int(s)
    except:
        return 0

def etiketle(ogr_siralama, taban, z_riskli):
    try:
        ogr_siralama = int(ogr_siralama)
        taban = int(taban)
    except:
        return "Bilinmiyor"
    if taban >= ogr_siralama:
        return "Uygun"
    elif z_riskli is not None and taban >= z_riskli:
        return "Riskli"
    else:
        return "Uygunsuz"

def analiz_yap(df, eklenenler):
    result = []
    for p in eklenenler:
        ogr_siralama_int = temizle_sayi(p["puan"])
        sinir_int = temizle_sayi(p["sinir"])
        riskli_t_int = temizle_sayi(p.get("riskli_t", 0))
        z = ogr_siralama_int - sinir_int
        z_riskli = z - riskli_t_int if riskli_t_int else None

        df_filtered = df.copy()
        if p["tur"] and 'Puan Türü' in df_filtered.columns:
            df_filtered = df_filtered[df_filtered['Puan Türü'].astype(str).str.strip().str.upper() == p["tur"]]

        bos_veya_eksi_bolumler = []
        if 'En Düşük Sıralama' in df_filtered.columns:
            def kontrol_et(x):
                try:
                    x_sayi = int(str(x).replace('.', '').replace(',', ''))
                except:
                    return False
                return x_sayi > (z_riskli if z_riskli is not None else z)

            df_filtered_main = df_filtered[df_filtered['En Düşük Sıralama'].apply(kontrol_et)]
            bos_veya_eksi_bolumler = df_filtered[df_filtered['En Düşük Sıralama'].apply(lambda x: pd.isna(x) or str(x).strip() in ["", "-"])]
        else:
            df_filtered_main = df_filtered

        for _, row in df_filtered_main.iterrows():
            program_adi = str(row.get('Program Adı', '')).strip()
            burs = str(row.get('Burs/İndirim', '')).strip() if 'Burs/İndirim' in row else ''
            if not burs:
                for burs_kw in ["Burslu", "Ücretli", "%50 İndirimli", "%25 İndirimli", "%75 İndirimli", "%100 Burslu"]:
                    if burs_kw.lower() in program_adi.lower():
                        burs = burs_kw
                        break
            ucret = row.get('Ücret', '')
            if ucret:
                try:
                    ucret_num = float(str(ucret).replace('.', '').replace(',', '').replace('₺', '').strip())
                    # 3.5 uyumlu
                    ucret = "{:,.0f}".format(ucret_num).replace(",", ".") + " TL"
                except:
                    ucret = str(ucret) + " TL"
            dil = "EN" if "(ingilizce)" in program_adi.lower() else "TR"
            etiket = etiketle(ogr_siralama_int, row.get('En Düşük Sıralama', 0), z_riskli)

            result.append({
                "bolum_adi": program_adi,
                "puan_turu": row.get('Puan Türü', ''),
                "burs_orani": burs,
                "taban_siralama": row.get('En Düşük Sıralama', ''),
                "taban_puan": row.get('Taban Puan', ''),
                "tavan_puan": row.get('Tavan Puan', ''),
                "ucret": ucret,
                "dil": dil,
                "etiket": etiket,
                "riskli_t": riskli_t_int,
                "z_riskli": z_riskli if z_riskli is not None else "",
                "parametre": "{}/ {}/ {}".format(p["tur"], p["puan"], p["sinir"])
            })

        for _, row in bos_veya_eksi_bolumler.iterrows():
            program_adi = str(row.get('Program Adı', '')).strip()
            burs = str(row.get('Burs/İndirim', '')).strip() if 'Burs/İndirim' in row else ''
            if not burs:
                for burs_kw in ["Burslu", "Ücretli", "%50 İndirimli", "%25 İndirimli", "%75 İndirimli", "%100 Burslu"]:
                    if burs_kw.lower() in program_adi.lower():
                        burs = burs_kw
                        break
            ucret = row.get('Ücret', '')
            if ucret:
                try:
                    ucret_num = float(str(ucret).replace('.', '').replace(',', '').replace('₺', '').strip())
                    ucret = "{:,.0f}".format(ucret_num).replace(",", ".") + " TL"
                except:
                    ucret = str(ucret) + " TL"
            dil = "EN" if "(ingilizce)" in program_adi.lower() else "TR"
            etiket = etiketle(ogr_siralama_int, row.get('En Düşük Sıralama', 0), z_riskli)

            # Python 3.5'te dictionary comprehension ve any fonksiyonu çalışır
            if not any(r['bolum_adi'] == program_adi and r['taban_siralama'] == row.get('En Düşük Sıralama', '') for r in result):
                result.append({
                    "bolum_adi": program_adi,
                    "puan_turu": row.get('Puan Türü', ''),
                    "burs_orani": burs,
                    "taban_siralama": row.get('En Düşük Sıralama', ''),
                    "taban_puan": row.get('Taban Puan', ''),
                    "tavan_puan": row.get('Tavan Puan', ''),
                    "ucret": ucret,
                    "dil": dil,
                    "etiket": etiket,
                    "riskli_t": riskli_t_int,
                    "z_riskli": z_riskli if z_riskli is not None else "",
                    "parametre": "{}/ {}/ {}".format(p["tur"], p["puan"], p["sinir"])
                })

    if not result:
        result = [{"bolum_adi": "Sonuç bulunamadı"}]
    return result

@app.route("/", methods=["GET", "POST"])
def index():
    eklenenler = []
    adsoyad = ""
    talep_bolum = ""
    result = None
    tablo_basliklari = TABLO_BASLIKLARI

    if request.method == "POST":
        adsoyad_ve_bolum = request.form.get("adsoyad", "")
        if "," in adsoyad_ve_bolum:
            adsoyad, talep_bolum = [x.strip() for x in adsoyad_ve_bolum.split(",", 1)]
        else:
            adsoyad = adsoyad_ve_bolum.strip()

        try:
            eklenenler = json.loads(request.form.get("eklenenler", "[]"))
        except:
            eklenenler = []

        file = request.files.get("veri_dosya")
        if not file or file.filename == "":
            flash("Excel dosyası yükleyin.", "danger")
        else:
            try:
                df = pd.read_excel(file)
                df.columns = df.columns.str.strip()
                result = analiz_yap(df, eklenenler)
                session["analiz_df"] = result
                session["adsoyad"] = adsoyad
                session["talep_bolum"] = talep_bolum
            except Exception as e:
                flash("Excel okuma hatası: {}".format(e), "danger")

        return render_template("index.html", adsoyad=adsoyad_ve_bolum, eklenenler=eklenenler, result=result, tablo_basliklari=tablo_basliklari)

    return render_template("index.html", adsoyad="", eklenenler=[], result=None, tablo_basliklari=tablo_basliklari)

@app.route("/indir")
def indir():
    analiz_df = session.get("analiz_df", None)
    adsoyad = session.get("adsoyad", "")
    talep_bolum = session.get("talep_bolum", "")

    if analiz_df:
        df = pd.DataFrame(analiz_df)
        output = io.BytesIO()
        import xlsxwriter
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sheet = 'Sonuclar'
            df.to_excel(writer, index=False, sheet_name=sheet, startrow=1)
            worksheet = writer.sheets[sheet]
            worksheet.write(0, 0, "Talep Edilen Bölüm: {}".format(talep_bolum))
        output.seek(0)
        dosya_adi = (adsoyad.strip().replace(" ", "_") if adsoyad else "analiz_sonuclari") + ".xlsx"
        return send_file(output, as_attachment=True, download_name=dosya_adi, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        flash("Önce analiz yapmalısınız.", "warning")
        return render_template("index.html", adsoyad="", eklenenler=[], result=None, tablo_basliklari=TABLO_BASLIKLARI)

if __name__ == "__main__":
    app.run(debug=True)
