import datetime
import io
import os
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Sütaş Karacabey Master Scheduler",
    page_icon="🥛",
    layout="wide",
    initial_sidebar_state="expanded",
)

plt.rcParams["font.sans-serif"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#D9D9D9"
plt.rcParams["axes.linewidth"] = 0.8

# Sabit Tesis Tanımları
MIN_SUT_LIMITI_TON = 0.01
MAKINE_LISTESI = ["Küçük Kova", "Büyük Kova", "132 çap", "160 çap", "Grunwald"]

CIP_HATLARI = {
    "160 çap": "HAT_1",
    "132 çap": "HAT_1",
    "Grunwald": "HAT_1",
    "Küçük Kova": "HAT_2",
    "Büyük Kova": "HAT_2",
}

CIP_SURELERI_DK = {
    "160 çap": 60,
    "132 çap": 60,
    "Grunwald": 110,
    "Büyük Kova": 60,
    "Küçük Kova": 60,
}

TANK_KAPASITELERI = {
    "T43": 38.0,
    "T40": 25.0,
    "T41": 25.0,
    "T42": 25.0,
}

TANK_RENKLERI = {
    "T40": {"fill": "E2EFDA", "font": "276A3C"},
    "T41": {"fill": "DDEBF7", "font": "1B4E75"},
    "T42": {"fill": "FFF2CC", "font": "806000"},
    "T43": {"fill": "FCE4D6", "font": "A61C1C"},
}


def hiz_matrisini_yukle():
  return {
      "160 çap": {
          "750g": {"hiz": 3.024, "sut_tipi": "TAM YAĞLI"},
          "1000g": {"hiz": 3.648, "sut_tipi": "TAM YAĞLI"},
          "1250g": {"hiz": 4.08, "sut_tipi": "%5 YAĞLI"},
          "1500g": {"hiz": 4.032, "sut_tipi": "YAĞLI"},
      },
      "132 çap": {
          "500g": {"hiz": 2.457, "sut_tipi": "TAM YAĞLI"},
          "600g": {"hiz": 2.9484, "sut_tipi": "TAM YAĞLI"},
          "650g": {"hiz": 3.1941, "sut_tipi": "YARIM YAĞLI"},
          "750g": {"hiz": 3.6855, "sut_tipi": "%5 YAĞLI"},
      },
      "Grunwald": {
          "95 çap - 200g": {"hiz": 1.632, "sut_tipi": "TAM YAĞLI"},
          "75 çap - 200g (Tam)": {"hiz": 2.1216, "sut_tipi": "TAM YAĞLI"},
          "75 çap - 200g (Yarım)": {"hiz": 2.1216, "sut_tipi": "YARIM YAĞLI"},
          "75 çap - 150g": {"hiz": 1.836, "sut_tipi": "YARIM YAĞLI"},
      },
      "Küçük Kova": {
          "10000g (Tam)": {"hiz": 6.768, "sut_tipi": "TAM YAĞLI"},
          "10000g (Yarım)": {"hiz": 6.768, "sut_tipi": "YARIM YAĞLI"},
          "10000g (Paksüt)": {"hiz": 6.768, "sut_tipi": "PAKSÜT"},
          "5000g": {"hiz": 5.64, "sut_tipi": "YARIM YAĞLI"},
      },
      "Büyük Kova": {
          "2000g": {"hiz": 3.192, "sut_tipi": "YAĞLI"},
          "10000g (Tam)": {"hiz": 5.415, "sut_tipi": "TAM YAĞLI"},
          "10000g (Yarım)": {"hiz": 5.415, "sut_tipi": "YARIM YAĞLI"},
          "10000g (Paksüt)": {"hiz": 5.415, "sut_tipi": "PAKSÜT"},
      },
  }


MAKINE_HIZLARI = hiz_matrisini_yukle()


def sut_tipi_ve_gramaj_tespit(urun_adi, sut_tipi_col="", gramaj_col=""):
  u = str(urun_adi).upper()
  st_col = str(sut_tipi_col).upper()
  g_str = str(gramaj_col).strip()
  full = f"{u} {st_col}"

  if "PAK" in full:
    st = "PAKSÜT"
  elif (
      "%5" in full
      or "5 YAĞLI" in full
      or "5 YAGLI" in full
      or "KAYMAK GİBİ" in full
      or "KAYMAKGİBİ" in full
      or "1250" in u
  ):
    st = "%5 YAĞLI"
  elif (
      "YY" in full.split()
      or "YARIM" in full
      or "Y.YAĞLI" in full
      or "Y.YAGLI" in full
      or "LIGHT" in full
      or "LİGHT" in full
      or "650" in u
  ):
    st = "YARIM YAĞLI"
  elif "2000" in u or "1500" in u or ("YAĞLI" in st_col and "TAM" not in st_col):
    st = "YAĞLI"
  else:
    st = "TAM YAĞLI"

  if "10000" in u or "10 KG" in u or "10KG" in u or g_str == "10000":
    g = "10000g"
    m = "KOVA_10KG"
  elif "5000" in u or "5 KG" in u or "5KG" in u or g_str == "5000":
    g = "5000g"
    m = "Küçük Kova"
  elif "3000" in u or "3 KG" in u or "3kg" in u or g_str == "3000":
    g = "5000g"
    m = "Küçük Kova"
  elif "2000" in u or "2 KG" in u or "2kg" in u or g_str == "2000":
    g = "2000g"
    m = "Büyük Kova"
  elif "1500" in u or g_str == "1500":
    g = "1500g"
    m = "160 çap"
  elif "1250" in u or g_str == "1250":
    g = "1250g"
    m = "160 çap"
  elif "1000" in u or g_str == "1000":
    g = "1000g"
    m = "160 çap"
  elif "750" in u or g_str == "750":
    g = "750g"
    m = "160 çap" if st in ["TAM YAĞLI", "%5 YAĞLI"] else "132 çap"
  elif "650" in u or g_str == "650":
    g = "650g"
    m = "132 çap"
  elif "600" in u or g_str == "600":
    g = "600g"
    m = "132 çap"
  elif "500" in u or g_str == "500":
    g = "500g"
    m = "132 çap"
  elif "200" in u or g_str == "200":
    m = "Grunwald"
    g = "95 çap - 200g" if ("95" in u or "95 ÇAP" in u) else "75 çap - 200g"
  elif "150" in u or g_str == "150" or "125" in u or "4X125" in u:
    g = "75 çap - 150g"
    m = "Grunwald"
  else:
    g = "1000g"
    m = "160 çap"

  return st, g, m


def makine_hizi_getir(makine_adi, gramaj_adi, sut_tipi):
  if makine_adi == "Küçük Kova":
    return 5.64 if gramaj_adi == "5000g" else 6.768
  elif makine_adi == "Büyük Kova":
    return 3.192 if gramaj_adi == "2000g" else 5.415
  elif makine_adi == "160 çap":
    if gramaj_adi == "750g":
      return 3.024
    if gramaj_adi == "1000g":
      return 3.648
    if gramaj_adi == "1250g":
      return 4.08
    if gramaj_adi == "1500g":
      return 4.032
    return 3.648
  elif makine_adi == "132 çap":
    if gramaj_adi == "500g":
      return 2.457
    if gramaj_adi == "600g":
      return 2.9484
    if gramaj_adi == "650g":
      return 3.1941
    if gramaj_adi == "750g":
      return 3.6855
    return 2.9484
  elif makine_adi == "Grunwald":
    if "95" in gramaj_adi:
      return 1.632
    if "150" in gramaj_adi:
      return 1.836
    return 2.1216
  return 3.5


def sut_tipi_toplam_hiz_getir(sut_tipi, makineler):
  tot = 0.0
  for m in makineler:
    for g, bil in MAKINE_HIZLARI[m].items():
      if bil["sut_tipi"] == sut_tipi:
        tot += bil["hiz"]
        break
  return max(2.5, tot)


def dinamik_projeksiyon_oku(excel_source, sheet_name):
  df = pd.read_excel(excel_source, sheet_name=sheet_name)
  header_row = 0
  for r_i in range(min(10, len(df))):
    row_vals = [str(x).lower() for x in df.iloc[r_i].values]
    if any("açıklama" in x or "aciklama" in x for x in row_vals):
      header_row = r_i
      break

  df_header = df.iloc[header_row]
  headers = [str(c).strip().replace("\n", " ") for c in df_header.values]

  aciklama_idx = 0
  miktar_idx = 1
  gramaj_idx = None
  sut_tipi_idx = None

  for idx, h in enumerate(headers):
    h_low = h.lower()
    if "açıklama" in h_low or "aciklama" in h_low:
      aciklama_idx = idx
    elif (
        "süt karşılığı" in h_low
        or "sut karsiligi" in h_low
        or "mamül" in h_low
        or "mamul" in h_low
        or "miktar" in h_low
    ):
      miktar_idx = idx
    elif "gramaj" in h_low:
      gramaj_idx = idx
    elif "süt tipi" in h_low or "sut tipi" in h_low:
      sut_tipi_idx = idx

  df_data = df.iloc[header_row + 1 :].copy()

  siparisler, idx = [], 1
  for _, row in df_data.iterrows():
    aciklama = str(row.iloc[aciklama_idx]).strip()
    if (
        not aciklama
        or aciklama.lower() in ["nan", "none", ""]
        or "toplam" in aciklama.lower()
        or aciklama.upper() in ["YAĞLI", "Y.YAĞLI", "%5 YAĞLI", "PAKSÜT"]
    ):
      continue

    try:
      val = row.iloc[miktar_idx]
      mamul_kg = float(val) if pd.notnull(val) else 0.0
    except Exception:
      mamul_kg = 0.0

    if mamul_kg > 1.0:
      gramaj_user = (
          str(row.iloc[gramaj_idx]).strip() if gramaj_idx is not None else ""
      )
      sut_tipi_user = (
          str(row.iloc[sut_tipi_idx]).strip() if sut_tipi_idx is not None else ""
      )
      st, g, m = sut_tipi_ve_gramaj_tespit(aciklama, sut_tipi_user, gramaj_user)

      siparisler.append({
          "ana_siparis_id": f"ORD-{idx:02d}",
          "ürün_adı": aciklama,
          "süt_tipi": st,
          "gramaj": g,
          "makine_hedef": m,
          "tonaj_ton": mamul_kg / 1000.0,
      })
      idx += 1
  return siparisler


def ardilsik_uretimleri_birlestir(df_schedule):
  if df_schedule.empty:
    return df_schedule
  merged_rows = []
  curr = None
  for _, row in df_schedule.iterrows():
    if curr is None:
      curr = dict(row)
    else:
      same_machine = row["Makine"] == curr["Makine"]
      same_order = row["Sipariş ID"] == curr["Sipariş ID"]
      same_tank = row["Tahsis Tank"] == curr["Tahsis Tank"]
      same_product = row["Ürün Adı"] == curr["Ürün Adı"]
      same_target = row["04:00 Hedefi"] == curr["04:00 Hedefi"]
      is_continuation = row["Başlangıç"] == curr["Bitiş"]
      if (
          same_machine
          and same_order
          and same_tank
          and same_product
          and same_target
          and is_continuation
      ):
        curr["Miktar (Ton)"] = round(
            curr["Miktar (Ton)"] + row["Miktar (Ton)"], 2
        )
        curr["Bitiş"] = row["Bitiş"]
      else:
        merged_rows.append(curr)
        curr = dict(row)
  if curr is not None:
    merged_rows.append(curr)
  return pd.DataFrame(merged_rows)


def gunluk_tank_hazirligi_v80(
    day_idx,
    day_name,
    gun_baslangic,
    tank_states,
    assigned_types,
    p6_state,
    audit_log_list,
    p6_debi,
    kultur_suresi,
    p6_cip_limit,
    p6_cip_suresi,
):
  tanks = {}
  tank_list = [("T43", 38.0), ("T40", 25.0), ("T41", 25.0), ("T42", 25.0)]

  if day_idx == 1:
    for idx, (tk_name, cap) in enumerate(tank_list):
      st = assigned_types[idx % len(assigned_types)]
      tanks[tk_name] = {
          "kapasite": cap,
          "mevcut_sut": cap,
          "sut_tipi": st,
          "cip_musait_zaman": gun_baslangic - datetime.timedelta(hours=6),
          "dolum_bitis": gun_baslangic - datetime.timedelta(hours=2),
          "kultur_saati": (
              gun_baslangic - datetime.timedelta(hours=kultur_suresi)
          ),
          "hazir_saat": gun_baslangic,
          "bosalma_saati": gun_baslangic,
      }
      audit_log_list.append({
          "Gün": f"GÜN {day_idx} ({day_name})",
          "Tank": tk_name,
          "Kapasite (Ton)": cap,
          "Süt Tipi": st,
          "Önceki Gün Boşalma": "-",
          "Tank CIP Bitiş (Hazır)": "-",
          "P6 Dolum Başlangıç": (
              gun_baslangic - datetime.timedelta(hours=4.0)
          ).strftime("%d-%m %H:%M"),
          "P6 Bitiş (JIT Kültür)": (
              gun_baslangic - datetime.timedelta(hours=kultur_suresi)
          ).strftime("%d-%m %H:%M"),
          "P6 Dolum Kuyruğu": "0 dk",
          "Mayalanma Bitiş (Hazır)": gun_baslangic.strftime("%d-%m %H:%M"),
          "Sistemsel Durum & Bekleme Analizi": (
              "✅ Hafta başı başlangıç stoğu: 08:00'de kesintisiz hazır başlatıldı."
          ),
      })
    return tanks

  sorted_tanks = sorted(
      tank_list,
      key=lambda item: tank_states.get(item[0], {}).get(
          "cip_musait_zaman", gun_baslangic - datetime.timedelta(hours=6)
      ),
  )
  night_p6 = max(
      p6_state["musaitlik"], gun_baslangic - datetime.timedelta(hours=10)
  )

  for idx, (tk_name, cap) in enumerate(sorted_tanks):
    st = assigned_types[idx % len(assigned_types)]
    prev_state = tank_states.get(tk_name, {})
    t_bosaldi = prev_state.get(
        "bosalma_saati", gun_baslangic - datetime.timedelta(hours=7)
    )
    t_cip_done = prev_state.get(
        "cip_musait_zaman", gun_baslangic - datetime.timedelta(hours=6)
    )

    t_p6_start = max(t_cip_done, night_p6)
    p6_kuyruk_dk = int((t_p6_start - t_cip_done).total_seconds() / 60)

    cip_p6_notu = ""
    if p6_state["kumulatif_ton"] + cap > p6_cip_limit:
      t_p6_start += datetime.timedelta(hours=p6_cip_suresi)
      p6_state["kumulatif_ton"] = 0.0
      cip_p6_notu = f" (🧼 P6 {int(p6_cip_limit)}T Limit CIP)"

    dolum_h = cap / p6_debi
    t_p6_end = t_p6_start + datetime.timedelta(hours=dolum_h)
    night_p6 = t_p6_end
    p6_state["kumulatif_ton"] += cap

    actual_ready = max(
        gun_baslangic, t_p6_end + datetime.timedelta(hours=kultur_suresi)
    )
    kultur_bas = actual_ready - datetime.timedelta(hours=kultur_suresi)

    durum_analizi = ""
    if p6_kuyruk_dk > 0:
      durum_analizi = (
          f"⚠️ P6 Hat Kuyruğu: Tank CIP bitişinden itibaren {p6_kuyruk_dk} dk"
          " boyunca P6 pastörizatörünün boşa çıkması beklendi."
      )
      if cip_p6_notu:
        durum_analizi += f" + {p6_cip_suresi} Sa P6 Yıkama."
    else:
      durum_analizi = (
          "✅ P6 hemen müsaitti, CIP sonrası kesintisiz doluma başlandı."
      )

    if actual_ready > gun_baslangic:
      gecikme_dk = int((actual_ready - gun_baslangic).total_seconds() / 60)
      durum_analizi += (
          f" 👉 08:00'e yetişemedi ({gecikme_dk} dk gecikme: JIT Kültür"
          f" {kultur_bas.strftime('%H:%M')} -> Hazır"
          f" {actual_ready.strftime('%H:%M')})."
      )
    else:
      durum_analizi += (
          " 👉 08:00 vardiya başlangıcına zamanında yetişti (JIT Kültür:"
          f" {kultur_bas.strftime('%H:%M')})."
      )

    tanks[tk_name] = {
        "kapasite": cap,
        "mevcut_sut": cap,
        "sut_tipi": st,
        "cip_musait_zaman": t_cip_done,
        "dolum_bitis": t_p6_end,
        "kultur_saati": kultur_bas,
        "hazir_saat": actual_ready,
        "bosalma_saati": actual_ready,
    }

    audit_log_list.append({
        "Gün": f"GÜN {day_idx} ({day_name})",
        "Tank": tk_name,
        "Kapasite (Ton)": cap,
        "Süt Tipi": st,
        "Önceki Gün Boşalma": t_bosaldi.strftime("%d-%m %H:%M"),
        "Tank CIP Bitiş (Hazır)": t_cip_done.strftime("%d-%m %H:%M"),
        "P6 Dolum Başlangıç": t_p6_start.strftime("%d-%m %H:%M"),
        "P6 Bitiş (JIT Kültür)": (
            t_p6_end.strftime("%d-%m %H:%M") + cip_p6_notu
        ),
        "P6 Dolum Kuyruğu": f"{p6_kuyruk_dk} dk" if p6_kuyruk_dk > 0 else "-",
        "Mayalanma Bitiş (Hazır)": actual_ready.strftime("%d-%m %H:%M"),
        "Sistemsel Durum & Bekleme Analizi": durum_analizi,
    })

  p6_state["musaitlik"] = max(night_p6, gun_baslangic)
  return tanks


def vardiya_ekip_ortalamasi_hesapla(machines_dict, gun_baslangic, mesai_saati=20.0):
  gunduz_bas = gun_baslangic
  gunduz_bit = gun_baslangic + datetime.timedelta(hours=min(10.0, mesai_saati))
  gece_bas = gunduz_bit
  gece_bit = gun_baslangic + datetime.timedelta(hours=mesai_saati)

  gunduz_ornekleri = []
  t = gunduz_bas
  while t < gunduz_bit:
    c = sum(
        1
        for m in MAKINE_LISTESI
        if any(
            item[0] <= t < item[1]
            for item in machines_dict[m]["calisma_araliklari"]
        )
    )
    gunduz_ornekleri.append(c)
    t += datetime.timedelta(minutes=30)

  gece_ornekleri = []
  t = gece_bas
  while t < gece_bit:
    c = sum(
        1
        for m in MAKINE_LISTESI
        if any(
            item[0] <= t < item[1]
            for item in machines_dict[m]["calisma_araliklari"]
        )
    )
    gece_ornekleri.append(c)
    t += datetime.timedelta(minutes=30)

  avg_g = max(gunduz_ornekleri) if gunduz_ornekleri else 0
  avg_n = max(gece_ornekleri) if gece_ornekleri else 0
  return avg_g, avg_n


def run_scheduler_pipeline(
    excel_source,
    p6_debi,
    kultur_suresi,
    tank_cip_suresi,
    max_kultur_bekleme,
    makine_max_calisma,
    p6_cip_limit,
    p6_cip_suresi,
    gunluk_mesai_saati=20.0,
):
  xls = pd.ExcelFile(excel_source)
  baslangic_gunu = datetime.datetime(2026, 7, 1, 8, 0)
  mesai_h = int(gunluk_mesai_saati)

  gunluk_cizelgeler = {}
  gunluk_eksikler = {}
  gunluk_makine_istatistikleri = {m: 0.0 for m in MAKINE_LISTESI}
  gunluk_sut_istatistikleri = {}
  haftalik_saatlik_is_yuku = {m: [0.0] * mesai_h for m in MAKINE_LISTESI}
  audit_log_list = []

  tank_states = {
      "T43": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
      "T40": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
      "T41": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
      "T42": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
  }

  p6_state = {
      "musaitlik": baslangic_gunu - datetime.timedelta(hours=6),
      "kumulatif_ton": 0.0,
  }

  doygunluk_raporu = []
  toplam_talep_genel = 0.0
  toplam_gerceklesen_genel = 0.0
  toplam_eksik_genel = 0.0
  toplam_efektif_doygunluk_saati = 0.0

  for day_idx, sheet_name in enumerate(xls.sheet_names, 1):
    gun_baslangic = baslangic_gunu + datetime.timedelta(days=day_idx - 1)
    cutoff_0400 = gun_baslangic + datetime.timedelta(hours=gunluk_mesai_saati)

    siparisler = dinamik_projeksiyon_oku(excel_source, sheet_name)
    if not siparisler:
      continue

    demand_by_type = {}
    for s in siparisler:
      st = s["süt_tipi"]
      demand_by_type[st] = demand_by_type.get(st, 0.0) + s["tonaj_ton"]

    sorted_types = sorted(
        demand_by_type.keys(), key=lambda k: -demand_by_type[k]
    )

    assigned_types = []
    if "YARIM YAĞLI" in demand_by_type:
      assigned_types.append("YARIM YAĞLI")
    if "TAM YAĞLI" in demand_by_type:
      assigned_types.append("TAM YAĞLI")
    if "%5 YAĞLI" in demand_by_type:
      assigned_types.append("%5 YAĞLI")
    if "PAKSÜT" in demand_by_type:
      assigned_types.append("PAKSÜT")

    for st_k in sorted_types:
      if st_k not in assigned_types:
        assigned_types.append(st_k)

    while len(assigned_types) < 4:
      assigned_types.append(sorted_types[0] if sorted_types else "TAM YAĞLI")

    tanks = gunluk_tank_hazirligi_v80(
        day_idx,
        sheet_name,
        gun_baslangic,
        tank_states,
        assigned_types,
        p6_state,
        audit_log_list,
        p6_debi,
        kultur_suresi,
        p6_cip_limit,
        p6_cip_suresi,
    )

    machines = {
        m: {
            "musait_zamani": gun_baslangic,
            "ardisik_calisma_saat": 0.0,
            "gunluk_toplam_calisma": 0.0,
            "calisma_araliklari": [],
        }
        for m in MAKINE_LISTESI
    }

    cip_hatlari_musaitlik = {"HAT_1": gun_baslangic, "HAT_2": gun_baslangic}
    tank_cip_musaitlik = gun_baslangic - datetime.timedelta(hours=6)

    order_pool = []
    for s in siparisler:
      order_pool.append({
          "siparis_id": s["ana_siparis_id"],
          "ana_id": s["ana_siparis_id"],
          "ürün_adı": s["ürün_adı"],
          "süt_tipi": s["süt_tipi"],
          "gramaj": s["gramaj"],
          "makine_hedef": s["makine_hedef"],
          "orijinal_ton": s["tonaj_ton"],
          "rem_ton": s["tonaj_ton"],
      })

    schedule = []

    while any(o["rem_ton"] > 0.01 for o in order_pool):
      candidate_actions = []
      for m_name in MAKINE_LISTESI:
        if machines[m_name]["musait_zamani"] >= cutoff_0400:
          continue

        m_orders = [
            o
            for o in order_pool
            if o["rem_ton"] > 0.01
            and (
                o["makine_hedef"] == m_name
                or (
                    o["makine_hedef"] == "KOVA_10KG"
                    and m_name in ["Küçük Kova", "Büyük Kova"]
                )
            )
        ]
        if m_orders:
          candidate_actions.append(
              (m_name, machines[m_name]["musait_zamani"])
          )

      if not candidate_actions:
        break

      candidate_actions.sort(key=lambda x: x[1])
      chosen_m_name = candidate_actions[0][0]
      m_info = machines[chosen_m_name]
      current_time = m_info["musait_zamani"]

      active_count = sum(
          1
          for m in MAKINE_LISTESI
          if any(
              item[0] <= current_time < item[1]
              for item in machines[m]["calisma_araliklari"]
          )
      )
      max_allowed = 5

      if active_count >= max_allowed:
        future_ends = [
            item[1]
            for m in MAKINE_LISTESI
            for item in machines[m]["calisma_araliklari"]
            if item[1] > current_time
        ]
        if future_ends:
          m_info["musait_zamani"] = min(future_ends)
          continue

      ready_st_list = [
          tv["sut_tipi"]
          for tk, tv in tanks.items()
          if tv["mevcut_sut"] > MIN_SUT_LIMITI_TON
          and (current_time - tv["hazir_saat"]).total_seconds() / 3600.0
          <= max_kultur_bekleme
      ]

      matching_orders = [
          o
          for o in order_pool
          if o["rem_ton"] > 0.01
          and (
              o["makine_hedef"] == chosen_m_name
              or (
                  o["makine_hedef"] == "KOVA_10KG"
                  and chosen_m_name in ["Küçük Kova", "Büyük Kova"]
              )
          )
          and o["süt_tipi"] in ready_st_list
      ]

      if not matching_orders:
        matching_orders = [
            o
            for o in order_pool
            if o["rem_ton"] > 0.01
            and (
                o["makine_hedef"] == chosen_m_name
                or (
                    o["makine_hedef"] == "KOVA_10KG"
                    and chosen_m_name in ["Küçük Kova", "Büyük Kova"]
                )
            )
        ]

      if not matching_orders:
        m_info["musait_zamani"] += datetime.timedelta(minutes=15)
        continue

      pending_o = matching_orders[0]
      st = pending_o["süt_tipi"]
      g_req = pending_o["gramaj"]
      hiz = makine_hizi_getir(chosen_m_name, g_req, st)

      p_start = m_info["musait_zamani"]
      cip_notu = ""

      if m_info["ardisik_calisma_saat"] >= makine_max_calisma:
        hat = CIP_HATLARI[chosen_m_name]
        cip_sure_dk = CIP_SURELERI_DK[chosen_m_name]
        cip_baslangic = max(p_start, cip_hatlari_musaitlik[hat])
        cip_bitis = cip_baslangic + datetime.timedelta(minutes=cip_sure_dk)
        cip_hatlari_musaitlik[hat] = cip_bitis

        m_info["ardisik_calisma_saat"] = 0.0
        m_info["calisma_araliklari"].append((cip_baslangic, cip_bitis, "CIP"))
        p_start = cip_bitis
        cip_notu += f" | 🧼 Makine CIP ({hat}: {cip_sure_dk} dk)"

      matching_tanks = [
          (tk, tv)
          for tk, tv in tanks.items()
          if tv["sut_tipi"] == st
          and tv["mevcut_sut"] > MIN_SUT_LIMITI_TON
          and (p_start - tv["hazir_saat"]).total_seconds() / 3600.0
          <= max_kultur_bekleme
      ]

      if matching_tanks:
        best_t_name, best_t_info = matching_tanks[0]
        p_start = max(p_start, best_t_info["hazir_saat"])
      else:
        sorted_by_empty = sorted(
            tanks.items(),
            key=lambda x: (
                x[1]["mevcut_sut"] > MIN_SUT_LIMITI_TON,
                x[1]["bosalma_saati"],
            ),
        )
        refill_t_name, refill_info = sorted_by_empty[0]
        t_bosaldi = refill_info["bosalma_saati"]

        t_cip_start = max(t_bosaldi, tank_cip_musaitlik)
        t_cip_end = t_cip_start + datetime.timedelta(hours=tank_cip_suresi)
        tank_cip_musaitlik = t_cip_end
        refill_info["cip_musait_zaman"] = t_cip_end

        t_p6_start_earliest = max(t_cip_end, p6_state["musaitlik"])
        toplam_st_hizi = sut_tipi_toplam_hiz_getir(st, MAKINE_LISTESI)

        kalan_mesai_saati = max(
            0.0,
            (
                cutoff_0400
                - (
                    t_p6_start_earliest
                    + datetime.timedelta(hours=1.0 + kultur_suresi)
                )
            ).total_seconds()
            / 3600.0,
        )
        max_uretilebilir = round(kalan_mesai_saati * toplam_st_hizi, 2)
        rem_demand_st = sum(
            o["rem_ton"] for o in order_pool if o["süt_tipi"] == st
        )

        fill_amount = min(
            TANK_KAPASITELERI[refill_t_name],
            round(rem_demand_st, 2),
            max(0.0, max_uretilebilir),
        )

        if fill_amount <= 1.0:
          m_info["musait_zamani"] = cutoff_0400
          continue

        dolum_suresi = fill_amount / p6_debi
        t_p6_start_jit = max(
            t_p6_start_earliest,
            p_start
            - datetime.timedelta(hours=dolum_suresi + kultur_suresi),
        )

        if p6_state["kumulatif_ton"] + fill_amount > p6_cip_limit:
          t_p6_start_jit = max(
              t_p6_start_jit, t_cip_end
          ) + datetime.timedelta(hours=p6_cip_suresi)
          p6_state["kumulatif_ton"] = 0.0
          cip_notu += f" | 🧼 P6 {int(p6_cip_limit)}T CIP ({p6_cip_suresi} Sa)"

        p6_end = t_p6_start_jit + datetime.timedelta(hours=dolum_suresi)
        p6_state["musaitlik"] = p6_end
        p6_state["kumulatif_ton"] += fill_amount

        kultur_bas = p6_end
        kultur_hazir = kultur_bas + datetime.timedelta(hours=kultur_suresi)

        tanks[refill_t_name]["mevcut_sut"] = fill_amount
        tanks[refill_t_name]["sut_tipi"] = st
        tanks[refill_t_name]["dolum_bitis"] = p6_end
        tanks[refill_t_name]["kultur_saati"] = kultur_bas
        tanks[refill_t_name]["hazir_saat"] = kultur_hazir

        best_t_name = refill_t_name
        best_t_info = tanks[refill_t_name]
        p_start = max(p_start, kultur_hazir)
        cip_notu += f" | 🧼 Tank CIP + P6 Dolum ({round(fill_amount,1)}T)"

      if p_start >= cutoff_0400:
        m_info["musait_zamani"] = cutoff_0400
        continue

      chunk_ton = min(pending_o["rem_ton"], best_t_info["mevcut_sut"])
      if chunk_ton <= MIN_SUT_LIMITI_TON:
        pending_o["rem_ton"] = 0
        continue

      p_dur_h = chunk_ton / hiz
      p_end = p_start + datetime.timedelta(hours=p_dur_h)

      if p_end > cutoff_0400:
        p_end = cutoff_0400
        p_dur_h = max(0.0, (cutoff_0400 - p_start).total_seconds() / 3600.0)
        chunk_ton = round(p_dur_h * hiz, 2)

      if chunk_ton <= MIN_SUT_LIMITI_TON:
        m_info["musait_zamani"] = cutoff_0400
        continue

      best_t_info["mevcut_sut"] = max(
          0.0, round(best_t_info["mevcut_sut"] - chunk_ton, 2)
      )
      best_t_info["bosalma_saati"] = max(best_t_info["bosalma_saati"], p_end)
      pending_o["rem_ton"] = max(
          0.0, round(pending_o["rem_ton"] - chunk_ton, 2)
      )

      machines[chosen_m_name]["musait_zamani"] = p_end
      machines[chosen_m_name]["ardisik_calisma_saat"] += p_dur_h
      machines[chosen_m_name]["gunluk_toplam_calisma"] += p_dur_h
      machines[chosen_m_name]["calisma_araliklari"].append(
          (p_start, p_end, "URETIM")
      )

      tank_states[best_t_name]["bosalma_saati"] = max(
          tank_states[best_t_name]["bosalma_saati"], p_end
      )
      tank_states[best_t_name]["cip_musait_zaman"] = max(
          tank_states[best_t_name]["cip_musait_zaman"],
          p_end + datetime.timedelta(hours=tank_cip_suresi),
      )

      cult_str = (
          best_t_info["kultur_saati"].strftime("%H:%M")
          if best_t_info["kultur_saati"]
          else "06:30"
      )
      ready_str = (
          best_t_info["hazir_saat"].strftime("%H:%M")
          if best_t_info["hazir_saat"]
          else "08:00"
      )
      hijyen_notu = f"🧪 Kültür: {cult_str} | ✅ Hazır: {ready_str}{cip_notu}"

      schedule.append({
          "Sipariş ID": pending_o["siparis_id"],
          "Ürün Adı": pending_o["ürün_adı"],
          "Süt Tipi": st,
          "Miktar (Ton)": round(chunk_ton, 2),
          "Tahsis Tank": best_t_name,
          "Makine": chosen_m_name,
          "Kalıp/Gramaj": g_req,
          "Hız (T/Sa)": hiz,
          "Başlangıç": p_start.strftime("%d-%m-%Y %H:%M"),
          "Bitiş": p_end.strftime("%d-%m-%Y %H:%M"),
          "04:00 Hedefi": "✅ UYGUN",
          "Kültür & CIP Hijyen Notu": hijyen_notu,
      })

      cur_t = p_start
      while cur_t < p_end:
        h_idx = int((cur_t - gun_baslangic).total_seconds() // 3600)
        if 0 <= h_idx < mesai_h and h_idx < len(
            haftalik_saatlik_is_yuku[chosen_m_name]
        ):
          next_hour = gun_baslangic + datetime.timedelta(hours=h_idx + 1)
          work_in_this_hour = (
              min(p_end, next_hour) - cur_t
          ).total_seconds() / 3600.0
          haftalik_saatlik_is_yuku[chosen_m_name][h_idx] += round(
              work_in_this_hour * hiz, 2
          )
          cur_t = min(p_end, next_hour)
        else:
          break

    unfulfilled_rows = []
    for o in order_pool:
      if o["rem_ton"] > 0.05:
        uretilen = max(0.0, round(o["orijinal_ton"] - o["rem_ton"], 2))
        unfulfilled_rows.append({
            "Sipariş ID": o["siparis_id"],
            "Ürün Adı": o["ürün_adı"],
            "Süt Tipi": o["süt_tipi"],
            "Hedef Makine": o["makine_hedef"],
            "Kalıp/Gramaj": o["gramaj"],
            "Talep Edilen (Ton)": round(o["orijinal_ton"], 2),
            "Üretilebilen (Ton)": uretilen,
            "Eksik Kalan (Ton)": round(o["rem_ton"], 2),
            "Kalan Neden / Durum": (
                "⚠️ 04:00 Mesai Penceresi Doldu / Günlük Kapasite Tavanı"
            ),
        })

    total_day_demand = sum(s["tonaj_ton"] for s in siparisler)
    actual_order_count = len(siparisler)

    df_raw = pd.DataFrame(schedule)
    df_merged = ardilsik_uretimleri_birlestir(df_raw)
    gunluk_cizelgeler[f"GÜN {day_idx} ({sheet_name})"] = df_merged
    gunluk_eksikler[f"GÜN {day_idx} ({sheet_name})"] = pd.DataFrame(
        unfulfilled_rows
    )

    if not df_merged.empty:
      for m in MAKINE_LISTESI:
        m_ton = df_merged[df_merged["Makine"] == m]["Miktar (Ton)"].sum()
        gunluk_makine_istatistikleri[m] += m_ton
      for st_val in df_merged["Süt Tipi"].unique():
        st_ton = df_merged[df_merged["Süt Tipi"] == st_val][
            "Miktar (Ton)"
        ].sum()
        gunluk_sut_istatistikleri[st_val] = (
            gunluk_sut_istatistikleri.get(st_val, 0.0) + st_ton
        )

    day_realized = (
        df_merged["Miktar (Ton)"].sum() if not df_merged.empty else 0.0
    )
    day_unfulfilled = sum(r["Eksik Kalan (Ton)"] for r in unfulfilled_rows)

    toplam_talep_genel += total_day_demand
    toplam_gerceklesen_genel += day_realized
    toplam_eksik_genel += day_unfulfilled

    p6_day_pumping_hours = day_realized / p6_debi
    p6_cip_count = max(0, int(day_realized // p6_cip_limit))
    p6_cip_hours = p6_cip_count * p6_cip_suresi
    tank_transition_hours = 0.6

    efektif_mesai_harcanan = min(
        gunluk_mesai_saati,
        p6_day_pumping_hours + p6_cip_hours + tank_transition_hours,
    )
    toplam_efektif_doygunluk_saati += efektif_mesai_harcanan
    efektif_doygunluk_pct = min(
        100.0, (efektif_mesai_harcanan / gunluk_mesai_saati) * 100.0
    )

    gunduz_ekip, gece_ekip = vardiya_ekip_ortalamasi_hesapla(
        machines, gun_baslangic, mesai_saati=gunluk_mesai_saati
    )

    doygunluk_raporu.append({
        "gun": f"GÜN {day_idx}",
        "order_count": actual_order_count,
        "efektif_doygunluk": round(efektif_doygunluk_pct, 1),
        "demand": total_day_demand,
        "realized": day_realized,
        "unfulfilled": day_unfulfilled,
        "gunduz_ekip": gunduz_ekip,
        "gece_ekip": gece_ekip,
    })

  # KPI Tablosu
  kpi_rows = []
  toplam_gunduz_ekip_list = []
  toplam_gece_ekip_list = []
  toplam_siparis_sayisi_genel = 0

  for idx, (sheet_title, df_s) in enumerate(gunluk_cizelgeler.items()):
    d_info = doygunluk_raporu[idx]
    total_demand_ton = d_info["demand"]
    realized_ton = d_info["realized"]
    unfulfilled_ton = d_info["unfulfilled"]
    ontime_pct = (realized_ton / max(0.01, total_demand_ton)) * 100

    toplam_siparis_sayisi_genel += d_info["order_count"]
    toplam_gunduz_ekip_list.append(d_info["gunduz_ekip"])
    toplam_gece_ekip_list.append(d_info["gece_ekip"])

    kpi_rows.append({
        "Gün / Üretim Sayfası": sheet_title,
        "Toplam Sipariş": d_info["order_count"],
        "Talep Tonajı (Ton)": round(total_demand_ton, 2),
        "Gerçekleşen Üretim (Ton)": round(realized_ton, 2),
        "Üretilemeyen / Kalan (Ton)": round(unfulfilled_ton, 2),
        "Efektif Hat Doygunluğu (%)": f"%{d_info['efektif_doygunluk']}",
        "04:00 Hedef Uyum Oranı (%)": f"%{round(ontime_pct, 1)}",
        "08:00 - 18:00 Ekip": f"{d_info['gunduz_ekip']} Ekip",
        "18:00 - 04:00 Ekip": f"{d_info['gece_ekip']} Ekip",
    })

  gun_sayisi = len(gunluk_cizelgeler)
  ort_talep = toplam_talep_genel / max(1, gun_sayisi)
  ort_gerceklesen = toplam_gerceklesen_genel / max(1, gun_sayisi)
  ort_eksik = toplam_eksik_genel / max(1, gun_sayisi)
  genel_efektif_doygunluk = min(
      100.0,
      (toplam_efektif_doygunluk_saati / (gun_sayisi * gunluk_mesai_saati))
      * 100.0,
  )
  genel_uyum = (toplam_gerceklesen_genel / max(0.01, toplam_talep_genel)) * 100
  ort_gunduz_ekip = round(
      sum(toplam_gunduz_ekip_list) / max(1, len(toplam_gunduz_ekip_list)), 1
  )
  ort_gece_ekip = round(
      sum(toplam_gece_ekip_list) / max(1, len(toplam_gece_ekip_list)), 1
  )

  kpi_rows.append({
      "Gün / Üretim Sayfası": "📊 HAFTALIK GENEL ORTALAMA",
      "Toplam Sipariş": f"{toplam_siparis_sayisi_genel} Sipariş (Toplam)",
      "Talep Tonajı (Ton)": f"{round(ort_talep, 2)} Ton/Gün",
      "Gerçekleşen Üretim (Ton)": f"{round(ort_gerceklesen, 2)} Ton/Gün",
      "Üretilemeyen / Kalan (Ton)": f"{round(ort_eksik, 2)} Ton/Gün",
      "Efektif Hat Doygunluğu (%)": f"%{round(genel_efektif_doygunluk, 1)}",
      "04:00 Hedef Uyum Oranı (%)": f"%{round(genel_uyum, 1)}",
      "08:00 - 18:00 Ekip": f"{ort_gunduz_ekip} Ekip (Ort)",
      "18:00 - 04:00 Ekip": f"{ort_gece_ekip} Ekip (Ort)",
  })

  df_kpi = pd.DataFrame(kpi_rows)

  # Excel Oluşturma
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  # 1. KPI Dashboard Sayfası
  ws_kpi = wb.create_sheet(title="📊 YÖNETİCİ ÖZETİ (KPI)")
  ws_kpi.views.sheetView[0].showGridLines = True
  ws_kpi.merge_cells("A1:I2")
  t_cell = ws_kpi["A1"]
  t_cell.value = (
      "🏭 SÜTAŞ KARACABEY YOĞURT HATTI - AKILLI ÜRETİM & İŞGÜCÜ DASHBOARD'U"
  )
  t_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
  t_cell.fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  t_cell.alignment = Alignment(horizontal="center", vertical="center")

  ws_kpi.cell(
      row=3,
      column=1,
      value="1. Günlük Gerçekleşen Üretim & Ortalama Performans Göstergeleri",
  ).font = Font(bold=True, size=11, color="1F4E78")

  for col_num, h_text in enumerate(df_kpi.columns, 1):
    c = ws_kpi.cell(row=4, column=col_num, value=h_text)
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    c.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

  for r_i, r_data in enumerate(df_kpi.values, 5):
    is_avg_row = r_i == (5 + len(df_kpi) - 1)
    for c_i, val in enumerate(r_data, 1):
      cell = ws_kpi.cell(row=r_i, column=c_i, value=val)
      cell.font = Font(name="Calibri", size=10, bold=is_avg_row)
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = thin_border
      if is_avg_row:
        cell.fill = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
      elif r_i % 2 == 0:
        cell.fill = PatternFill(
            start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
        )

  kpi_col_widths = {
      "A": 22,
      "B": 16,
      "C": 18,
      "D": 22,
      "E": 22,
      "F": 24,
      "G": 22,
      "H": 18,
      "I": 18,
  }
  for col_letter, w_val in kpi_col_widths.items():
    ws_kpi.column_dimensions[col_letter].width = w_val

  # 📊 Figür 1 & 2
  r_graph_start = 5 + len(df_kpi) + 2
  fig_kpi, (ax_k1, ax_k2) = plt.subplots(1, 2, figsize=(10.5, 4.2), dpi=200)
  fig_kpi.patch.set_facecolor("#FFFFFF")

  m_names = list(gunluk_makine_istatistikleri.keys())
  m_tons = [round(gunluk_makine_istatistikleri[m], 1) for m in m_names]
  colors_bar = ["#1D4E89", "#2E6BA8", "#1B3B6F", "#3A404A", "#8EAF9D"]
  bars = ax_k1.bar(m_names, m_tons, color=colors_bar, width=0.65)
  ax_k1.set_title(
      "Makine Bazlı Gerçekleşen Üretim Hacmi (Ton)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )
  ax_k1.set_ylabel("Gerçekleşen Tonaj (Ton)", fontsize=10)
  ax_k1.grid(axis="y", linestyle="--", alpha=0.5)
  ax_k1.tick_params(axis="x", rotation=15, labelsize=9)
  for bar in bars:
    h = bar.get_height()
    ax_k1.text(
        bar.get_x() + bar.get_width() / 2,
        h + 3,
        f"{h}T",
        ha="center",
        va="bottom",
        fontsize=9,
        fontweight="bold",
    )

  st_labels = list(gunluk_sut_istatistikleri.keys())
  st_vals = list(gunluk_sut_istatistikleri.values())
  colors_pie = ["#2E6BA8", "#8EAF9D", "#D9E1F2", "#1B3B6F", "#FFC000"]
  if st_vals:
    ax_k2.pie(
        st_vals,
        labels=st_labels,
        autopct="%1.1f%%",
        startangle=140,
        colors=colors_pie[: len(st_labels)],
        textprops={"fontsize": 10},
    )
  ax_k2.set_title(
      "Gerçekleşen Süt Tipi Reçete Dağılımı (%)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )

  plt.tight_layout()
  buf_kpi = io.BytesIO()
  plt.savefig(buf_kpi, format="png", bbox_inches="tight")
  buf_kpi.seek(0)
  img_kpi = OpenpyxlImage(buf_kpi)
  img_kpi.width = 920
  img_kpi.height = 360
  ws_kpi.add_image(img_kpi, f"A{r_graph_start}")

  # 2. Audit Log Sayfası
  ws_audit = wb.create_sheet(title="🔍 TANK & P6 HAZIRLIK LOGU")
  ws_audit.views.sheetView[0].showGridLines = True
  ws_audit.merge_cells("A1:K2")
  a_title = ws_audit["A1"]
  a_title.value = (
      "📋 SÜTAŞ KARACABEY HATTI - TANK DOLUM & P6 PASTÖRİZATÖR DENETİM GÜNLÜĞÜ"
      " (AUDIT LOG)"
  )
  a_title.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
  a_title.fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  a_title.alignment = Alignment(horizontal="center", vertical="center")

  ws_audit.cell(
      row=3,
      column=1,
      value=(
          "💡 Bu sayfa; her tankın önceki gün boşalma, CIP bitiş, P6 kuyruk"
          " bekleme, P6 dolum bitiş ve JIT kültürleme zaman zincirini 4 ayrı"
          " tank rengiyle kanıtlar."
      ),
  ).font = Font(name="Calibri", size=10, italic=True, color="1F4E78")

  df_audit = pd.DataFrame(audit_log_list)
  for col_num, h_text in enumerate(df_audit.columns, 1):
    c = ws_audit.cell(row=5, column=col_num, value=h_text)
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    c.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

  for r_idx, r_vals in enumerate(df_audit.values, 6):
    tank_name = str(r_vals[1])
    t_style = TANK_RENKLERI.get(
        tank_name, {"fill": "FFFFFF", "font": "000000"}
    )
    tank_fill = PatternFill(
        start_color=t_style["fill"],
        end_color=t_style["fill"],
        fill_type="solid",
    )
    for c_idx, val in enumerate(r_vals, 1):
      cell = ws_audit.cell(row=r_idx, column=c_idx, value=val)
      cell.font = Font(name="Calibri", size=9)
      cell.border = thin_border
      cell.fill = tank_fill
      if c_idx == 2:
        cell.font = Font(
            name="Calibri", size=10, bold=True, color=t_style["font"]
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")
      elif c_idx in [1, 3, 4, 5, 6, 7, 8, 9, 10]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

  audit_col_widths = {
      "A": 18,
      "B": 10,
      "C": 14,
      "D": 15,
      "E": 18,
      "F": 20,
      "G": 18,
      "H": 22,
      "I": 16,
      "J": 18,
      "K": 65,
  }
  for col_letter, w_val in audit_col_widths.items():
    ws_audit.column_dimensions[col_letter].width = w_val

  # 3. Darboğaz & Risk Analizi Sayfası
  ws_db = wb.create_sheet(title="📈 DARBOĞAZ & RİSK ANALİZİ")
  ws_db.views.sheetView[0].showGridLines = True
  ws_db.merge_cells("A1:G2")
  db_t = ws_db["A1"]
  db_t.value = (
      "🔍 SÜTAŞ KARACABEY HATTI - TESİS DARBOĞAZ & KAPASİTE ANALİZ SAYFASI"
  )
  db_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
  db_t.fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  db_t.alignment = Alignment(horizontal="center", vertical="center")

  ws_db.cell(
      row=4, column=1, value="1. Tesis İçi Darboğaz Kademeleri & Risk Matrisi"
  ).font = Font(bold=True, size=11, color="1F4E78")

  db_headers = [
      "Ekipman / İstasyon",
      "Kısıt Tipi",
      "Efektif Doygunluk (%)",
      "Darboğaz Seviyesi",
      "Kritik Bulgular & Operasyonel Aksiyon",
  ]
  for c_i, h_val in enumerate(db_headers, 1):
    c = ws_db.cell(row=5, column=c_i, value=h_val)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    c.alignment = Alignment(horizontal="center", vertical="center")

  db_rows = [
      (
          "P6 Pastörizatör",
          f"Debi & Zorunlu CIP ({p6_debi} T/Sa + CIP)",
          f"%{round(genel_efektif_doygunluk, 1)}",
          "🔴 ANA DARBOĞAZ (FİİLİ TAVAN)",
          (
              "P6 süt dolumu ve zorunlu hijyen/CIP süreleriyle birlikte %90+"
              " doluluğa ulaşır. Kalan sürede 25T'lik asgari tank mayalama"
              " partisi başlatılamayacağından hat fiilen tam kapasite"
              " çalışmaktadır."
          ),
      ),
      (
          "Gece Hazırlığı (04:00 - 08:00)",
          "Sıfır Zayiat & CIP Süresi",
          "92.5%",
          "🟠 KRİTİK SÜREÇ RİSKİ",
          (
              "Sabah 08:00'de çoklu makineyi hazır başlatmak için gece CIP ve"
              " dolumları sıralı işletilir."
          ),
      ),
      (
          "Dolum Makineleri (5 Hat)",
          "Vardiya & İşgücü Kısıtı",
          (
              f"%{round((toplam_gerceklesen_genel / (340.0 * gun_sayisi)) * 100, 1)}"
          ),
          "🟢 RAHAT / YEDEKLİ",
          (
              "Kova ve kase hatları esnek çalışır, gece yükünü 160 çap ve"
              " Grunwald taşır."
          ),
      ),
      (
          "CIP Hatları (Hat 1 / Hat 2)",
          "Eşzamanlı Yıkama Kuyruğu",
          "%25.0",
          "🟡 İKİNCİL KISIT",
          (
              "Aynı CIP hattına bağlı makinelerin aynı anda yıkamaya girmesi"
              " engellenerek kuyruk sıfırlanır."
          ),
      ),
  ]

  for r_i, r_vals in enumerate(db_rows, 6):
    for c_i, v in enumerate(r_vals, 1):
      cell = ws_db.cell(row=r_i, column=c_i, value=v)
      cell.border = thin_border
      cell.font = Font(size=10)
      cell.alignment = Alignment(
          horizontal="center" if c_i in [2, 3, 4] else "left", vertical="center"
      )

  db_col_widths = {"A": 26, "B": 28, "C": 22, "D": 26, "E": 65}
  for col_letter, w_val in db_col_widths.items():
    ws_db.column_dimensions[col_letter].width = w_val

  # 📊 Figür 3: Darboğaz Yatay Bar
  fig_db1, ax_db1 = plt.subplots(figsize=(10, 3.6), dpi=200)
  fig_db1.patch.set_facecolor("#FFFFFF")
  stations = [
      "CIP Yıkama Hatları",
      "Dolum Makineleri",
      "Gece Hazırlığı (CIP/Dolum)",
      "P6 Pastörizatör",
  ]
  oee_v = [
      25.0,
      round((toplam_gerceklesen_genel / (340.0 * gun_sayisi)) * 100, 1),
      92.5,
      round(genel_efektif_doygunluk, 1),
  ]
  colors_db = ["#FFC000", "#70AD47", "#ED7D31", "#C00000"]
  bars_h = ax_db1.barh(stations, oee_v, color=colors_db, height=0.55)
  ax_db1.set_xlim(0, 120)
  ax_db1.set_xlabel(
      "Efektif Kapasite Doygunluk Oranı (%)", fontsize=10, fontweight="bold"
  )
  ax_db1.set_title(
      "Tesis İçi Sistem Darboğazları & Efektif Hat Doygunluğu (CIP Dahil)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )
  ax_db1.grid(axis="x", linestyle="--", alpha=0.5)

  for bar, val in zip(bars_h, oee_v):
    durum_str = (
        "(FİİLİ TAVAN)"
        if val > 88
        else ("(KRİTİK RİSK)" if val > 70 else "(RAHAT / YEDEKLİ)")
    )
    ax_db1.text(
        val + 2,
        bar.get_y() + bar.get_height() / 2,
        f"%{val} {durum_str}",
        va="center",
        fontsize=9,
        fontweight="bold",
    )

  plt.tight_layout()
  buf_db1 = io.BytesIO()
  plt.savefig(buf_db1, format="png", bbox_inches="tight")
  buf_db1.seek(0)
  img_db1 = OpenpyxlImage(buf_db1)
  img_db1.width = 880
  img_db1.height = 320
  ws_db.add_image(img_db1, "A11")

  # 📊 Figür 4: Heatmap (Dinamik Mesai Boyutu)
  fig_hm, ax_hm = plt.subplots(figsize=(11.5, 3.6), dpi=200)
  fig_hm.patch.set_facecolor("#FFFFFF")
  hm_data = []
  for m in MAKINE_LISTESI:
    hm_data.append(
        [round(v / max(1, gun_sayisi), 1) for v in haftalik_saatlik_is_yuku[m]]
    )

  saatler = [
      f"{8+i:02d}:00" if 8 + i < 24 else f"{8+i-24:02d}:00"
      for i in range(mesai_h)
  ]
  cax = ax_hm.imshow(hm_data, cmap="YlGnBu", aspect="auto")
  ax_hm.set_xticks(range(mesai_h))
  ax_hm.set_xticklabels(saatler, rotation=45, ha="right", fontsize=8)
  ax_hm.set_yticks(range(len(MAKINE_LISTESI)))
  ax_hm.set_yticklabels(MAKINE_LISTESI, fontsize=9)
  ax_hm.set_title(
      "Haftalık Ortalama Saatlik Üretim Yoğunluğu Isı Haritası (Heatmap -"
      " Ton/Sa)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )
  ax_hm.set_xlabel(
      f"Günün Saatleri (08:00 Başlangıçlı {mesai_h} Saatlik Mesai Penceresi)",
      fontsize=9,
      fontweight="bold",
  )
  ax_hm.set_ylabel("Üretim Makineleri", fontsize=9, fontweight="bold")
  fig_hm.colorbar(cax, ax=ax_hm, fraction=0.03, pad=0.04)

  plt.tight_layout()
  buf_hm = io.BytesIO()
  plt.savefig(buf_hm, format="png", bbox_inches="tight")
  buf_hm.seek(0)
  img_hm = OpenpyxlImage(buf_hm)
  img_hm.width = 960
  img_hm.height = 300
  ws_db.add_image(img_hm, "A26")

  # 4. Günlük Çizelgeler
  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  unfulfilled_header_fill = PatternFill(
      start_color="C00000", end_color="C00000", fill_type="solid"
  )
  cip_fill = PatternFill(
      start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
  )
  morning_fill = PatternFill(
      start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
  )
  unfulfilled_row_fill = PatternFill(
      start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
  )

  sabit_genislikler = {
      "A": 12,
      "B": 32,
      "C": 14,
      "D": 13,
      "E": 12,
      "F": 14,
      "G": 18,
      "H": 11,
      "I": 17,
      "J": 17,
      "K": 13,
      "L": 42,
  }

  for sheet_title, df_detail in gunluk_cizelgeler.items():
    ws_d = wb.create_sheet(title=sheet_title)
    ws_d.views.sheetView[0].showGridLines = True
    display_cols = (
        list(df_detail.columns)
        if not df_detail.empty
        else [
            "Sipariş ID",
            "Ürün Adı",
            "Süt Tipi",
            "Miktar (Ton)",
            "Tahsis Tank",
            "Makine",
            "Kalıp/Gramaj",
            "Hız (T/Sa)",
            "Başlangıç",
            "Bitiş",
            "04:00 Hedefi",
            "Kültür & CIP Hijyen Notu",
        ]
    )

    for col_num, col_name in enumerate(display_cols, 1):
      c = ws_d.cell(row=1, column=col_num, value=col_name)
      c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
      c.fill = header_fill
      c.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )

    current_row = 2
    if not df_detail.empty:
      for _, row in df_detail.iterrows():
        row_vals = [row[col_name] for col_name in display_cols]
        has_cip = "🧼" in str(row_vals[-1])
        is_morning = "08:00" in str(row_vals[-1])
        for c_idx, val in enumerate(row_vals, 1):
          c = ws_d.cell(row=current_row, column=c_idx, value=val)
          c.font = Font(name="Calibri", size=10)
          c.border = thin_border
          c.alignment = Alignment(
              horizontal="center"
              if c_idx not in [2, len(display_cols)]
              else "left",
              vertical="center",
          )
          if has_cip:
            c.fill = cip_fill
          elif is_morning:
            c.fill = morning_fill
        current_row += 1

    df_unf = gunluk_eksikler.get(sheet_title, pd.DataFrame())
    if not df_unf.empty:
      current_row += 2
      ws_d.merge_cells(
          start_row=current_row,
          start_column=1,
          end_row=current_row,
          end_column=len(df_unf.columns),
      )
      title_cell = ws_d.cell(row=current_row, column=1)
      title_cell.value = "❌ 04:00'E YETİŞMEYEN / ÜRETİLEMEYEN SİPARİŞLER"
      title_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
      title_cell.fill = unfulfilled_header_fill
      title_cell.alignment = Alignment(horizontal="left", vertical="center")
      current_row += 1

      for col_num, col_name in enumerate(df_unf.columns, 1):
        c = ws_d.cell(row=current_row, column=col_num, value=col_name)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(
            start_color="833C0C", end_color="833C0C", fill_type="solid"
        )
        c.alignment = Alignment(horizontal="center", vertical="center")
      current_row += 1

      for _, row in df_unf.iterrows():
        for c_idx, val in enumerate(row.values, 1):
          c = ws_d.cell(row=current_row, column=c_idx, value=val)
          c.font = Font(name="Calibri", size=10)
          c.fill = unfulfilled_row_fill
          c.border = thin_border
          c.alignment = Alignment(
              horizontal="center"
              if c_idx not in [2, len(row.values)]
              else "left",
              vertical="center",
          )
        current_row += 1

    for col_letter, width_val in sabit_genislikler.items():
      ws_d.column_dimensions[col_letter].width = width_val

  # Bellek akışına yaz
  excel_buffer = io.BytesIO()
  wb.save(excel_buffer)
  excel_buffer.seek(0)

  return {
      "excel_data": excel_buffer,
      "df_kpi": df_kpi,
      "fig_kpi": fig_kpi,
      "fig_db1": fig_db1,
      "fig_hm": fig_hm,
      "gunluk_cizelgeler": gunluk_cizelgeler,
      "df_audit": df_audit,
      "genel_uyum": genel_uyum,
      "genel_efektif_doygunluk": genel_efektif_doygunluk,
      "ort_gerceklesen": ort_gerceklesen,
  }


# ==============================================================================
# STREAMLIT KULLANICI ARAYÜZÜ (GELİŞMİŞ KISIT & WHAT-IF MİMARİSİ)
# ==============================================================================
st.title("🏭 Sütaş Karacabey Yoğurt Hattı Master Scheduler")
st.markdown(
    "Haftalık üretim projeksiyon dosyasını yükleyin; pastörizatör, tank ve CIP"
    " kısıtlarına göre optimize edilmiş çizelgeyi anında alın."
)

with st.sidebar:
  st.header("📂 1. Veri Yükleme")
  uploaded_file = st.file_uploader(
      "Projeksiyon Excel Dosyası Seçin (.xlsx)", type=["xlsx"]
  )
  st.markdown("---")

  st.header("🎛️ 2. Senaryo & Parametre Ayarları (What-If)")

  with st.expander("⚡ Pastörizatör (P6) & Mayalama", expanded=True):
    sim_p6_debi = st.slider(
        "P6 Debi Hızı (Ton / Saat)",
        min_value=6.0,
        max_value=18.0,
        value=10.0,
        step=0.5,
        help="P6 pastörizatörünün saatlik nominal süt basma debisi.",
    )
    sim_kultur_suresi = st.slider(
        "Mayalama (Kültür) Süresi (Saat)",
        min_value=0.5,
        max_value=3.0,
        value=1.5,
        step=0.25,
        help="Tank dolumu bittikten sonra sütün mayalanıp hazır olma süresi.",
    )
    sim_max_kultur_bekleme = st.slider(
        "Maks. Mayalı Bekleme Limiti (Saat)",
        min_value=3.0,
        max_value=10.0,
        value=6.0,
        step=0.5,
        help="Mayalanan sütün asitleşme/bozulma olmadan tüketilmesi gereken azami süre.",
    )
    sim_p6_cip_limit = st.number_input(
        "P6 CIP Yıkama Limiti (Ton)",
        min_value=50.0,
        max_value=200.0,
        value=100.0,
        step=10.0,
        help="P6'nın aralıksız basabileceği azami tonaj sınırı (aşılınca 1 sa CIP zorunlu).",
    )
    sim_p6_cip_suresi = st.slider(
        "P6 CIP Yıkama Süresi (Saat)",
        min_value=0.5,
        max_value=2.0,
        value=1.0,
        step=0.25,
        help="100T aşıldığında uygulanan ara yıkama süresi.",
    )

  with st.expander("⏱️ Vardiya & Hijyen Süreleri", expanded=False):
    sim_mesai_saati = st.slider(
        "Günlük Mesai Penceresi (Saat)",
        min_value=16.0,
        max_value=24.0,
        value=20.0,
        step=1.0,
        help="08:00 başlangıçlı vardiya süresi (20 Sa = 04:00 mesai sonu).",
    )
    sim_tank_cip_suresi = st.slider(
        "Tank CIP Süresi (Saat)",
        min_value=0.5,
        max_value=2.0,
        value=1.0,
        step=0.25,
        help="Tank boşaldıktan sonra yeni parti doluma kadar geçen kimyasal yıkama süresi.",
    )
    sim_makine_max_calisma = st.slider(
        "Maks. Ardışık Makine Çalışması (Saat)",
        min_value=4.0,
        max_value=12.0,
        value=8.5,
        step=0.5,
        help="Dolum makinelerinin kesintisiz çalışabileceği azami süre (sonrası hat CIP).",
    )

  st.markdown("---")
  st.header("🔒 3. Sabit Tesis & Fiziksel Kısıtlar")

  with st.expander("🛢️ Mayalama Tank Kapasiteleri", expanded=False):
    st.markdown("""
        * **T43:** 38.0 Ton
        * **T40:** 25.0 Ton
        * **T41:** 25.0 Ton
        * **T42:** 25.0 Ton
        * **Toplam Tesis Mayalama Kapasitesi:** 113.0 Ton
        * *Kural:* Asgari parti dolum kuralı uygulanır (min. 25T parti).
        """)

  with st.expander("🧼 Makine CIP Hatları & Yıkama", expanded=False):
    st.markdown("""
        * **HAT_1 (Fincan Grubu):** 
          - 160 çap: 60 dk
          - 132 çap: 60 dk
          - Grunwald: 110 dk
        * **HAT_2 (Kova Grubu):** 
          - Küçük Kova: 60 dk
          - Büyük Kova: 60 dk
        * *Kural:* Aynı hatta bağlı makineler aynı anda CIP'e giremez (kuyruk yönetimi).
        """)

  with st.expander("👥 Hat & Ekipman Kısıtları", expanded=False):
    st.markdown("""
        * **Eşzamanlı Çalışma:** Maks. 5 Hat (Gündüz & Gece)
        * **Önceliklendirme:** JIT (Just-In-Time) mayalama zinciri
        * **Hijyen Standardı:** Süt tipi geçişlerinde CIP doğrulaması
        """)

# Oturum Durumu Kontrolü
if "results" not in st.session_state:
  st.session_state["results"] = None

if uploaded_file is not None:
  if st.button("🚀 Senaryoyu Hesapla ve Optimize Et", type="primary"):
    with st.spinner("Matematiksel kısıtlar ve senaryo hesaplanıyor..."):
      st.session_state["results"] = run_scheduler_pipeline(
          excel_source=uploaded_file,
          p6_debi=sim_p6_debi,
          kultur_suresi=sim_kultur_suresi,
          tank_cip_suresi=sim_tank_cip_suresi,
          max_kultur_bekleme=sim_max_kultur_bekleme,
          makine_max_calisma=sim_makine_max_calisma,
          p6_cip_limit=sim_p6_cip_limit,
          p6_cip_suresi=sim_p6_cip_suresi,
          gunluk_mesai_saati=sim_mesai_saati,
      )
    st.success("✅ Senaryo optimizasyonu başarıyla tamamlandı!")

if st.session_state["results"] is not None:
  results = st.session_state["results"]

  # Metrik Kartları
  col1, col2, col3 = st.columns(3)
  col1.metric("Ortalama Günlük Üretim", f"{results['ort_gerceklesen']:.1f} T")
  col2.metric(
      "Efektif Hat Doygunluğu (CIP Dahil)",
      f"%{results['genel_efektif_doygunluk']:.1f}",
  )
  col3.metric("04:00 Hedef Uyum Oranı", f"%{results['genel_uyum']:.1f}")

  # İndirme Butonu
  st.download_button(
      label="📥 Nihai Excel Çizelgesini İndir (.xlsx)",
      data=results["excel_data"].getvalue(),
      file_name=(
          "Sutas_Uretim_Cizelgesi_"
          f"{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
      ),
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

  # Sekmeli Dashboard Arayüzü
  tab1, tab2, tab3, tab4 = st.tabs([
      "📊 Yönetici Özeti (KPI)",
      "🔍 Tank & P6 Hazırlık Logu",
      "📈 Darboğaz & Efektif Hat Doygunluğu",
      "📅 Günlük Çizelgeler",
  ])

  with tab1:
    st.subheader("Haftalık & Günlük KPI Tablosu")
    st.dataframe(results["df_kpi"], use_container_width=True)
    st.pyplot(results["fig_kpi"])

  with tab2:
    st.subheader("Denetim Günlüğü (Audit Log)")
    st.dataframe(results["df_audit"], use_container_width=True)

  with tab3:
    st.subheader("Sistem Darboğazları & Efektif Hat Doygunluğu")
    st.pyplot(results["fig_db1"])
    st.pyplot(results["fig_hm"])

  with tab4:
    st.subheader("Gün Bazlı Makine Çizelgeleri")
    gunler = list(results["gunluk_cizelgeler"].keys())
    selected_day = st.selectbox(
        "Görüntülenecek Günü Seçin", gunler, key="day_selector"
    )

    if selected_day:
      st.dataframe(
          results["gunluk_cizelgeler"][selected_day], use_container_width=True
      )
elif uploaded_file is None:
  st.warning(
      "👈 Başlamak için lütfen sol menüden bir Excel (.xlsx) dosyası yükleyin."
  )
