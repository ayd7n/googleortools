import argparse
import calendar
import json
import re
import sqlite3
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

KAYNAK_TABLO = "Talepler"
HEDEF_TABLO = "talepler_x"
TALEPLER_X_SILINME_LOG_TABLOSU = "talepler_x_silinme_log"
TALEPLER_X_ADIM_OZET_TABLOSU = "talepler_x_adim_ozet"
TALEPLER_X_TOPLULASTIRMA_LOG_TABLOSU = "talepler_x_toplulastirma_log"
MALZEME_UYUM_TABLOSU = "Malzeme_Fabrika_Firin_Hat_Tarih_Uyumu"
PD_TALEPLER_TABLOSU = "PD_Talepler"
HAM_TALEP_SNAPSHOT_TABLOSU = "talepler_x_ham_snapshot"
PLANLAMA_SQLITE_DOSYASI = "planlama.sqlite"
PLANLAMA_TABLOLARI = (
    PD_TALEPLER_TABLOSU,
    MALZEME_UYUM_TABLOSU,
    "Playground_x",
    "Renk_Plani_X",
    "Aylik_Teorik_Kapasite_X",
)
PLANLAMA_OPSIYONEL_TABLOLARI = (
    "Hattaki_Aylik_Max_Malzeme_Sayis",
)
AYLIK_ADT_DESENI = re.compile(r"^(\d{2})\.(\d{4})\s+ADT$")
TARIH_SUTUN_DESENI = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
MALZEME_UYUM_HATA_TIPLERI = (
    "Gramaj_Eksik",
    "Verim_Eksik",
    "Damla_Sayisi_Eksik",
    "Kapasite_Eksik",
    "Kapasite_Yetersiz",
    "Gecersiz_Tarih_Araligi",
)
JOB_PROGRESS_PREFIX = "JOB_PROGRESS::"

# Not:
# Excel tarafinda bazi dosyalarda Turkce karakterler bozulmus gelebiliyor.
# Bu nedenle hem duzgun hem de bozuk yazimlari birlikte tutuyoruz.
SILINECEK_SUTUNLAR = {
    "Malzeme kısa metni",
    "Ürün Hiyerarşisi 3",
    "Ürün Hiyerarşisi 4",
    "Ambalaj Tipi Tanımı",
    "Palet İç Miktarı",
    "Net ağırlığı",
    "Satış bürosu tanım",
    "Satış bürosu",
    "YI/YD",
    "Müşteri",
    "Müşteri Tanımı",
    "Yaratan",
    "Yaratma tarihi",
    "Saat",
    "Versiyon",
    "Malzeme kÄ±sa metni",
    "ÃœrÃ¼n HiyerarÅŸisi 3",
    "ÃœrÃ¼n HiyerarÅŸisi 4",
    "Ambalaj Tipi TanÄ±mÄ±",
    "Palet Ä°Ã§ MiktarÄ±",
    "Net aÄŸÄ±rlÄ±ÄŸÄ±",
    "SatÄ±ÅŸ bÃ¼rosu tanÄ±m",
    "SatÄ±ÅŸ bÃ¼rosu",
    "MÃ¼ÅŸteri",
    "MÃ¼ÅŸteri TanÄ±mÄ±",
}


def sql_kimlik_kacaga_dayanikli(metin: str) -> str:
    """Tablo/sutun adlarini SQL icinde guvenli sekilde cift tirnaklar."""
    return '"' + metin.replace('"', '""') + '"'


def normalize_lookup_token(deger: Any) -> str:
    metin = str(deger or "").strip().lower().replace("\u0131", "i")
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(ch for ch in metin if not unicodedata.combining(ch))
    metin = metin.translate(
        str.maketrans(
            {
                "\u00ff": "s",
                "\u00fe": "s",
                "\u00f0": "g",
                "\u0153": "u",
                "\u00e5": "a",
                "\u00e4": "a",
                "\u00e3": "a",
                "\u00b1": "i",
                "\u00a7": "c",
            }
        )
    )
    return "".join(ch for ch in metin if ch.isalnum())


def sutun_adi_coz(mevcut_sutunlar: set[str], aday_sutunlar: tuple[str, ...]) -> str | None:
    for sutun in aday_sutunlar:
        if sutun in mevcut_sutunlar:
            return sutun

    token_haritasi: dict[str, str] = {}
    for sutun in sorted(mevcut_sutunlar):
        token = normalize_lookup_token(sutun)
        if token and token not in token_haritasi:
            token_haritasi[token] = sutun

    for aday in aday_sutunlar:
        token = normalize_lookup_token(aday)
        if token and token in token_haritasi:
            return token_haritasi[token]
    return None


def emit_progress(percent: int, phase_code: str, phase_label: str) -> None:
    payload = {
        "percent": max(0, min(100, int(round(percent)))),
        "phase_code": str(phase_code).strip(),
        "phase_label": str(phase_label).strip(),
    }
    print(f"{JOB_PROGRESS_PREFIX}{json.dumps(payload, ensure_ascii=True)}", flush=True)


def sayiya_cevir(deger: Any) -> float | None:
    """Metin veya sayi degerini float'a cevirir; bos degerler icin None doner."""
    if deger is None:
        return None
    if isinstance(deger, (int, float)):
        return float(deger)

    metin = str(deger).strip()
    if not metin:
        return None

    metin = metin.replace(" ", "")

    if "," in metin and "." in metin:
        # Son ayrac ondalik kabul edilir.
        if metin.rfind(",") > metin.rfind("."):
            normalize = metin.replace(".", "").replace(",", ".")
        else:
            normalize = metin.replace(",", "")
    elif "," in metin:
        normalize = metin.replace(",", ".")
    elif metin.count(".") > 1:
        normalize = metin.replace(".", "")
    else:
        normalize = metin

    return float(normalize)


def tam_sayi_metnine_cevir(deger: Any) -> str | None:
    """Degeri tam sayiya cevirip metin olarak dondurur."""
    try:
        sayi = sayiya_cevir(deger)
    except (TypeError, ValueError):
        return None

    if sayi is None:
        return None

    return str(int(sayi))


def urun_agaci_malzeme_bilesen_ayni_kayitlari_sil(
    baglanti: sqlite3.Connection,
) -> int:
    """
    Urun_Agaci tablosunda Malzeme ve Bilesen degeri ayni olan satirlari siler.
    Bu temizlik ETL'in en erken asamasinda calistirilir.
    """
    if not tablo_var_mi(baglanti, "Urun_Agaci"):
        return 0

    sutunlar = tablo_sutunlarini_getir(baglanti, "Urun_Agaci")
    malzeme_sutunu = sutun_adi_coz(sutunlar, ("Malzeme",))
    bilesen_sutunu = sutun_adi_coz(sutunlar, ("BileÅŸen", "Bileşen", "Bilesen"))
    if malzeme_sutunu is None or bilesen_sutunu is None:
        return 0

    baglanti.create_function("tam_sayi_metni", 1, tam_sayi_metnine_cevir)
    imlec = baglanti.cursor()
    kacisli_tablo = sql_kimlik_kacaga_dayanikli("Urun_Agaci")
    kacisli_malzeme = sql_kimlik_kacaga_dayanikli(malzeme_sutunu)
    kacisli_bilesen = sql_kimlik_kacaga_dayanikli(bilesen_sutunu)

    imlec.execute(
        f"""
        DELETE FROM {kacisli_tablo}
        WHERE tam_sayi_metni(CAST({kacisli_malzeme} AS TEXT)) IS NOT NULL
          AND tam_sayi_metni(CAST({kacisli_malzeme} AS TEXT)) =
              tam_sayi_metni(CAST({kacisli_bilesen} AS TEXT))
        """
    )
    silinen = imlec.rowcount
    log_kaydet(
        baglanti,
        "URUN_AGACI_KENDI_ESIT_SIL",
        "Urun_Agaci tablosunda Malzeme=Bilesen olan kayitlar temizlendi.",
        silinen,
    )
    baglanti.commit()
    return silinen


def adt_sutunlarini_ay_sonu_tarihine_cevir(baglanti: sqlite3.Connection) -> list[str]:
    """
    MM.YYYY ADT adindaki sutunlari DD.MM.YYYY formatina cevirir.
    Ornek: 03.2026 ADT -> 31.03.2026
    """
    imlec = baglanti.cursor()
    sutunlar = [satir[1] for satir in imlec.execute(f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})")]

    yeniden_adlandir = []
    for sutun in sutunlar:
        eslesme = AYLIK_ADT_DESENI.match(sutun)
        if not eslesme:
            continue

        ay = int(eslesme.group(1))
        yil = int(eslesme.group(2))
        if ay < 1 or ay > 12:
            continue

        ay_sonu = calendar.monthrange(yil, ay)[1]
        yeni_ad = f"{ay_sonu:02d}.{ay:02d}.{yil}"
        yeniden_adlandir.append((sutun, yeni_ad))

    mevcut = set(sutunlar)
    for eski_ad, yeni_ad in yeniden_adlandir:
        if yeni_ad in mevcut and yeni_ad != eski_ad:
            raise ValueError(f"Sutun adi cakismasi: {eski_ad} -> {yeni_ad}")
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli(eski_ad)} "
            f"TO {sql_kimlik_kacaga_dayanikli(yeni_ad)}"
        )
        mevcut.remove(eski_ad)
        mevcut.add(yeni_ad)

    return [yeni_ad for _, yeni_ad in yeniden_adlandir]


def adt_sutunlarini_bin_ile_carp(
    baglanti: sqlite3.Connection, adt_tarih_sutunlari: list[str]
) -> int:
    """Verilen sutunlardaki sayisal degerleri 1000 ile carpar."""

    def bin_ile_carp(deger: Any) -> Any:
        try:
            sayi = sayiya_cevir(deger)
        except (TypeError, ValueError):
            return deger
        if sayi is None:
            return None
        return sayi * 1000

    baglanti.create_function("bin_ile_carp", 1, bin_ile_carp)
    imlec = baglanti.cursor()

    for sutun in adt_tarih_sutunlari:
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET {sql_kimlik_kacaga_dayanikli(sutun)} = "
            f"bin_ile_carp({sql_kimlik_kacaga_dayanikli(sutun)})"
        )

    return len(adt_tarih_sutunlari)


def tarihli_kolonlari_temizle(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    talepler_x tablosunda tarih formatindaki (DD.MM.YYYY) kolonlari tespit eder
    ve sayisal olmayan degerleri 0'a cevirir.
    
    Returns:
        tuple: (tarihli_kolon_sayisi, temizlenen_kayit_sayisi)
    """
    imlec = baglanti.cursor()
    
    # Tarihli kolonlari tespit et (DD.MM.YYYY formati)
    tarih_deseni = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    
    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]
    
    tarihli_sutunlar = [s for s in sutunlar if tarih_deseni.match(s)]
    kalip_ifadesi = gruplu_metin_ifadesi(set(sutunlar), ("KalÄ±p NumarasÄ±",))
    renk_ifadesi = gruplu_metin_ifadesi(set(sutunlar), ("Renk", "Renk TanÄ±mÄ±"))
    tanim_ifadesi = gruplu_metin_ifadesi(set(sutunlar), ("TanÄ±m",))

    imlec.execute(
        f"""
        INSERT INTO {sql_kimlik_kacaga_dayanikli(TALEPLER_X_TOPLULASTIRMA_LOG_TABLOSU)} (
            Log_Zamani,
            Malzeme,
            Onceki_Kayit_Sayisi,
            Birlesen_Kayit_Sayisi,
            Kalip_Numaralari,
            Renkler,
            Tanimlar
        )
        SELECT
            ?,
            CAST(Malzeme AS TEXT) AS Malzeme,
            COUNT(*) AS Onceki_Kayit_Sayisi,
            COUNT(*) - 1 AS Birlesen_Kayit_Sayisi,
            {kalip_ifadesi} AS Kalip_Numaralari,
            {renk_ifadesi} AS Renkler,
            {tanim_ifadesi} AS Tanimlar
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        GROUP BY Malzeme
        HAVING COUNT(*) > 1
        """,
        (simdi_metni(),),
    )
    
    if not tarihli_sutunlar:
        return 0, 0
    
    # Her tarihli kolonda sayisal olmayan degerleri 0 yap
    temizlenen_toplam = 0
    for sutun in tarihli_sutunlar:
        # NULL degerleri 0 yap
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET {sql_kimlik_kacaga_dayanikli(sutun)} = '0' "
            f"WHERE {sql_kimlik_kacaga_dayanikli(sutun)} IS NULL"
        )
        temizlenen_toplam += imlec.rowcount
        
        # Sayisal olmayan degerleri 0 yap
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET {sql_kimlik_kacaga_dayanikli(sutun)} = '0' "
            f"WHERE {sql_kimlik_kacaga_dayanikli(sutun)} IS NOT NULL "
            f"AND {sql_kimlik_kacaga_dayanikli(sutun)} != '' "
            f"AND typeof({sql_kimlik_kacaga_dayanikli(sutun)}) != 'real' "
            f"AND CAST({sql_kimlik_kacaga_dayanikli(sutun)} AS REAL) = 0 "
            f"AND {sql_kimlik_kacaga_dayanikli(sutun)} != '0'"
        )
        temizlenen_toplam += imlec.rowcount
    
    log_kaydet(baglanti, "TARIHLI_KOLON_TEMIZLE", f"Tarihli kolonlardaki sayisal olmayan degerler temizlendi", temizlenen_toplam)
    
    return len(tarihli_sutunlar), temizlenen_toplam


def stok_tahsis_et(
    baglanti: sqlite3.Connection,
) -> tuple[int, int, dict[str, dict[str, Any]]]:
    """
    TANIM <> 'SADE' olan malzemelerde stokla talepleri karşılar.
    Tarihli kolonları (DD.MM.YYYY) tarihsel sıraya göre sıralar.
    Her malzeme için:
    - Stok miktarını alır
    - Her tarihli kolonu sırayla kontrol eder
    - Stok >= Talep ise: Talep 0 olur, Stok'tan düşülür
    - Stok < Talep ise: Talep = Talep - Stok olur, Stok 0 olur
    Sonuçta talep kolonları ve Başlangıç Stoku kolonu güncellenir.

    Returns:
        tuple: (
            guncellenen_kayit_sayisi,
            islem_goren_kolon_sayisi,
            stok_tahsis_haritasi,
        )
    """
    imlec = baglanti.cursor()

    # Tarihli kolonlari tespit et
    tarih_deseni = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]

    tarihli_sutunlar = [s for s in sutunlar if tarih_deseni.match(s)]

    if not tarihli_sutunlar:
        return 0, 0, {}

    # Tarihli kolonları tarihsel olarak sırala (YYYYMMDD'ye çevirip sırala)
    def tarih_sirala_anahtari(kolon_adi: str) -> int:
        # DD.MM.YYYY -> YYYYMMDD (int olarak karşılaştır)
        gun, ay, yil = kolon_adi.split(".")
        return int(f"{yil}{ay}{gun}")

    tarihli_sutunlar_sirali = sorted(tarihli_sutunlar, key=tarih_sirala_anahtari)

    # SADE olmayan kayıtları al
    imlec.execute(
        f"SELECT Malzeme, \"Başlangıç Stoku\", {', '.join(sql_kimlik_kacaga_dayanikli(s) for s in tarihli_sutunlar_sirali)} "
        f"FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"WHERE Tanım IS NOT NULL AND Tanım <> 'SADE'"
    )
    kayitlar = imlec.fetchall()

    guncellenen = 0
    stok_tahsis_haritasi: dict[str, dict[str, Any]] = {}

    for kayit in kayitlar:
        malzeme = kayit[0]
        orijinal_stok = sayiya_cevir(kayit[1]) or 0
        stok = orijinal_stok
        talepler = list(kayit[2:])  # Tarihli kolon değerleri
        stokla_eriyen: dict[str, float] = {}

        # Her talep için stok tahsisi yap
        for i, talep_deger in enumerate(talepler):
            talep = sayiya_cevir(talep_deger) or 0
            stoktan_karsilanan = min(max(stok, 0), max(talep, 0))
            stokla_eriyen[tarihli_sutunlar_sirali[i]] = stoktan_karsilanan
            talepler[i] = max(0.0, talep - stoktan_karsilanan)
            stok = max(0.0, stok - stoktan_karsilanan)

        stok_tahsis_haritasi[str(malzeme)] = {
            "orijinal_baslangic_stoku": float(orijinal_stok),
            "stokla_eriyen": stokla_eriyen,
        }

        # Güncelleme yap
        guncelleme_degerleri = talepler + [stok, malzeme]
        guncelleme_ifadesi = ", ".join(
            f"{sql_kimlik_kacaga_dayanikli(sutun)} = ?" for sutun in tarihli_sutunlar_sirali
        )
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET {guncelleme_ifadesi}, \"Başlangıç Stoku\" = ? "
            f"WHERE Malzeme = ?",
            guncelleme_degerleri
        )
        guncellenen += imlec.rowcount

    if guncellenen > 0:
        log_kaydet(baglanti, "STOK_TAHSIS", f"Başlangıç Stoku ile talep tahsisi yapıldı: {guncellenen} kayıt", guncellenen)

        # Toplam_Talep kolonunu güncelle (tüm tarihli kolonların toplamı)
        toplam_ifadesi = " + ".join(
            f"COALESCE({sql_kimlik_kacaga_dayanikli(sutun)}, 0)"
            for sutun in tarihli_sutunlar_sirali
        )
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET Toplam_Talep = CAST({toplam_ifadesi} AS INTEGER) "
            f"WHERE Tanım IS NOT NULL AND Tanım <> 'SADE'"
        )
        log_kaydet(baglanti, "TOPLAM_TALEP_GUNCELLE", "Stok tahsisinden sonra Toplam_Talep güncellendi", guncellenen)

    baglanti.commit()
    return guncellenen, len(tarihli_sutunlar_sirali), stok_tahsis_haritasi


def sifir_talepleri_sil(baglanti: sqlite3.Connection) -> int:
    """
    talepler_x tablosunda Toplam_Talep = 0 olan kayıtları siler.

    Returns:
        int: silinen kayıt sayısı
    """
    imlec = baglanti.cursor()

    silme_kosulu = "Toplam_Talep = 0"
    silme_nedeni = "Toplam_Talep stok tahsisi ve aktarimlar sonrasinda 0 oldugu icin kayit silindi."
    if tablo_var_mi(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU):
        snapshot_sutunlar = tablo_sutunlarini_getir(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU)
        snapshot_malzeme_sutunu = sutun_adi_coz(snapshot_sutunlar, ("Malzeme",))
        snapshot_toplam_talep_sutunu = sutun_adi_coz(snapshot_sutunlar, ("Toplam_Talep",))
        if snapshot_malzeme_sutunu and snapshot_toplam_talep_sutunu:
            silme_kosulu = (
                "Toplam_Talep = 0 "
                f"AND CAST(Malzeme AS TEXT) NOT IN ("
                f"SELECT CAST({sql_kimlik_kacaga_dayanikli(snapshot_malzeme_sutunu)} AS TEXT) "
                f"FROM {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)} "
                f"WHERE COALESCE(CAST({sql_kimlik_kacaga_dayanikli(snapshot_toplam_talep_sutunu)} AS REAL), 0) > 0"
                f")"
            )
            silme_nedeni = (
                "Toplam_Talep 0 olanlardan, ham snapshot'ta talebi pozitif olanlar korunarak kayit silindi."
            )

    talepler_x_silinme_logu_kaydet(
        baglanti,
        "SIFIR_TALEP_SIL",
        silme_nedeni,
        silme_kosulu,
    )
    imlec.execute(
        f"DELETE FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"WHERE {silme_kosulu}"
    )
    silinen = imlec.rowcount

    if silinen > 0:
        log_kaydet(baglanti, "SIFIR_TALEP_SIL", f"Toplam_Talep = 0 olan {silinen} kayıt silindi", silinen)

    baglanti.commit()
    return silinen


def sade_kolonlarini_sil(baglanti: sqlite3.Connection) -> int:
    """
    talepler_x tablosundan SADE_Malzemesi ve SADE_Nerede kolonlarını siler.

    Returns:
        int: silinen kolon sayısı
    """
    imlec = baglanti.cursor()

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]

    silinecekler = ["SADE_Malzemesi", "SADE_Nerede", "Tanım"]
    silinen = 0

    for sutun in silinecekler:
        if sutun in sutunlar:
            imlec.execute(
                f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
                f"DROP COLUMN {sql_kimlik_kacaga_dayanikli(sutun)}"
            )
            silinen += 1

    if silinen > 0:
        log_kaydet(baglanti, "SADE_KOLON_SIL", f"SADE kolonları ve Tanım silindi: {silinen} kolon", silinen)

    baglanti.commit()
    return silinen


def sadeye_talep_aktar(
    baglanti: sqlite3.Connection,
    stok_tahsis_haritasi: dict[str, dict[str, Any]] | None = None,
) -> tuple[int, int]:
    """
    TANIM <> 'SADE' olan malzemelerin talep değerlerini,
    SADE_Malzemesi kolonundaki SADE malzeme koduna aktarır.
    SADE olmayan malzemede stokla eritilen talebi görünür tutar.
    Toplam_Talep kolonları güncellenir.

    Returns:
        tuple: (aktarilan_kayit_sayisi, guncellenen_sutun_sayisi)
    """
    imlec = baglanti.cursor()

    # Tarihli kolonlari tespit et
    tarih_deseni = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]

    tarihli_sutunlar = [s for s in sutunlar if tarih_deseni.match(s)]
    baslangic_stoku_sutunu = baslangic_stoku_sutununu_coz(set(sutunlar))

    if not tarihli_sutunlar:
        return 0, 0

    # SADE olmayan ve SADE_Malzemesi belirtilmiş kayıtları al
    imlec.execute(
        f"SELECT Malzeme, SADE_Malzemesi, {', '.join(sql_kimlik_kacaga_dayanikli(s) for s in tarihli_sutunlar)} "
        f"FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"WHERE Tanım IS NOT NULL AND Tanım <> 'SADE' AND SADE_Malzemesi IS NOT NULL AND SADE_Malzemesi != ''"
    )
    kayitlar = imlec.fetchall()

    aktarilan = 0
    sade_kodlari_kumesi = set()  # Kaç farklı SADE koduna aktarıldı
    kolon_toplamlari = {sutun: 0 for sutun in tarihli_sutunlar}  # Her kolon için toplam aktarım
    genel_toplam = 0  # Tüm kolonlardaki toplam aktarım
    
    # Her SADE malzemesi için detaylı istatistik
    sade_bazli_istatistik = {}  # {sade_kodu: {'kayit_sayisi': int, 'toplam': float, 'kolonlar': {sutun: float}}}

    for kayit in kayitlar:
        malzeme_kodu = kayit[0]
        sade_kodu = kayit[1]
        talepler = list(kayit[2:])  # Tarihli kolon değerleri

        # Bu SADE kodunu kümeye ekle
        sade_kodlari_kumesi.add(sade_kodu)
        
        # SADE bazlı istatistik başlat
        if sade_kodu not in sade_bazli_istatistik:
            sade_bazli_istatistik[sade_kodu] = {
                'kayit_sayisi': 0,
                'toplam': 0,
                'kolonlar': {sutun: 0 for sutun in tarihli_sutunlar}
            }
        
        # Bu SADE-olmayan malzemeden aktarım yapıldı, sayacı artır
        sade_bazli_istatistik[sade_kodu]['kayit_sayisi'] += 1

        # SADE malzemesinin mevcut talep değerlerini al
        imlec.execute(
            f"SELECT {', '.join(sql_kimlik_kacaga_dayanikli(s) for s in tarihli_sutunlar)} "
            f"FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"WHERE Malzeme = ?",
            (sade_kodu,)
        )
        sade_mevcut = imlec.fetchone()

        if sade_mevcut:
            # SADE malzemesinin yeni değerlerini hesapla (mevcut + aktarılacak)
            sade_yeni = []
            for i, sade_deger in enumerate(sade_mevcut):
                sade_me = sayiya_cevir(sade_deger) or 0
                aktarilan_deger = sayiya_cevir(talepler[i]) or 0
                
                # Kolon toplamına ekle
                kolon_toplamlari[tarihli_sutunlar[i]] += aktarilan_deger
                genel_toplam += aktarilan_deger
                
                # SADE bazlı istatistiğe ekle
                sade_bazli_istatistik[sade_kodu]['toplam'] += aktarilan_deger
                sade_bazli_istatistik[sade_kodu]['kolonlar'][tarihli_sutunlar[i]] += aktarilan_deger
                
                sade_yeni.append(sade_me + aktarilan_deger)

            # SADE malzemesini güncelle
            guncelleme_ifadesi = ", ".join(
                f"{sql_kimlik_kacaga_dayanikli(sutun)} = ?" for sutun in tarihli_sutunlar
            )
            imlec.execute(
                f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
                f"SET {guncelleme_ifadesi} "
                f"WHERE Malzeme = ?",
                sade_yeni + [sade_kodu]
            )

        stok_tahsis_detayi = (
            stok_tahsis_haritasi.get(str(malzeme_kodu), {})
            if stok_tahsis_haritasi
            else {}
        )
        stokla_eriyen = stok_tahsis_detayi.get("stokla_eriyen", {})
        if not isinstance(stokla_eriyen, dict):
            stokla_eriyen = {}
        orijinal_baslangic_stoku = stok_tahsis_detayi.get("orijinal_baslangic_stoku")

        # SADE olmayan malzemede stokla eritilen kısım görünür kalır.
        non_sade_yeni_talepler = [
            sayiya_cevir(stokla_eriyen.get(sutun)) or 0
            for sutun in tarihli_sutunlar
        ]
        guncelleme_parcalari = [
            f"{sql_kimlik_kacaga_dayanikli(sutun)} = ?"
            for sutun in tarihli_sutunlar
        ]
        guncelleme_degerleri: list[Any] = list(non_sade_yeni_talepler)
        if baslangic_stoku_sutunu and orijinal_baslangic_stoku is not None:
            guncelleme_parcalari.append(
                f"{sql_kimlik_kacaga_dayanikli(baslangic_stoku_sutunu)} = ?"
            )
            guncelleme_degerleri.append(sayiya_cevir(orijinal_baslangic_stoku) or 0)
        guncelleme_ifadesi = ", ".join(guncelleme_parcalari)
        guncelleme_degerleri.append(malzeme_kodu)
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET {guncelleme_ifadesi} "
            f"WHERE Malzeme = ?",
            guncelleme_degerleri,
        )

        aktarilan += 1

    if aktarilan > 0:
        # Özet log
        log_kaydet(baglanti, "SADEYE_TALEP_AKTAR", 
                   f"{aktarilan} SADE-olmayan malzemeden {len(sade_kodlari_kumesi)} farklı SADE malzemesine talep aktarıldı. Toplam: {int(genel_toplam):,}", 
                   aktarilan)

        # Her tarihli kolon için detaylı log
        for sutun, toplam in kolon_toplamlari.items():
            if toplam > 0:
                log_kaydet(baglanti, "SADEYE_TALEP_AKTAR_KOLON", 
                           f"{sutun} kolonuna {int(toplam):,} adet aktarıldı", 
                           aktarilan)

        # En çok talep aktarılan SADE malzemesinden örnek
        if sade_bazli_istatistik:
            en_cok_sade = max(sade_bazli_istatistik.items(), key=lambda x: x[1]['toplam'])
            sade_kodu, istatistik = en_cok_sade
            log_kaydet(baglanti, "SADEYE_TALEP_AKTAR_ORNEK",
                       f"Örnek: SADE Malzeme #{sade_kodu}'na {istatistik['kayit_sayisi']} kayıttan toplam {int(istatistik['toplam']):,} adet aktarıldı",
                       int(istatistik['toplam']))

        # Toplam_Talep kolonunu tüm kayıtlar için güncelle
        toplam_ifadesi = " + ".join(
            f"COALESCE({sql_kimlik_kacaga_dayanikli(sutun)}, 0)"
            for sutun in tarihli_sutunlar
        )
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET Toplam_Talep = CAST({toplam_ifadesi} AS INTEGER)"
        )
        log_kaydet(baglanti, "TOPLAM_TALEP_GUNCELLE", "Talep aktarımından sonra Toplam_Talep güncellendi", aktarilan)

    baglanti.commit()
    return aktarilan, len(tarihli_sutunlar)


def talepler_x_toplula(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    talepler_x tablosunda ayni Malzeme'ye sahip kayitlari birlestirir.
    Tarihli kolonlardaki degerleri toplar, diğer kolonlarda ilk degeri saklar.
    Toplam_Talep kolonu ekler ve tüm tarihli kolonların toplamını hesaplar.
    Stok kolonu ekler ve stok_x tablosundan stok değerlerini getirir.

    Returns:
        tuple: (eski_kayit_sayisi, yeni_kayit_sayisi)
    """
    imlec = baglanti.cursor()

    # Eski kayıt sayısını al
    imlec.execute(f"SELECT COUNT(*) FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}")
    eski_sayisi = imlec.fetchone()[0]

    # Tarihli kolonlari tespit et
    tarih_deseni = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]

    tarihli_sutunlar = [s for s in sutunlar if tarih_deseni.match(s)]
    
    # Gruplama için tarihli kolonların toplam ifadesi
    toplam_ifadeleri = ", ".join(
        f"SUM(COALESCE({sql_kimlik_kacaga_dayanikli(sutun)}, 0)) as {sql_kimlik_kacaga_dayanikli(sutun)}"
        for sutun in tarihli_sutunlar
    )

    # Diğer kolonlar (Malzeme hariç): temsilci satırdan seç.
    # Aynı malzeme birden fazla satırda geçiyorsa MIN/MAX gibi leksik seçim
    # kalıp/renk karışmasına yol açabiliyor. Bu nedenle en yüksek Toplam_Talep
    # (eşitlikte en küçük rowid) satırı temsilci kabul edilir.
    diger_sutunlar = [s for s in sutunlar if not tarih_deseni.match(s) and s != 'Malzeme']
    talep_skor_ifadesi = (
        "COALESCE(Toplam_Talep, 0)"
        if "Toplam_Talep" in sutunlar
        else (
            " + ".join(
                f"COALESCE({sql_kimlik_kacaga_dayanikli(sutun)}, 0)"
                for sutun in tarihli_sutunlar
            )
            if tarihli_sutunlar
            else "0"
        )
    )
    diger_ifadeler = ", ".join(
        (
            f"MAX(CASE WHEN __rn = 1 THEN {sql_kimlik_kacaga_dayanikli(sutun)} END) "
            f"as {sql_kimlik_kacaga_dayanikli(sutun)}"
        )
        for sutun in diger_sutunlar
    )

    # Geçici tablo oluştur
    imlec.execute("DROP TABLE IF EXISTS talepler_x_temp")
    secim_parcalari = ["Malzeme"]
    if diger_ifadeler:
        secim_parcalari.append(diger_ifadeler)
    if toplam_ifadeleri:
        secim_parcalari.append(toplam_ifadeleri)
    secim_sql = ",\n            ".join(secim_parcalari)
    imlec.execute(
        f"""CREATE TABLE talepler_x_temp AS
        WITH sirali AS (
            SELECT
                *,
                ROW_NUMBER() OVER (
                    PARTITION BY Malzeme
                    ORDER BY {talep_skor_ifadesi} DESC, rowid ASC
                ) AS __rn
            FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        )
        SELECT 
            {secim_sql}
        FROM sirali
        GROUP BY Malzeme"""
    )
    
    # Toplam_Talep kolonu ekle
    imlec.execute(
        f"ALTER TABLE talepler_x_temp ADD COLUMN Toplam_Talep INTEGER"
    )
    
    # Toplam_Talep'i hesapla (tüm tarihli kolonların toplamı)
    toplam_kolonlar = " + ".join(
        f"COALESCE({sql_kimlik_kacaga_dayanikli(sutun)}, 0)"
        for sutun in tarihli_sutunlar
    )
    imlec.execute(
        f"UPDATE talepler_x_temp SET Toplam_Talep = CAST({toplam_kolonlar} AS INTEGER)"
    )
    
    # Stok kolonu ekle
    imlec.execute(
        f"ALTER TABLE talepler_x_temp ADD COLUMN Stok INTEGER"
    )
    
    # Stok değerlerini stok_x'ten getir (eşleşmeyenler 0 olur)
    imlec.execute(
        f"""UPDATE talepler_x_temp 
        SET Stok = COALESCE((SELECT s.\"Toplam Stok\" FROM stok_x s 
                    WHERE s.Malzeme = talepler_x_temp.Malzeme), 0)"""
    )
    
    # Eski tabloyu sil ve yenisini oluştur
    imlec.execute(f"DROP TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}")
    imlec.execute(
        f"ALTER TABLE talepler_x_temp RENAME TO {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}"
    )
    
    # Yeni kayıt sayısını al
    imlec.execute(f"SELECT COUNT(*) FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}")
    yeni_sayisi = imlec.fetchone()[0]
    
    log_kaydet(baglanti, "TOPLULA", f"talepler_x toplulaştırıldı: {eski_sayisi} -> {yeni_sayisi} kayıt", yeni_sayisi)
    log_kaydet(baglanti, "TOPLAM_TALEP", "Toplam_Talep kolonu hesaplandı", yeni_sayisi)
    log_kaydet(baglanti, "STOK_GETIR", "stok_x'ten Başlangıç Stoku kolonu dolduruldu", yeni_sayisi)

    # Stok kolonunun adını Başlangıç Stoku olarak değiştir
    imlec.execute(
        f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('Stok')} "
        f"TO {sql_kimlik_kacaga_dayanikli('Başlangıç Stoku')}"
    )
    log_kaydet(baglanti, "STOK_RENAME", "Stok kolonu Başlangıç Stoku olarak yeniden adlandırıldı", yeni_sayisi)

    baglanti.commit()
    return eski_sayisi, yeni_sayisi


def kalip_bilgisi_kontrol_et(baglanti: sqlite3.Connection) -> tuple[int, int, int, int]:
    """
    talepler_x tablosundaki kalıp numaralarını Kalıp_Bilgisi sayfasında tarar.
    Hiç eşleşme bulamayan kayıtları Kalıp_Bilgisi_Eksik_X tablosuna ekler.

    Returns:
        tuple: (toplam_kalip_sayisi, eslesen_kalip_sayisi, eksik_kalip_sayisi, eksik_kayit_sayisi)
    """
    imlec = baglanti.cursor()

    # Kalıp Numarası olan tüm distinct kalıpları al
    imlec.execute(
        f"SELECT DISTINCT \"Kalıp Numarası\" FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"WHERE \"Kalıp Numarası\" IS NOT NULL"
    )
    kalip_numaralari = imlec.fetchall()
    toplam_kalip_sayisi = len(kalip_numaralari)

    eslesen_kalip_sayisi = 0
    eksik_kalip_numaralari = []

    for (kalip_no,) in kalip_numaralari:
        # Kalıp_Bilgisi tablosunda eşleşme var mı kontrol et
        imlec.execute(
            "SELECT 1 FROM Kalıp_Bilgisi WHERE \"Kalıp Kodu\" = ? LIMIT 1",
            (kalip_no,)
        )
        if imlec.fetchone():
            eslesen_kalip_sayisi += 1
        else:
            eksik_kalip_numaralari.append(kalip_no)

    eksik_kalip_sayisi = len(eksik_kalip_numaralari)

    # Eşleşmeyen kalıp numaralarını Kalıp_Bilgisi_Eksik_X tablosuna ekle
    imlec.execute("DROP TABLE IF EXISTS Kalıp_Bilgisi_Eksik_X")
    if eksik_kalip_numaralari:
        yer_tutucular = ','.join('?' * len(eksik_kalip_numaralari))
        imlec.execute(
            f"""CREATE TABLE Kalıp_Bilgisi_Eksik_X AS
            SELECT Malzeme, "Kalıp Numarası", "Renk Tanımı", "Başlangıç Stoku", Toplam_Talep
            FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
            WHERE "Kalıp Numarası" IS NOT NULL
            AND "Kalıp Numarası" IN ({yer_tutucular})""",
            eksik_kalip_numaralari
        )
    else:
        imlec.execute(
            f"""CREATE TABLE Kalıp_Bilgisi_Eksik_X AS
            SELECT Malzeme, "Kalıp Numarası", "Renk Tanımı", "Başlangıç Stoku", Toplam_Talep
            FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
            WHERE 0 = 1"""
        )

    # Eksik kayıt sayısını al
    imlec.execute(f"SELECT COUNT(*) FROM Kalıp_Bilgisi_Eksik_X")
    eksik_kayit_sayisi = imlec.fetchone()[0]

    if eksik_kayit_sayisi > 0:
        log_kaydet(baglanti, "KALIP_BILGISI_EKSIK", f"Kalıp_Bilgisi'nde eşleşme bulunamayan {eksik_kayit_sayisi} kayıt Kalıp_Bilgisi_Eksik_X'e eklendi", eksik_kayit_sayisi)

        # Eşleşme bulunamayan kayıtlarda stoklu malzemeleri koru.
        mevcut_sutunlar = tablo_sutunlarini_getir(baglanti, HEDEF_TABLO)
        kalip_sutunu = sutun_adi_coz(mevcut_sutunlar, ("Kalıp Numarası", "KalÄ±p NumarasÄ±"))
        if kalip_sutunu is None:
            kalip_sutunu = "Kalıp Numarası"
        baslangic_stoku_sutunu = baslangic_stoku_sutununu_coz(mevcut_sutunlar)
        stok_kosulu = (
            f"COALESCE(CAST({sql_kimlik_kacaga_dayanikli(baslangic_stoku_sutunu)} AS REAL), 0) <= 0"
            if baslangic_stoku_sutunu
            else "1 = 1"
        )
        yer_tutucular = ",".join("?" * len(eksik_kalip_numaralari))
        silme_kosulu = (
            f"{sql_kimlik_kacaga_dayanikli(kalip_sutunu)} IS NOT NULL "
            f"AND {sql_kimlik_kacaga_dayanikli(kalip_sutunu)} IN ({yer_tutucular}) "
            f"AND {stok_kosulu}"
        )
        talepler_x_silinme_logu_kaydet(
            baglanti,
            "KALIP_BILGISI_SIL",
            "Kalip numarasi Kalip_Bilgisi tablosunda bulunamadigi icin kayit silindi.",
            silme_kosulu,
            tuple(eksik_kalip_numaralari),
        )
        imlec.execute(
            f"DELETE FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"WHERE {silme_kosulu}",
            eksik_kalip_numaralari
        )
        silinen = imlec.rowcount
        log_kaydet(baglanti, "KALIP_BILGISI_SIL", f"Kalıp_Bilgisi'nde eşleşme bulunamayan {silinen} kayıt talepler_x'ten silindi", silinen)

    log_kaydet(baglanti, "KALIP_BILGISI_KONTROL", 
               f"Toplam {toplam_kalip_sayisi} distinct kalıp: {eslesen_kalip_sayisi} eşleşen, {eksik_kalip_sayisi} eşleşmeyen ({eksik_kayit_sayisi} kayıt)", 
               toplam_kalip_sayisi)

    baglanti.commit()
    return toplam_kalip_sayisi, eslesen_kalip_sayisi, eksik_kalip_sayisi, eksik_kayit_sayisi


def gramaj_kolonu_ekle(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    talepler_x tablosuna Gramaj kolonu ekler ve Parametre tablosundan
    Kalıp Numarası ile eşleşen ilk Gramaj değerini getirir.
    Gramaj bulunamayan kayıtları Parametre_Bilgisi_Eksik_X tablosuna ekler.

    Returns:
        tuple: (doldurulan_sayisi, eksik_sayisi)
    """
    imlec = baglanti.cursor()

    # Gramaj kolonu ekle
    imlec.execute(
        f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"ADD COLUMN Gramaj REAL"
    )

    # Parametre tablosundan Kalıp Numarası ile eşleşen ilk Gramaj değerini al
    imlec.execute(
        f"""UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        SET Gramaj = (
            SELECT p."Fiili Gramaj" FROM Parametre p
            WHERE p.Kalıp = {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}."Kalıp Numarası"
            LIMIT 1
        )
        WHERE "Kalıp Numarası" IS NOT NULL"""
    )
    doldurulan = imlec.rowcount

    # Gramaj bulunamayan kayıtları Parametre_Bilgisi_Eksik_X tablosuna ekle
    imlec.execute("DROP TABLE IF EXISTS Parametre_Bilgisi_Eksik_X")
    imlec.execute(
        f"""CREATE TABLE Parametre_Bilgisi_Eksik_X AS
        SELECT Malzeme, "Kalıp Numarası", "Renk Tanımı", "Başlangıç Stoku", Toplam_Talep, Gramaj
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        WHERE Gramaj IS NULL"""
    )

    # Eksik kayıt sayısını al
    imlec.execute(f"SELECT COUNT(*) FROM Parametre_Bilgisi_Eksik_X")
    eksik = imlec.fetchone()[0]

    if eksik > 0:
        log_kaydet(baglanti, "PARAMETRE_BILGISI_EKSIK", f"Parametre bilgisi eksik {eksik} kayıt Parametre_Bilgisi_Eksik_X'e eklendi", eksik)

        # Gramaj bulunamayan kayıtlarda stoklu malzemeleri koru.
        mevcut_sutunlar = tablo_sutunlarini_getir(baglanti, HEDEF_TABLO)
        baslangic_stoku_sutunu = baslangic_stoku_sutununu_coz(mevcut_sutunlar)
        stok_kosulu = (
            f"COALESCE(CAST({sql_kimlik_kacaga_dayanikli(baslangic_stoku_sutunu)} AS REAL), 0) <= 0"
            if baslangic_stoku_sutunu
            else "1 = 1"
        )
        silme_kosulu = f"Gramaj IS NULL AND {stok_kosulu}"
        talepler_x_silinme_logu_kaydet(
            baglanti,
            "GRAMAJSIZ_SIL",
            "Parametre tablosundan gramaj bilgisi alinamadigi icin kayit silindi.",
            silme_kosulu,
        )
        imlec.execute(
            f"DELETE FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"WHERE {silme_kosulu}"
        )
        silinen = imlec.rowcount
        log_kaydet(baglanti, "GRAMAJSIZ_SIL", f"Gramaj bulunamayan {silinen} kayıt talepler_x'ten silindi", silinen)

    log_kaydet(baglanti, "GRAMAJ_EKLE", f"Gramaj kolonu eklendi: {doldurulan} kayıt dolduruldu, {eksik} kayıt eksik", doldurulan)

    baglanti.commit()
    return doldurulan, eksik


def ayarlar_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Kalip_Bilgisi ve Parametre tablolarinin ortak Fabrika-Kalip-Hat No
    kayitlarindan, sadece talepler_x'te bulunan kalip numaralari icin
    Ayarlar_X tablosunu olusturur.
    """
    baglanti.create_function("tam_sayi_metni", 1, tam_sayi_metnine_cevir)
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Ayarlar_X")
    imlec.execute(
        """
        CREATE TABLE Ayarlar_X AS
        WITH kalip_bilgisi_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                tam_sayi_metni("Kalıp Kodu") AS Kalıp,
                tam_sayi_metni("Hat No") AS "Hat No",
                Max_Kol_Sayisi
            FROM Kalıp_Bilgisi
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND tam_sayi_metni("Kalıp Kodu") IS NOT NULL
              AND tam_sayi_metni("Hat No") IS NOT NULL
        ),
        parametre_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                tam_sayi_metni(Kalıp) AS Kalıp,
                tam_sayi_metni("Hat No") AS "Hat No",
                "Fiili Hız dam/dak",
                "1.Gün Sonrası Verimi"
            FROM Parametre
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND tam_sayi_metni(Kalıp) IS NOT NULL
              AND tam_sayi_metni("Hat No") IS NOT NULL
        ),
        hat_bilgisi_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                tam_sayi_metni("Hat No") AS "Hat No",
                Firin,
                Hatlar
            FROM Hat_Bilgisi
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND tam_sayi_metni("Hat No") IS NOT NULL
        ),
        talepler_kaliplari AS (
            SELECT DISTINCT
                tam_sayi_metni("Kalıp Numarası") AS Kalıp
            FROM talepler_x
            WHERE tam_sayi_metni("Kalıp Numarası") IS NOT NULL
        )
        SELECT DISTINCT
            k.Kalıp,
            k.Fabrika,
            h.Firin,
            h.Hatlar,
            k."Hat No",
            k.Max_Kol_Sayisi AS Calisabilecegi_Max_Kol_Sayisi,
            p."Fiili Hız dam/dak" AS Fiili_Hiz_Dam_Dak,
            p."1.Gün Sonrası Verimi" AS Birinci_Gun_Sonrasi_Verimi
        FROM kalip_bilgisi_norm k
        INNER JOIN parametre_norm p
            ON k.Fabrika = p.Fabrika
           AND k.Kalıp = p.Kalıp
           AND k."Hat No" = p."Hat No"
        INNER JOIN hat_bilgisi_norm h
            ON k.Fabrika = h.Fabrika
           AND k."Hat No" = h."Hat No"
        INNER JOIN talepler_kaliplari t
            ON k.Kalıp = t.Kalıp
        ORDER BY k.Fabrika, k.Kalıp, k."Hat No"
        """
    )

    imlec.execute("SELECT COUNT(*) FROM Ayarlar_X")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "AYARLAR_X_OLUSTUR",
        f"Ayarlar_X tablosu {kayit_sayisi} ortak kayit ile olusturuldu",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi


def kapasiteler_x_olustur(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    Kapasiteler tablosunu Kapasiteler_x olarak yeniden duzenler:
    - Uy -> Fabrika
    - Hat_Bilgisi'nden Firin bilgisini ekler
    - Kolon sirasini Fabrika, Firin, Hat olarak duzenler
    """
    baglanti.create_function("tam_sayi_metni", 1, tam_sayi_metnine_cevir)
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Kapasiteler_x")
    imlec.execute(
        """
        CREATE TABLE Kapasiteler_x AS
        WITH kapasiteler_norm AS (
            SELECT
                tam_sayi_metni(Uy) AS Fabrika,
                TRIM(CAST(Hat AS TEXT)) AS Hat,
                Renk,
                Gunluk_Kapasite
            FROM Kapasiteler
            WHERE tam_sayi_metni(Uy) IS NOT NULL
              AND TRIM(CAST(Hat AS TEXT)) != ''
        ),
        hat_bilgisi_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                TRIM(CAST(Hatlar AS TEXT)) AS Hat,
                Firin
            FROM Hat_Bilgisi
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND TRIM(CAST(Hatlar AS TEXT)) != ''
        )
        SELECT
            k.Fabrika,
            h.Firin,
            k.Hat,
            k.Renk,
            k.Gunluk_Kapasite AS "Günlük_Üretim_Tonajı"
        FROM kapasiteler_norm k
        LEFT JOIN hat_bilgisi_norm h
            ON k.Fabrika = h.Fabrika
           AND k.Hat = h.Hat
        ORDER BY k.Fabrika, h.Firin, k.Hat
        """
    )

    imlec.execute("SELECT COUNT(*) FROM Kapasiteler_x")
    kayit_sayisi = imlec.fetchone()[0]
    imlec.execute("SELECT COUNT(*) FROM Kapasiteler_x WHERE Firin IS NULL")
    firin_eksik_sayisi = imlec.fetchone()[0]

    log_kaydet(
        baglanti,
        "KAPASITELER_X_OLUSTUR",
        f"Kapasiteler_x tablosu olusturuldu: {kayit_sayisi} kayit, {firin_eksik_sayisi} kayitta Firin eksik",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi, firin_eksik_sayisi


def damla_sayisi_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Hat_Bilgisi tablosundan zaman bagli Damla_Sayisi_x tablosunu olusturur.
    """
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Damla_Sayisi_x")
    imlec.execute(
        """
        CREATE TABLE Damla_Sayisi_x AS
        SELECT
            Fabrika,
            Firin,
            Hatlar,
            Damla,
            Baslangic_Tarihi,
            Bitis_Tarihi
        FROM Hat_Bilgisi
        ORDER BY Fabrika, Firin, Hatlar, Baslangic_Tarihi, Bitis_Tarihi
        """
    )

    imlec.execute("SELECT COUNT(*) FROM Damla_Sayisi_x")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "DAMLA_SAYISI_X_OLUSTUR",
        f"Damla_Sayisi_x tablosu {kayit_sayisi} kayit ile olusturuldu",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi


def playground_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Hat_Bilgisi tablosundan Fabrika-Firin-Hatlar kombinasyonlarini tekil olarak
    Playground_x tablosuna yazar.
    """
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Playground_x")
    imlec.execute(
        """
        CREATE TABLE Playground_x AS
        SELECT DISTINCT
            Fabrika,
            Firin,
            Hatlar
        FROM Hat_Bilgisi
        ORDER BY Fabrika, Firin, Hatlar
        """
    )

    imlec.execute("SELECT COUNT(*) FROM Playground_x")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "PLAYGROUND_X_OLUSTUR",
        f"Playground_x tablosu {kayit_sayisi} tekil kayit ile olusturuldu",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi


def aylik_teorik_kapasite_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Renk_Plani_X, Hat_Bilgisi ve Kapasiteler_x tablolarini kullanarak
    aylik hat bazli teorik maksimum tonaj tablosunu olusturur.
    """
    def guvenli_sayiya_cevir(deger: Any) -> float | None:
        try:
            return sayiya_cevir(deger)
        except (TypeError, ValueError):
            return None

    baglanti.create_function("tam_sayi_metni", 1, tam_sayi_metnine_cevir)
    baglanti.create_function("sayiya_cevir", 1, guvenli_sayiya_cevir)
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Aylik_Teorik_Kapasite_X")
    imlec.execute(
        """
        CREATE TABLE Aylik_Teorik_Kapasite_X AS
        WITH RECURSIVE
        renk_plani_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                TRIM(CAST(Firin AS TEXT)) AS Firin,
                DATE("Başlangıç Tarihi") AS Baslangic_Tarihi,
                DATE("Bitiş Tarihi") AS Bitis_Tarihi,
                TRIM(CAST(Renk AS TEXT)) AS Renk
            FROM Renk_Plani_X
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND TRIM(CAST(Firin AS TEXT)) != ''
              AND TRIM(CAST(Renk AS TEXT)) != ''
              AND DATE("Başlangıç Tarihi") IS NOT NULL
              AND DATE("Bitiş Tarihi") IS NOT NULL
              AND DATE("Bitiş Tarihi") >= DATE("Başlangıç Tarihi")
        ),
        renk_aylari AS (
            SELECT
                Fabrika,
                Firin,
                Renk,
                Baslangic_Tarihi,
                Bitis_Tarihi,
                DATE(Baslangic_Tarihi, 'start of month') AS Ay_Basi
            FROM renk_plani_norm

            UNION ALL

            SELECT
                Fabrika,
                Firin,
                Renk,
                Baslangic_Tarihi,
                Bitis_Tarihi,
                DATE(Ay_Basi, '+1 month') AS Ay_Basi
            FROM renk_aylari
            WHERE DATE(Ay_Basi, '+1 month') <= DATE(Bitis_Tarihi, 'start of month')
        ),
        renk_ay_segmentleri AS (
            SELECT
                Fabrika,
                Firin,
                Renk,
                STRFTIME('%Y-%m', Ay_Basi) AS Ay,
                CASE
                    WHEN Baslangic_Tarihi > Ay_Basi THEN Baslangic_Tarihi
                    ELSE Ay_Basi
                END AS Segment_Baslangic,
                CASE
                    WHEN Bitis_Tarihi < DATE(Ay_Basi, '+1 month', '-1 day') THEN Bitis_Tarihi
                    ELSE DATE(Ay_Basi, '+1 month', '-1 day')
                END AS Segment_Bitis
            FROM renk_aylari
        ),
        hat_bilgisi_norm AS (
            SELECT DISTINCT
                tam_sayi_metni(Fabrika) AS Fabrika,
                TRIM(CAST(Firin AS TEXT)) AS Firin,
                TRIM(CAST(Hatlar AS TEXT)) AS Hat,
                DATE(Baslangic_Tarihi) AS Hat_Baslangic_Tarihi,
                DATE(Bitis_Tarihi) AS Hat_Bitis_Tarihi
            FROM Hat_Bilgisi
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND TRIM(CAST(Firin AS TEXT)) != ''
              AND TRIM(CAST(Hatlar AS TEXT)) != ''
              AND DATE(Baslangic_Tarihi) IS NOT NULL
              AND DATE(Bitis_Tarihi) IS NOT NULL
              AND DATE(Bitis_Tarihi) >= DATE(Baslangic_Tarihi)
        ),
        hat_ay_segmentleri AS (
            SELECT
                r.Ay,
                r.Fabrika,
                r.Firin,
                h.Hat,
                r.Renk,
                CASE
                    WHEN r.Segment_Baslangic > h.Hat_Baslangic_Tarihi THEN r.Segment_Baslangic
                    ELSE h.Hat_Baslangic_Tarihi
                END AS Segment_Baslangic,
                CASE
                    WHEN r.Segment_Bitis < h.Hat_Bitis_Tarihi THEN r.Segment_Bitis
                    ELSE h.Hat_Bitis_Tarihi
                END AS Segment_Bitis
            FROM renk_ay_segmentleri r
            INNER JOIN hat_bilgisi_norm h
                ON r.Fabrika = h.Fabrika
               AND r.Firin = h.Firin
               AND h.Hat_Bitis_Tarihi >= r.Segment_Baslangic
               AND h.Hat_Baslangic_Tarihi <= r.Segment_Bitis
        ),
        kapasiteler_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                TRIM(CAST(Firin AS TEXT)) AS Firin,
                TRIM(CAST(Hat AS TEXT)) AS Hat,
                TRIM(CAST(Renk AS TEXT)) AS Renk,
                MAX(sayiya_cevir("Günlük_Üretim_Tonajı")) AS Gunluk_Uretim_Tonaji
            FROM Kapasiteler_x
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
              AND TRIM(CAST(Firin AS TEXT)) != ''
              AND TRIM(CAST(Hat AS TEXT)) != ''
              AND TRIM(CAST(Renk AS TEXT)) != ''
            GROUP BY
                tam_sayi_metni(Fabrika),
                TRIM(CAST(Firin AS TEXT)),
                TRIM(CAST(Hat AS TEXT)),
                TRIM(CAST(Renk AS TEXT))
        )
        SELECT
            h.Ay,
            h.Fabrika,
            h.Firin,
            h.Hat,
            SUM(CAST(JULIANDAY(h.Segment_Bitis) - JULIANDAY(h.Segment_Baslangic) + 1 AS INTEGER)) AS Toplam_Gun,
            SUM(
                CASE
                    WHEN h.Renk = 'HAT KAPALI ÇALIŞILMAZ'
                    THEN CAST(JULIANDAY(h.Segment_Bitis) - JULIANDAY(h.Segment_Baslangic) + 1 AS INTEGER)
                    ELSE 0
                END
            ) AS Kapali_Gun,
            SUM(
                CASE
                    WHEN h.Renk != 'HAT KAPALI ÇALIŞILMAZ'
                    THEN CAST(JULIANDAY(h.Segment_Bitis) - JULIANDAY(h.Segment_Baslangic) + 1 AS INTEGER)
                    ELSE 0
                END
            ) AS Uretim_Gunu,
            ROUND(
                SUM(
                    CASE
                        WHEN h.Renk = 'HAT KAPALI ÇALIŞILMAZ' THEN 0
                        ELSE (JULIANDAY(h.Segment_Bitis) - JULIANDAY(h.Segment_Baslangic) + 1)
                             * COALESCE(k.Gunluk_Uretim_Tonaji, 0)
                    END
                ),
                3
            ) AS Teorik_Max_Ton,
            SUM(
                CASE
                    WHEN h.Renk != 'HAT KAPALI ÇALIŞILMAZ' AND k.Gunluk_Uretim_Tonaji IS NULL
                    THEN CAST(JULIANDAY(h.Segment_Bitis) - JULIANDAY(h.Segment_Baslangic) + 1 AS INTEGER)
                    ELSE 0
                END
            ) AS Kapasite_Eksik_Gun
        FROM hat_ay_segmentleri h
        LEFT JOIN kapasiteler_norm k
            ON h.Fabrika = k.Fabrika
           AND h.Firin = k.Firin
           AND h.Hat = k.Hat
           AND h.Renk = k.Renk
        GROUP BY
            h.Ay,
            h.Fabrika,
            h.Firin,
            h.Hat
        ORDER BY
            h.Ay,
            h.Fabrika,
            h.Firin,
            h.Hat
        """
    )

    imlec.execute("SELECT COUNT(*) FROM Aylik_Teorik_Kapasite_X")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "AYLIK_TEORIK_KAPASITE_X_OLUSTUR",
        f"Aylik_Teorik_Kapasite_X tablosu {kayit_sayisi} kayit ile olusturuldu",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi


def hatali_kol_sayilari_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Ayarlar_X tablosundaki calisabilecegi max kol sayisini Hat_Bilgisi'ndeki
    Kol Sayisi ile karsilastirir. Ayarlar_X degeri daha buyukse kaydi
    Hatali_Kol_Sayilari_X tablosuna ekler.
    """
    baglanti.create_function("tam_sayi_metni", 1, tam_sayi_metnine_cevir)
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Hatali_Kol_Sayilari_X")
    imlec.execute(
        """
        CREATE TABLE Hatali_Kol_Sayilari_X AS
        WITH ayarlar_norm AS (
            SELECT
                Kalıp,
                tam_sayi_metni(Fabrika) AS Fabrika,
                TRIM(CAST(Firin AS TEXT)) AS Firin,
                TRIM(CAST(Hatlar AS TEXT)) AS Hatlar,
                "Hat No",
                Calisabilecegi_Max_Kol_Sayisi,
                Fiili_Hiz_Dam_Dak,
                Birinci_Gun_Sonrasi_Verimi
            FROM Ayarlar_X
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
        ),
        hat_bilgisi_norm AS (
            SELECT
                tam_sayi_metni(Fabrika) AS Fabrika,
                TRIM(CAST(Firin AS TEXT)) AS Firin,
                TRIM(CAST(Hatlar AS TEXT)) AS Hatlar,
                "Hat No",
                "Kol Sayısı" AS Hat_Bilgisindeki_Kol_Sayisi
            FROM Hat_Bilgisi
            WHERE tam_sayi_metni(Fabrika) IS NOT NULL
        )
        SELECT
            a.Kalıp,
            a.Fabrika,
            a.Firin,
            a.Hatlar,
            a."Hat No",
            a.Calisabilecegi_Max_Kol_Sayisi,
            h.Hat_Bilgisindeki_Kol_Sayisi,
            a.Fiili_Hiz_Dam_Dak,
            a.Birinci_Gun_Sonrasi_Verimi
        FROM ayarlar_norm a
        INNER JOIN hat_bilgisi_norm h
            ON a.Fabrika = h.Fabrika
           AND a.Firin = h.Firin
           AND a.Hatlar = h.Hatlar
        WHERE CAST(a.Calisabilecegi_Max_Kol_Sayisi AS INTEGER) >
              CAST(h.Hat_Bilgisindeki_Kol_Sayisi AS INTEGER)
        ORDER BY a.Fabrika, a.Firin, a.Hatlar, a.Kalıp
        """
    )

    imlec.execute("SELECT COUNT(*) FROM Hatali_Kol_Sayilari_X")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "HATALI_KOL_SAYILARI_X_OLUSTUR",
        f"Hatali_Kol_Sayilari_X tablosu {kayit_sayisi} kayit ile olusturuldu",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi


def hatalar_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Farkli hata tablolarini tek bir Hatalar_X tablosunda toplar ve
    kaynak hata tablolarini siler.
    """
    imlec = baglanti.cursor()

    imlec.execute("DROP TABLE IF EXISTS Hatalar_X")
    imlec.execute(
        """
        CREATE TABLE Hatalar_X AS
        SELECT *
        FROM (
            SELECT
                'Kalip_Bilgisi_Eksik' AS Hata_Tipi,
                'Kalıp numarası Kalıp_Bilgisi tablosunda bulunamadı; kayıt talepler_x tablosundan çıkarıldı.' AS Hata_Aciklamasi,
                'Excel: Kalıp_Bilgisi sayfasına ilgili Kalıp Numarası için satır ekleyin; en az Kalıp Kodu, Fabrika, Hat No ve Max_Kol_Sayisi bilgilerini girin.' AS Excel_Duzeltme,
                Malzeme,
                "Kalıp Numarası" AS Kalıp,
                'Renk=' || COALESCE(CAST("Renk Tanımı" AS TEXT), '') ||
                ' | Baslangic_Stoku=' || COALESCE(CAST("Başlangıç Stoku" AS TEXT), '') ||
                ' | Toplam_Talep=' || COALESCE(CAST(Toplam_Talep AS TEXT), '') AS Detay
            FROM "Kalıp_Bilgisi_Eksik_X"

            UNION ALL

            SELECT
                'Parametre_Bilgisi_Eksik' AS Hata_Tipi,
                'Kalıp numarası Parametre tablosunda bulundu ama Gramaj bilgisi alınamadı; kayıt talepler_x tablosundan çıkarıldı.' AS Hata_Aciklamasi,
                'Excel: Parametre sayfasında ilgili Kalıp için satırı tamamlayın; en az Fabrika, Kalıp, Hat No, Fiili Gramaj, Fiili Hız dam/dak ve 1.Gün Sonrası Verimi bilgilerini kontrol edin.' AS Excel_Duzeltme,
                Malzeme,
                "Kalıp Numarası" AS Kalıp,
                'Renk=' || COALESCE(CAST("Renk Tanımı" AS TEXT), '') ||
                ' | Baslangic_Stoku=' || COALESCE(CAST("Başlangıç Stoku" AS TEXT), '') ||
                ' | Toplam_Talep=' || COALESCE(CAST(Toplam_Talep AS TEXT), '') ||
                ' | Gramaj=' || COALESCE(CAST(Gramaj AS TEXT), '') AS Detay
            FROM "Parametre_Bilgisi_Eksik_X"

            UNION ALL

            SELECT
                'Sadesi_Bulunamayan' AS Hata_Tipi,
                'Ürün ağacında SADE karşılığı bulunamadı; kayıt talepler_x tablosundan çıkarıldı.' AS Hata_Aciklamasi,
                'Excel: önce Urun_Agaci sayfasında bu Malzeme için SADE''ye giden bileşen ilişkisini ekleyin; gerekiyorsa Malzeme_Data sayfasında ilgili SADE malzemeyi Hiyerarşi=SADE olarak tanımlayın.' AS Excel_Duzeltme,
                Malzeme,
                NULL AS Kalıp,
                'Tanim=' || COALESCE(CAST(Tanım AS TEXT), '') ||
                ' | SADE_Malzemesi=' || COALESCE(CAST(SADE_Malzemesi AS TEXT), '') AS Detay
            FROM "Sadesi_Bulunamayanlar_X"

            UNION ALL

            SELECT
                'Hatali_Kol_Sayisi' AS Hata_Tipi,
                'Kalıp verilerindeki max kol sayısı, Hat_Bilgisi tablosundaki kol sayısını aşıyor.' AS Hata_Aciklamasi,
                'Excel: öncelikle Hat_Bilgisi sayfasında ilgili Fabrika-Firin-Hatlar için Kol Sayısı değerini kontrol edin; gerekiyorsa Kalıp_Bilgisi sayfasında Max_Kol_Sayisi değerini düzeltin.' AS Excel_Duzeltme,
                NULL AS Malzeme,
                Kalıp,
                'Fabrika=' || COALESCE(CAST(Fabrika AS TEXT), '') ||
                ' | Firin=' || COALESCE(CAST(Firin AS TEXT), '') ||
                ' | Hatlar=' || COALESCE(CAST(Hatlar AS TEXT), '') ||
                ' | Hat_No=' || COALESCE(CAST("Hat No" AS TEXT), '') ||
                ' | Calisabilecegi_Max_Kol_Sayisi=' || COALESCE(CAST(Calisabilecegi_Max_Kol_Sayisi AS TEXT), '') ||
                ' | Hat_Bilgisindeki_Kol_Sayisi=' || COALESCE(CAST(Hat_Bilgisindeki_Kol_Sayisi AS TEXT), '') ||
                ' | Fiili_Hiz_Dam_Dak=' || COALESCE(CAST(Fiili_Hiz_Dam_Dak AS TEXT), '') ||
                ' | Birinci_Gun_Sonrasi_Verimi=' || COALESCE(CAST(Birinci_Gun_Sonrasi_Verimi AS TEXT), '') AS Detay
            FROM "Hatali_Kol_Sayilari_X"
        )
        ORDER BY COALESCE(Malzeme, 2147483647), COALESCE(Kalıp, 2147483647), Hata_Tipi
        """
    )

    for tablo_adi in (
        "Hatali_Kol_Sayilari_X",
        "Kalıp_Bilgisi_Eksik_X",
        "Parametre_Bilgisi_Eksik_X",
        "Sadesi_Bulunamayanlar_X",
    ):
        imlec.execute(f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(tablo_adi)}")

    imlec.execute("SELECT COUNT(*) FROM Hatalar_X")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "HATALAR_X_OLUSTUR",
        f"Hatalar_X tablosu {kayit_sayisi} kayit ile olusturuldu ve kaynak hata tablolari silindi",
        kayit_sayisi,
    )

    baglanti.commit()
    return kayit_sayisi


def tarih_metnini_coz(deger: Any) -> datetime | None:
    """Farkli tarih metinlerini datetime nesnesine cevirir."""
    if deger is None:
        return None

    metin = str(deger).strip()
    if not metin:
        return None

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(metin, fmt)
        except ValueError:
            continue

    return None


def tarih_metnine_cevir(tarih: datetime) -> str:
    """Datetime nesnesini YYYY-MM-DD olarak yazar."""
    return tarih.strftime("%Y-%m-%d")


def simdi_metni() -> str:
    """Kayit loglari icin ortak zaman metni uretir."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def tablo_sutunlarini_getir(baglanti: sqlite3.Connection, tablo_adi: str) -> set[str]:
    """Verilen tablonun mevcut sutun adlarini dondurur."""
    imlec = baglanti.cursor()
    return {
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(tablo_adi)})"
        )
    }


def tablo_kayit_sayisi(baglanti: sqlite3.Connection, tablo_adi: str) -> int:
    """Tablodaki toplam kayit sayisini dondurur."""
    imlec = baglanti.cursor()
    imlec.execute(f"SELECT COUNT(*) FROM {sql_kimlik_kacaga_dayanikli(tablo_adi)}")
    return imlec.fetchone()[0]


def tablo_var_mi(baglanti: sqlite3.Connection, tablo_adi: str) -> bool:
    imlec = baglanti.cursor()
    imlec.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ? LIMIT 1",
        (tablo_adi,),
    )
    return imlec.fetchone() is not None


def normalize_malzeme_anahtari(deger: Any) -> str:
    if deger is None:
        return ""
    tam_sayi = tam_sayi_metnine_cevir(deger)
    if tam_sayi is not None:
        return tam_sayi
    return str(deger).strip()


def sirali_tarih_sutunlarini_getir(mevcut_sutunlar: set[str]) -> list[str]:
    return sorted(
        [sutun for sutun in mevcut_sutunlar if TARIH_SUTUN_DESENI.match(sutun)],
        key=lambda sutun: tarih_metnini_coz(sutun) or datetime.max,
    )


def baslangic_stoku_sutununu_coz(mevcut_sutunlar: set[str]) -> str | None:
    baslangic_stoku_sutunu = sutun_adi_coz(
        mevcut_sutunlar,
        (
            "BaÅŸlangÄ±Ã§ Stoku",
            "Başlangıç Stoku",
            "Baslangic Stoku",
            "BaÃ…Å¸langÃ„Â±ÃƒÂ§ Stoku",
            "BaÃ…Â£langÃƒÂ½ÃƒÂ§ Stoku",
            "BaÃƒâ€¦Ã…Â¸langÃƒâ€Ã‚Â±ÃƒÆ’Ã‚Â§ Stoku",
        ),
    )
    if baslangic_stoku_sutunu is None:
        for sutun in sorted(mevcut_sutunlar):
            token = normalize_lookup_token(sutun)
            if not token or "stok" not in token:
                continue
            if "baslang" in token or "batlang" in token:
                baslangic_stoku_sutunu = sutun
                break
            if token.startswith("ba") and token.endswith(("stok", "stoku")):
                baslangic_stoku_sutunu = sutun
                break
    return baslangic_stoku_sutunu


def stok_tablosundan_pozitif_stoklu_malzemeler(
    baglanti: sqlite3.Connection,
) -> set[str]:
    if not tablo_var_mi(baglanti, "Stok"):
        return set()

    stok_sutunlari = tablo_sutunlarini_getir(baglanti, "Stok")
    malzeme_sutunu = sutun_adi_coz(stok_sutunlari, ("Malzeme",))
    stok_miktari_sutunu = sutun_adi_coz(
        stok_sutunlari,
        ("Toplam Stok Adet", "Toplam Stok", "Stok"),
    )
    if stok_miktari_sutunu is None:
        for sutun in sorted(stok_sutunlari):
            token = normalize_lookup_token(sutun)
            if "stok" in token and ("adet" in token or token.endswith(("stok", "stoku"))):
                stok_miktari_sutunu = sutun
                break
    if malzeme_sutunu is None or stok_miktari_sutunu is None:
        return set()

    imlec = baglanti.cursor()
    satirlar = imlec.execute(
        f"""
        SELECT
            CAST({sql_kimlik_kacaga_dayanikli(malzeme_sutunu)} AS TEXT) AS Malzeme,
            COALESCE(SUM(CAST({sql_kimlik_kacaga_dayanikli(stok_miktari_sutunu)} AS REAL)), 0) AS Toplam_Stok
        FROM Stok
        GROUP BY CAST({sql_kimlik_kacaga_dayanikli(malzeme_sutunu)} AS TEXT)
        HAVING COALESCE(SUM(CAST({sql_kimlik_kacaga_dayanikli(stok_miktari_sutunu)} AS REAL)), 0) > 0
        """
    ).fetchall()
    return {normalize_malzeme_anahtari(satir[0]) for satir in satirlar}


def gerekli_sutunu_coz(
    baglanti: sqlite3.Connection,
    tablo_adi: str,
    aday_sutunlar: tuple[str, ...],
) -> str:
    if not tablo_var_mi(baglanti, tablo_adi):
        raise ValueError(f"Gerekli tablo bulunamadi: {tablo_adi}")
    mevcut_sutunlar = tablo_sutunlarini_getir(baglanti, tablo_adi)
    sutun = sutun_adi_coz(mevcut_sutunlar, aday_sutunlar)
    if sutun is None:
        raise ValueError(
            f"{tablo_adi} tablosunda gerekli sutun bulunamadi: {', '.join(aday_sutunlar)}"
        )
    return sutun


def temp_anahtar_tablosunu_yenile(
    baglanti: sqlite3.Connection,
    temp_tablo_adi: str,
    anahtarlar: set[str],
) -> None:
    imlec = baglanti.cursor()
    kacisli_tablo = sql_kimlik_kacaga_dayanikli(temp_tablo_adi)
    imlec.execute(f"DROP TABLE IF EXISTS temp.{kacisli_tablo}")
    imlec.execute(
        f"CREATE TEMP TABLE {kacisli_tablo} (Anahtar TEXT PRIMARY KEY)"
    )
    if anahtarlar:
        imlec.executemany(
            f"INSERT INTO {kacisli_tablo} (Anahtar) VALUES (?)",
            [(anahtar,) for anahtar in sorted(anahtarlar)],
        )


def tabloyu_anahtar_kumesine_gore_daralt(
    baglanti: sqlite3.Connection,
    tablo_adi: str,
    sutun_adi: str,
    temp_tablo_adi: str,
) -> tuple[int, int]:
    imlec = baglanti.cursor()
    onceki = tablo_kayit_sayisi(baglanti, tablo_adi)
    kacisli_tablo = sql_kimlik_kacaga_dayanikli(tablo_adi)
    kacisli_sutun = sql_kimlik_kacaga_dayanikli(sutun_adi)
    kacisli_temp = sql_kimlik_kacaga_dayanikli(temp_tablo_adi)

    imlec.execute(
        f"""
        DELETE FROM {kacisli_tablo}
        WHERE NOT EXISTS (
            SELECT 1
            FROM {kacisli_temp} hedef
            WHERE hedef.Anahtar = normalize_anahtar(CAST({kacisli_tablo}.{kacisli_sutun} AS TEXT))
        )
        """
    )
    sonraki = tablo_kayit_sayisi(baglanti, tablo_adi)
    return onceki, sonraki


def hedef_malzeme_ve_kalip_anahtarlarini_hesapla(
    baglanti: sqlite3.Connection,
) -> tuple[set[str], set[str]]:
    baglanti.create_function("normalize_anahtar", 1, normalize_malzeme_anahtari)
    imlec = baglanti.cursor()

    talepler_malzeme_sutunu = gerekli_sutunu_coz(
        baglanti,
        KAYNAK_TABLO,
        ("Malzeme",),
    )
    talepler_tanim_sutunu = gerekli_sutunu_coz(
        baglanti,
        KAYNAK_TABLO,
        ("Tanım", "TanÄ±m", "Tanim"),
    )
    urun_agaci_malzeme_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Urun_Agaci",
        ("Malzeme",),
    )
    urun_agaci_bilesen_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Urun_Agaci",
        ("Bileşen", "BileÅŸen", "Bilesen"),
    )
    malzeme_data_malzeme_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Malzeme_Data",
        ("Malzeme",),
    )
    malzeme_data_kalip_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Malzeme_Data",
        (
            "Kalıp (Cam Ambalaj)",
            "KalÄ±p (Cam Ambalaj)",
            "Kalip (Cam Ambalaj)",
            "Kalıp",
            "KalÄ±p",
            "Kalip",
        ),
    )

    sade_malzemeler: set[str] = set()
    sade_olmayan_malzemeler: set[str] = set()
    kacisli_talepler = sql_kimlik_kacaga_dayanikli(KAYNAK_TABLO)
    kacisli_talepler_malzeme = sql_kimlik_kacaga_dayanikli(talepler_malzeme_sutunu)
    kacisli_talepler_tanim = sql_kimlik_kacaga_dayanikli(talepler_tanim_sutunu)
    for malzeme, tanim in imlec.execute(
        f"""
        SELECT
            CAST({kacisli_talepler_malzeme} AS TEXT) AS Malzeme,
            UPPER(TRIM(COALESCE(CAST({kacisli_talepler_tanim} AS TEXT), ''))) AS Tanim
        FROM {kacisli_talepler}
        """
    ).fetchall():
        malzeme_anahtari = normalize_malzeme_anahtari(malzeme)
        if not malzeme_anahtari:
            continue
        if tanim == "SADE":
            sade_malzemeler.add(malzeme_anahtari)
        else:
            sade_olmayan_malzemeler.add(malzeme_anahtari)

    sade_olmayan_temp = "__tmp_sade_olmayan_malzeme"
    temp_anahtar_tablosunu_yenile(baglanti, sade_olmayan_temp, sade_olmayan_malzemeler)

    bilesen_malzemeler: set[str] = set()
    kacisli_urun_agaci = sql_kimlik_kacaga_dayanikli("Urun_Agaci")
    kacisli_ua_malzeme = sql_kimlik_kacaga_dayanikli(urun_agaci_malzeme_sutunu)
    kacisli_ua_bilesen = sql_kimlik_kacaga_dayanikli(urun_agaci_bilesen_sutunu)
    kacisli_sade_olmayan_temp = sql_kimlik_kacaga_dayanikli(sade_olmayan_temp)
    for (bilesen_anahtari,) in imlec.execute(
        f"""
        SELECT DISTINCT normalize_anahtar(CAST(ua.{kacisli_ua_bilesen} AS TEXT)) AS Bilesen_Anahtar
        FROM {kacisli_urun_agaci} ua
        INNER JOIN {kacisli_sade_olmayan_temp} hedef
            ON normalize_anahtar(CAST(ua.{kacisli_ua_malzeme} AS TEXT)) = hedef.Anahtar
        WHERE normalize_anahtar(CAST(ua.{kacisli_ua_bilesen} AS TEXT)) != ''
        """
    ).fetchall():
        bilesen_malzemeler.add(bilesen_anahtari)

    hedef_malzeme_anahtarlari = sade_malzemeler | bilesen_malzemeler

    hedef_malzeme_temp = "__tmp_hedef_malzeme"
    temp_anahtar_tablosunu_yenile(baglanti, hedef_malzeme_temp, hedef_malzeme_anahtarlari)

    hedef_kalip_anahtarlari: set[str] = set()
    kacisli_malzeme_data = sql_kimlik_kacaga_dayanikli("Malzeme_Data")
    kacisli_md_malzeme = sql_kimlik_kacaga_dayanikli(malzeme_data_malzeme_sutunu)
    kacisli_md_kalip = sql_kimlik_kacaga_dayanikli(malzeme_data_kalip_sutunu)
    kacisli_hedef_malzeme_temp = sql_kimlik_kacaga_dayanikli(hedef_malzeme_temp)
    for (kalip_anahtari,) in imlec.execute(
        f"""
        SELECT DISTINCT normalize_anahtar(CAST(md.{kacisli_md_kalip} AS TEXT)) AS Kalip_Anahtar
        FROM {kacisli_malzeme_data} md
        INNER JOIN {kacisli_hedef_malzeme_temp} hedef
            ON normalize_anahtar(CAST(md.{kacisli_md_malzeme} AS TEXT)) = hedef.Anahtar
        WHERE normalize_anahtar(CAST(md.{kacisli_md_kalip} AS TEXT)) != ''
        """
    ).fetchall():
        hedef_kalip_anahtarlari.add(kalip_anahtari)

    return hedef_malzeme_anahtarlari, hedef_kalip_anahtarlari


def kaynak_tablolari_hedefe_gore_daralt(
    baglanti: sqlite3.Connection,
) -> tuple[int, int, dict[str, tuple[int, int]]]:
    baglanti.create_function("normalize_anahtar", 1, normalize_malzeme_anahtari)
    imlec = baglanti.cursor()

    hedef_malzeme_anahtarlari, hedef_kalip_anahtarlari = (
        hedef_malzeme_ve_kalip_anahtarlarini_hesapla(baglanti)
    )

    hedef_malzeme_temp = "__tmp_hedef_malzeme"
    hedef_kalip_temp = "__tmp_hedef_kalip"
    temp_anahtar_tablosunu_yenile(baglanti, hedef_malzeme_temp, hedef_malzeme_anahtarlari)
    temp_anahtar_tablosunu_yenile(baglanti, hedef_kalip_temp, hedef_kalip_anahtarlari)

    kalip_bilgisi_kalip_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Kalıp_Bilgisi",
        ("Kalıp Kodu", "KalÄ±p Kodu", "Kalip Kodu"),
    )
    parametre_kalip_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Parametre",
        ("Kalıp", "KalÄ±p", "Kalip"),
    )
    malzeme_data_malzeme_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Malzeme_Data",
        ("Malzeme",),
    )
    malzeme_data_kalip_sutunu = gerekli_sutunu_coz(
        baglanti,
        "Malzeme_Data",
        (
            "Kalıp (Cam Ambalaj)",
            "KalÄ±p (Cam Ambalaj)",
            "Kalip (Cam Ambalaj)",
            "Kalıp",
            "KalÄ±p",
            "Kalip",
        ),
    )

    gecisler: dict[str, tuple[int, int]] = {}
    stok_kayit_sayisi = tablo_kayit_sayisi(baglanti, "Stok")
    # Stok tablosu daraltma adimindan cikarildi; oldugu gibi korunur.
    gecisler["Stok"] = (stok_kayit_sayisi, stok_kayit_sayisi)
    gecisler["Kalıp_Bilgisi"] = tabloyu_anahtar_kumesine_gore_daralt(
        baglanti,
        "Kalıp_Bilgisi",
        kalip_bilgisi_kalip_sutunu,
        hedef_kalip_temp,
    )
    gecisler["Parametre"] = tabloyu_anahtar_kumesine_gore_daralt(
        baglanti,
        "Parametre",
        parametre_kalip_sutunu,
        hedef_kalip_temp,
    )

    malzeme_data_kayit_sayisi = tablo_kayit_sayisi(baglanti, "Malzeme_Data")
    # Malzeme_Data tablosu daraltma disi birakildi; oldugu gibi korunur.
    gecisler["Malzeme_Data"] = (malzeme_data_kayit_sayisi, malzeme_data_kayit_sayisi)

    log_kaydet(
        baglanti,
        "KAYNAK_DARALT_HEDEF_MALZEME",
        "Kaynak daraltma icin hedef malzeme listesi olusturuldu.",
        len(hedef_malzeme_anahtarlari),
    )
    log_kaydet(
        baglanti,
        "KAYNAK_DARALT_HEDEF_KALIP",
        "Kaynak daraltma icin hedef kalip listesi olusturuldu.",
        len(hedef_kalip_anahtarlari),
    )
    for tablo_adi, (onceki, sonraki) in gecisler.items():
        if tablo_adi in {"Stok", "Malzeme_Data"}:
            aciklama = f"{tablo_adi} tablosu daraltma disi birakildi: {onceki} -> {sonraki}"
        else:
            aciklama = f"{tablo_adi} tablosu hedef listeye gore daraltildi: {onceki} -> {sonraki}"
        log_kaydet(
            baglanti,
            "KAYNAK_DARALT_TABLO",
            aciklama,
            sonraki,
        )

    baglanti.commit()
    return len(hedef_malzeme_anahtarlari), len(hedef_kalip_anahtarlari), gecisler


def talepler_x_ham_snapshot_olustur(baglanti: sqlite3.Connection) -> int:
    mevcut_sutunlar = tablo_sutunlarini_getir(baglanti, HEDEF_TABLO)
    malzeme_sutunu = sutun_adi_coz(mevcut_sutunlar, ("Malzeme",))
    tarih_sutunlari = sirali_tarih_sutunlarini_getir(mevcut_sutunlar)
    toplam_talep_sutunu = sutun_adi_coz(mevcut_sutunlar, ("Toplam_Talep",))
    baslangic_stoku_sutunu = baslangic_stoku_sutununu_coz(mevcut_sutunlar)

    eksikler: list[str] = []
    if malzeme_sutunu is None:
        eksikler.append("Malzeme")
    if not tarih_sutunlari:
        eksikler.append("Tarih sutunlari")
    if toplam_talep_sutunu is None:
        eksikler.append("Toplam_Talep")
    if baslangic_stoku_sutunu is None:
        eksikler.append("Baslangic Stoku")
    if eksikler:
        raise ValueError(
            f"{HEDEF_TABLO} ham snapshot icin gerekli sutunlar eksik: {', '.join(eksikler)}"
        )

    secilecek_sutunlar = [
        malzeme_sutunu,
        *tarih_sutunlari,
        toplam_talep_sutunu,
        baslangic_stoku_sutunu,
    ]
    secim_ifadesi = ", ".join(
        f"{sql_kimlik_kacaga_dayanikli(sutun)} AS {sql_kimlik_kacaga_dayanikli(sutun)}"
        for sutun in secilecek_sutunlar
    )

    imlec = baglanti.cursor()
    imlec.execute(
        f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)}"
    )
    imlec.execute(
        f"""
        CREATE TABLE {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)} AS
        SELECT {secim_ifadesi}
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        """
    )
    imlec.execute(
        f"SELECT COUNT(*) FROM {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)}"
    )
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "HAM_TALEP_SNAPSHOT",
        f"{HAM_TALEP_SNAPSHOT_TABLOSU} tablosu olusturuldu",
        kayit_sayisi,
    )
    baglanti.commit()
    return kayit_sayisi


def talepler_x_ham_talebi_stoklu_hatali_malzemeler_icin_geri_yukle(
    baglanti: sqlite3.Connection,
) -> int:
    if not tablo_var_mi(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU):
        return 0
    if not tablo_var_mi(baglanti, "Hatalar_X"):
        return 0

    hedef_sutunlar = tablo_sutunlarini_getir(baglanti, HEDEF_TABLO)
    snapshot_sutunlar = tablo_sutunlarini_getir(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU)
    hedef_malzeme_sutunu = sutun_adi_coz(hedef_sutunlar, ("Malzeme",))
    snapshot_malzeme_sutunu = sutun_adi_coz(snapshot_sutunlar, ("Malzeme",))
    snapshot_stok_sutunu = baslangic_stoku_sutununu_coz(snapshot_sutunlar)

    if hedef_malzeme_sutunu is None or snapshot_malzeme_sutunu is None or snapshot_stok_sutunu is None:
        return 0

    hedef_tarih_sutunlari = sirali_tarih_sutunlarini_getir(hedef_sutunlar)
    snapshot_tarih_sutunlari = set(sirali_tarih_sutunlarini_getir(snapshot_sutunlar))
    guncellenecek_tarih_sutunlari = [
        sutun for sutun in hedef_tarih_sutunlari if sutun in snapshot_tarih_sutunlari
    ]

    hedef_toplam_talep_sutunu = sutun_adi_coz(hedef_sutunlar, ("Toplam_Talep",))
    snapshot_toplam_talep_sutunu = sutun_adi_coz(snapshot_sutunlar, ("Toplam_Talep",))
    hedef_stok_sutunu = baslangic_stoku_sutununu_coz(hedef_sutunlar)
    snapshot_stok_hedef_sutunu = baslangic_stoku_sutununu_coz(snapshot_sutunlar)

    guncellenecek_sutunlar: list[str] = list(guncellenecek_tarih_sutunlari)
    if (
        hedef_toplam_talep_sutunu
        and snapshot_toplam_talep_sutunu
        and hedef_toplam_talep_sutunu == snapshot_toplam_talep_sutunu
    ):
        guncellenecek_sutunlar.append(hedef_toplam_talep_sutunu)
    if (
        hedef_stok_sutunu
        and snapshot_stok_hedef_sutunu
        and hedef_stok_sutunu == snapshot_stok_hedef_sutunu
        and hedef_stok_sutunu not in guncellenecek_sutunlar
    ):
        guncellenecek_sutunlar.append(hedef_stok_sutunu)

    if not guncellenecek_sutunlar:
        return 0

    secim_sutunlari = [snapshot_malzeme_sutunu, *guncellenecek_sutunlar]
    secim_ifadesi = ", ".join(
        f"s.{sql_kimlik_kacaga_dayanikli(sutun)} AS {sql_kimlik_kacaga_dayanikli(sutun)}"
        for sutun in secim_sutunlari
    )
    snapshot_stok_ifadesi = (
        f"COALESCE(CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_stok_sutunu)} AS REAL), 0)"
    )

    imlec = baglanti.cursor()
    satirlar = imlec.execute(
        f"""
        SELECT {secim_ifadesi}
        FROM {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)} s
        WHERE {snapshot_stok_ifadesi} > 0
          AND EXISTS (
              SELECT 1
              FROM Hatalar_X h
              WHERE CAST(h.Malzeme AS TEXT) = CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_malzeme_sutunu)} AS TEXT)
          )
        """
    ).fetchall()

    if not satirlar:
        return 0

    set_ifadesi = ", ".join(
        f"{sql_kimlik_kacaga_dayanikli(sutun)} = ?"
        for sutun in guncellenecek_sutunlar
    )
    guncellenen = 0
    for satir in satirlar:
        guncelleme_degerleri = [satir[idx] for idx in range(1, len(secim_sutunlari))]
        guncelleme_degerleri.append(satir[0])
        imlec.execute(
            f"""
            UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
            SET {set_ifadesi}
            WHERE CAST({sql_kimlik_kacaga_dayanikli(hedef_malzeme_sutunu)} AS TEXT) = CAST(? AS TEXT)
            """,
            guncelleme_degerleri,
        )
        guncellenen += imlec.rowcount

    if guncellenen > 0:
        log_kaydet(
            baglanti,
            "HAM_TALEP_GERI_YUKLE",
            "Stoklu ve hatali malzemelerin ham talep/stok degerleri snapshot'tan geri yuklendi",
            guncellenen,
        )
    baglanti.commit()
    return guncellenen


def hatalar_x_stok_tam_karsilayanlari_temizle(baglanti: sqlite3.Connection) -> int:
    """Stogu toplam talebi tamamen karsilayan malzemeleri Hatalar_X'ten temizler."""
    if not tablo_var_mi(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU):
        return 0
    if not tablo_var_mi(baglanti, "Hatalar_X"):
        return 0

    snapshot_sutunlar = tablo_sutunlarini_getir(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU)
    snapshot_malzeme_sutunu = sutun_adi_coz(snapshot_sutunlar, ("Malzeme",))
    snapshot_stok_sutunu = baslangic_stoku_sutununu_coz(snapshot_sutunlar)
    snapshot_toplam_talep_sutunu = sutun_adi_coz(snapshot_sutunlar, ("Toplam_Talep",))

    if (
        snapshot_malzeme_sutunu is None
        or snapshot_stok_sutunu is None
        or snapshot_toplam_talep_sutunu is None
    ):
        return 0

    imlec = baglanti.cursor()
    imlec.execute(
        f"""
        WITH snapshot_ozet AS (
            SELECT
                CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_malzeme_sutunu)} AS TEXT) AS Malzeme_Anahtari,
                COALESCE(SUM(CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_toplam_talep_sutunu)} AS REAL)), 0) AS Toplam_Talep,
                COALESCE(MAX(CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_stok_sutunu)} AS REAL)), 0) AS Baslangic_Stok
            FROM {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)} s
            GROUP BY CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_malzeme_sutunu)} AS TEXT)
        )
        DELETE FROM Hatalar_X
        WHERE Malzeme IS NOT NULL
          AND EXISTS (
              SELECT 1
              FROM snapshot_ozet o
              WHERE o.Malzeme_Anahtari = CAST(Hatalar_X.Malzeme AS TEXT)
                AND o.Toplam_Talep > 0
                AND o.Baslangic_Stok > 0
                AND o.Baslangic_Stok >= o.Toplam_Talep
          )
        """
    )
    silinen = imlec.rowcount
    if silinen < 0:
        imlec.execute("SELECT changes()")
        sonuc = imlec.fetchone()
        silinen = int(sonuc[0]) if sonuc is not None else 0
    if silinen > 0:
        log_kaydet(
            baglanti,
            "HATALAR_X_STOK_TAM_KARSILAMA_TEMIZLE",
            "Stogu toplam talebi tamamen karsilayan malzemeler Hatalar_X tablosundan temizlendi",
            silinen,
        )
    baglanti.commit()
    return silinen


def hatalar_x_sifir_talep_silinenleri_temizle(baglanti: sqlite3.Connection) -> int:
    """SIFIR_TALEP_SIL adiminda elenen malzemeleri Hatalar_X'ten temizler."""
    if not tablo_var_mi(baglanti, "Hatalar_X"):
        return 0
    if not tablo_var_mi(baglanti, TALEPLER_X_SILINME_LOG_TABLOSU):
        return 0

    imlec = baglanti.cursor()
    imlec.execute(
        f"""
        DELETE FROM Hatalar_X
        WHERE Malzeme IS NOT NULL
          AND EXISTS (
              SELECT 1
              FROM {sql_kimlik_kacaga_dayanikli(TALEPLER_X_SILINME_LOG_TABLOSU)} l
              WHERE l.Silme_Adimi = 'SIFIR_TALEP_SIL'
                AND CAST(l.Malzeme AS TEXT) = CAST(Hatalar_X.Malzeme AS TEXT)
          )
        """
    )
    silinen = imlec.rowcount
    if silinen < 0:
        imlec.execute("SELECT changes()")
        sonuc = imlec.fetchone()
        silinen = int(sonuc[0]) if sonuc is not None else 0
    if silinen > 0:
        log_kaydet(
            baglanti,
            "HATALAR_X_SIFIR_TALEP_TEMIZLE",
            "SIFIR_TALEP_SIL adiminda elenen malzemeler Hatalar_X tablosundan temizlendi",
            silinen,
        )
    baglanti.commit()
    return silinen


def talepler_x_ham_snapshot_temizle(baglanti: sqlite3.Connection) -> None:
    if not tablo_var_mi(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU):
        return
    imlec = baglanti.cursor()
    imlec.execute(
        f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)}"
    )
    log_kaydet(
        baglanti,
        "HAM_TALEP_SNAPSHOT_TEMIZLE",
        f"{HAM_TALEP_SNAPSHOT_TABLOSU} tablosu temizlendi",
        0,
    )
    baglanti.commit()


def gruplu_metin_ifadesi(mevcut_sutunlar: set[str], aday_sutunlar: tuple[str, ...]) -> str:
    """Ilk bulunan sutun icin GROUP_CONCAT ifadesi dondurur."""
    for sutun in aday_sutunlar:
        if sutun in mevcut_sutunlar:
            return (
                f"GROUP_CONCAT(DISTINCT CAST({sql_kimlik_kacaga_dayanikli(sutun)} AS TEXT))"
            )
    return "NULL"


def talepler_x_silinme_logu_olustur(baglanti: sqlite3.Connection) -> None:
    """talepler_x tablosundan cikan malzemeleri nedenleriyle saklar."""
    imlec = baglanti.cursor()
    imlec.execute(
        f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(TALEPLER_X_SILINME_LOG_TABLOSU)}"
    )
    imlec.execute(
        f"""CREATE TABLE {sql_kimlik_kacaga_dayanikli(TALEPLER_X_SILINME_LOG_TABLOSU)} (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Log_Zamani TEXT,
            Silme_Adimi TEXT,
            Silme_Nedeni TEXT,
            Malzeme TEXT,
            Silinen_Kayit_Sayisi INTEGER,
            Kalip_Numaralari TEXT,
            Renkler TEXT,
            Tanimlar TEXT,
            Toplam_Talep_Toplami INTEGER
        )"""
    )


def talepler_x_adim_ozeti_olustur(baglanti: sqlite3.Connection) -> None:
    """Talepler'den final talepler_x'e kadar sayisal gecisleri tutar."""
    imlec = baglanti.cursor()
    imlec.execute(
        f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(TALEPLER_X_ADIM_OZET_TABLOSU)}"
    )
    imlec.execute(
        f"""CREATE TABLE {sql_kimlik_kacaga_dayanikli(TALEPLER_X_ADIM_OZET_TABLOSU)} (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Log_Zamani TEXT,
            Adim TEXT,
            Onceki_Kayit_Sayisi INTEGER,
            Sonraki_Kayit_Sayisi INTEGER,
            Fark INTEGER,
            Aciklama TEXT
        )"""
    )


def talepler_x_toplulastirma_logu_olustur(baglanti: sqlite3.Connection) -> None:
    """Toplulastirma nedeniyle birlesen kayitlari malzeme bazinda tutar."""
    imlec = baglanti.cursor()
    imlec.execute(
        f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(TALEPLER_X_TOPLULASTIRMA_LOG_TABLOSU)}"
    )
    imlec.execute(
        f"""CREATE TABLE {sql_kimlik_kacaga_dayanikli(TALEPLER_X_TOPLULASTIRMA_LOG_TABLOSU)} (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Log_Zamani TEXT,
            Malzeme TEXT,
            Onceki_Kayit_Sayisi INTEGER,
            Birlesen_Kayit_Sayisi INTEGER,
            Kalip_Numaralari TEXT,
            Renkler TEXT,
            Tanimlar TEXT
        )"""
    )


def talepler_x_adim_ozeti_kaydet(
    baglanti: sqlite3.Connection,
    adim: str,
    onceki_kayit_sayisi: int,
    sonraki_kayit_sayisi: int,
    aciklama: str,
) -> None:
    """Bir ETL adiminin kayit sayisina etkisini yazar."""
    imlec = baglanti.cursor()
    imlec.execute(
        f"""INSERT INTO {sql_kimlik_kacaga_dayanikli(TALEPLER_X_ADIM_OZET_TABLOSU)} (
            Log_Zamani,
            Adim,
            Onceki_Kayit_Sayisi,
            Sonraki_Kayit_Sayisi,
            Fark,
            Aciklama
        ) VALUES (?, ?, ?, ?, ?, ?)""",
        (
            simdi_metni(),
            adim,
            onceki_kayit_sayisi,
            sonraki_kayit_sayisi,
            sonraki_kayit_sayisi - onceki_kayit_sayisi,
            aciklama,
        ),
    )


def talepler_x_silinme_logu_kaydet(
    baglanti: sqlite3.Connection,
    silme_adimi: str,
    silme_nedeni: str,
    kosul_sql: str,
    parametreler: tuple[Any, ...] = (),
) -> None:
    """
    talepler_x'ten silinecek kayitlari, malzeme bazinda tek satir olacak sekilde loglar.
    """
    mevcut_sutunlar = tablo_sutunlarini_getir(baglanti, HEDEF_TABLO)
    kalip_ifadesi = gruplu_metin_ifadesi(mevcut_sutunlar, ("KalÄ±p NumarasÄ±",))
    renk_ifadesi = gruplu_metin_ifadesi(mevcut_sutunlar, ("Renk", "Renk TanÄ±mÄ±"))
    tanim_ifadesi = gruplu_metin_ifadesi(mevcut_sutunlar, ("TanÄ±m",))
    toplam_talep_ifadesi = (
        "SUM(COALESCE(Toplam_Talep, 0))" if "Toplam_Talep" in mevcut_sutunlar else "NULL"
    )

    imlec = baglanti.cursor()
    imlec.execute(
        f"""
        INSERT INTO {sql_kimlik_kacaga_dayanikli(TALEPLER_X_SILINME_LOG_TABLOSU)} (
            Log_Zamani,
            Silme_Adimi,
            Silme_Nedeni,
            Malzeme,
            Silinen_Kayit_Sayisi,
            Kalip_Numaralari,
            Renkler,
            Tanimlar,
            Toplam_Talep_Toplami
        )
        SELECT
            ?,
            ?,
            ?,
            CAST(Malzeme AS TEXT) AS Malzeme,
            COUNT(*) AS Silinen_Kayit_Sayisi,
            {kalip_ifadesi} AS Kalip_Numaralari,
            {renk_ifadesi} AS Renkler,
            {tanim_ifadesi} AS Tanimlar,
            {toplam_talep_ifadesi} AS Toplam_Talep_Toplami
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        WHERE {kosul_sql}
        GROUP BY Malzeme
        """,
        (simdi_metni(), silme_adimi, silme_nedeni, *tuple(parametreler)),
    )


def guvenli_sayiya_cevir(deger: Any) -> float | None:
    """Sayisal olmayan metinlerde hata firlatmadan None doner."""
    try:
        return sayiya_cevir(deger)
    except (TypeError, ValueError):
        return None


def malzeme_uyum_hata_satiri(
    hata_tipi: str,
    hata_aciklamasi: str,
    excel_duzeltme: str,
    malzeme: Any,
    kalip: Any,
    detay: str,
) -> tuple[str, str, str, Any, Any, str]:
    """Hatalar_X icin tek satirlik kayit hazirlar."""
    return (hata_tipi, hata_aciklamasi, excel_duzeltme, malzeme, kalip, detay)


def malzeme_fabrika_firin_hat_tarih_uyumu_olustur(
    baglanti: sqlite3.Connection,
) -> tuple[int, int]:
    """
    talepler_x, Ayarlar_X, Renk_Plani_X, Hat_Bilgisi ve Kapasiteler_x
    tablolarini birlestirerek Malzeme_Fabrika_Firin_Hat_Tarih_Uyumu tablosunu
    olusturur. Hatali satirlari Hatalar_X tablosuna ekler.
    """
    onceki_row_factory = baglanti.row_factory
    baglanti.row_factory = sqlite3.Row
    imlec = baglanti.cursor()

    imlec.execute(f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(MALZEME_UYUM_TABLOSU)}")
    imlec.execute(
        f"""
        CREATE TABLE {sql_kimlik_kacaga_dayanikli(MALZEME_UYUM_TABLOSU)} (
            Malzeme TEXT,
            Baslangic_Tarihi DATE,
            Bitis_Tarihi DATE,
            Fabrika TEXT,
            Firin TEXT,
            Hat TEXT,
            Birinci_Gun_Uretim_Adedi INTEGER,
            Birinci_Gunden_Sonraki_Gunlerdeki_Uretim_Adedi INTEGER,
            Gunluk_Cekis_Ton REAL,
            Kol_Sayisi INTEGER,
            Damla_Sayisi INTEGER,
            Gunluk_Kapasite_Ton REAL,
            Kapasite_Kullanim_Orani REAL
        )
        """
    )

    aday_satirlar = imlec.execute(
        """
        SELECT
            t.Malzeme,
            t."Kalıp Numarası" AS Kalip,
            t.Renk,
            t.Gramaj,
            a.Fabrika,
            a.Firin,
            a.Hatlar AS Hat,
            a.Calisabilecegi_Max_Kol_Sayisi AS Kol_Sayisi,
            a.Fiili_Hiz_Dam_Dak,
            a.Birinci_Gun_Sonrasi_Verimi AS Verim,
            rp."Başlangıç Tarihi" AS Baslangic_Tarihi,
            rp."Bitiş Tarihi" AS Bitis_Tarihi
        FROM talepler_x t
        INNER JOIN Ayarlar_X a
            ON CAST(t."Kalıp Numarası" AS TEXT) = CAST(a.Kalıp AS TEXT)
        INNER JOIN Renk_Plani_X rp
            ON CAST(a.Fabrika AS TEXT) = CAST(rp.Fabrika AS TEXT)
           AND CAST(a.Firin AS TEXT) = CAST(rp.Firin AS TEXT)
           AND TRIM(CAST(t.Renk AS TEXT)) = TRIM(CAST(rp.Renk AS TEXT))
        ORDER BY t.Malzeme, rp."Başlangıç Tarihi", rp."Bitiş Tarihi", a.Fabrika, a.Firin, a.Hatlar
        """
    ).fetchall()

    kapasite_map: dict[tuple[str, str, str, str], float] = {}
    for satir in imlec.execute(
        """
        SELECT Fabrika, Firin, Hat, Renk, "Günlük_Üretim_Tonajı" AS Gunluk_Kapasite_Ton
        FROM Kapasiteler_x
        """
    ).fetchall():
        anahtar = (
            str(satir["Fabrika"]).strip(),
            str(satir["Firin"]).strip(),
            str(satir["Hat"]).strip(),
            str(satir["Renk"]).strip(),
        )
        kapasite = guvenli_sayiya_cevir(satir["Gunluk_Kapasite_Ton"])
        if kapasite is not None:
            kapasite_map[anahtar] = kapasite

    damla_map: dict[tuple[str, str, str], list[sqlite3.Row]] = {}
    for satir in imlec.execute(
        """
        SELECT Fabrika, Firin, Hatlar, Damla, Baslangic_Tarihi, Bitis_Tarihi
        FROM Hat_Bilgisi
        ORDER BY Fabrika, Firin, Hatlar, Baslangic_Tarihi, Bitis_Tarihi
        """
    ).fetchall():
        anahtar = (
            str(satir["Fabrika"]).strip(),
            str(satir["Firin"]).strip(),
            str(satir["Hatlar"]).strip(),
        )
        damla_map.setdefault(anahtar, []).append(satir)

    sonuc_kayitlari: list[tuple[Any, ...]] = []
    hata_kayitlari: list[tuple[str, str, str, Any, Any, str]] = []

    for satir in aday_satirlar:
        malzeme = satir["Malzeme"]
        kalip = satir["Kalip"]
        renk = str(satir["Renk"]).strip()
        fabrika = str(satir["Fabrika"]).strip()
        firin = str(satir["Firin"]).strip()
        hat = str(satir["Hat"]).strip()

        baslangic = tarih_metnini_coz(satir["Baslangic_Tarihi"])
        bitis = tarih_metnini_coz(satir["Bitis_Tarihi"])
        gramaj = guvenli_sayiya_cevir(satir["Gramaj"])
        kol_sayisi = guvenli_sayiya_cevir(satir["Kol_Sayisi"])
        hiz = guvenli_sayiya_cevir(satir["Fiili_Hiz_Dam_Dak"])
        verim = guvenli_sayiya_cevir(satir["Verim"])

        detay_ortak = (
            f"Fabrika={fabrika} | Firin={firin} | Hat={hat} | Renk={renk} "
            f"| Baslangic_Tarihi={satir['Baslangic_Tarihi']} | Bitis_Tarihi={satir['Bitis_Tarihi']}"
        )

        if baslangic is None or bitis is None or baslangic > bitis:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Gecersiz_Tarih_Araligi",
                    "Baslangic tarihi bitis tarihinden sonra ya da tarih formati gecersiz.",
                    "Excel'de Renk_Plani_X sayfasindaki tarih alanlarini kontrol edin.",
                    malzeme,
                    kalip,
                    detay_ortak,
                )
            )
            continue

        if gramaj is None or gramaj <= 0:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Gramaj_Eksik",
                    "Malzemenin gramaj bilgisi eksik veya gecersiz.",
                    "Excel'de Parametre veya talepler kaynagindaki gramaj bilgisini duzeltin.",
                    malzeme,
                    kalip,
                    detay_ortak,
                )
            )
            continue

        if verim is None or verim <= 0 or hiz is None or hiz <= 0 or kol_sayisi is None or kol_sayisi <= 0:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Verim_Eksik",
                    "Verim, hiz veya kol sayisi bilgisi eksik ya da gecersiz.",
                    "Excel'de Ayarlar_X kaynagini besleyen Parametre ve Kalip_Bilgisi sayfalarini kontrol edin.",
                    malzeme,
                    kalip,
                    detay_ortak
                    + f" | Kol_Sayisi={satir['Kol_Sayisi']} | Fiili_Hiz_Dam_Dak={satir['Fiili_Hiz_Dam_Dak']} | Verim={satir['Verim']}",
                )
            )
            continue

        damla_satiri = None
        for aday_damla_satiri in damla_map.get((fabrika, firin, hat), []):
            damla_baslangic = tarih_metnini_coz(aday_damla_satiri["Baslangic_Tarihi"])
            damla_bitis = tarih_metnini_coz(aday_damla_satiri["Bitis_Tarihi"])
            if damla_baslangic is None or damla_bitis is None:
                continue
            if baslangic <= damla_bitis and bitis >= damla_baslangic:
                if damla_satiri is None:
                    damla_satiri = aday_damla_satiri
                    continue

                secili_baslangic = tarih_metnini_coz(damla_satiri["Baslangic_Tarihi"])
                if secili_baslangic is None or damla_baslangic > secili_baslangic:
                    damla_satiri = aday_damla_satiri

        if damla_satiri is None:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Damla_Sayisi_Eksik",
                    "Belirtilen tarih araliginda damla sayisi bulunamadi.",
                    "Excel'de Hat_Bilgisi sayfasinda ilgili Fabrika-Firin-Hat icin tarih kapsamini ve Damla bilgisini tamamlayin.",
                    malzeme,
                    kalip,
                    detay_ortak,
                )
            )
            continue

        damla_sayisi = guvenli_sayiya_cevir(damla_satiri["Damla"])
        if damla_sayisi is None or damla_sayisi <= 0:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Damla_Sayisi_Eksik",
                    "Damla sayisi eksik veya gecersiz.",
                    "Excel'de Hat_Bilgisi sayfasindaki Damla kolonunu kontrol edin.",
                    malzeme,
                    kalip,
                    detay_ortak + f" | Damla={damla_satiri['Damla']}",
                )
            )
            continue

        kapasite = kapasite_map.get((fabrika, firin, hat, renk))
        if kapasite is None or kapasite <= 0:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Kapasite_Eksik",
                    "Fabrika-Firin-Hat-Renk kombinasyonu icin kapasite tanimli degil.",
                    "Excel'de Kapasiteler sayfasinda ilgili kombinasyon icin gunluk tonaj tanimlayin.",
                    malzeme,
                    kalip,
                    detay_ortak,
                )
            )
            continue

        maksimum_kol_sayisi = int(kapasite * 1_000_000 // (gramaj * damla_sayisi * hiz * 1440))
        kullanilacak_kol_sayisi = min(int(kol_sayisi), maksimum_kol_sayisi)

        if kullanilacak_kol_sayisi < 1:
            hata_kayitlari.append(
                malzeme_uyum_hata_satiri(
                    "Kapasite_Yetersiz",
                    "En dusuk kol sayisinda bile kapasite asiliyor.",
                    "Excel'de kapasite tonajini veya urun parametrelerini gozden gecirin.",
                    malzeme,
                    kalip,
                    detay_ortak
                    + f" | Gunluk_Kapasite_Ton={kapasite} | Hesaplanan_Maksimum_Kol_Sayisi={maksimum_kol_sayisi}",
                )
            )
            continue

        birinci_gun_uretim_adedi = int(round(kullanilacak_kol_sayisi * damla_sayisi * hiz * verim * 0.8 * 1440))
        sonraki_gun_uretim_adedi = int(round(kullanilacak_kol_sayisi * damla_sayisi * hiz * verim * 1440))
        gunluk_cekis_ton = (kullanilacak_kol_sayisi * gramaj * damla_sayisi * hiz * 1440) / 1_000_000
        kapasite_kullanim_orani = (gunluk_cekis_ton / kapasite) * 100

        sonuc_kayitlari.append(
            (
                str(malzeme),
                tarih_metnine_cevir(baslangic),
                tarih_metnine_cevir(bitis),
                fabrika,
                firin,
                hat,
                birinci_gun_uretim_adedi,
                sonraki_gun_uretim_adedi,
                round(gunluk_cekis_ton, 4),
                kullanilacak_kol_sayisi,
                int(damla_sayisi),
                round(kapasite, 4),
                round(kapasite_kullanim_orani, 4),
            )
        )

    sonuc_kayitlari.sort(key=lambda x: (x[0], x[3], x[4], x[5], x[1], x[2]))

    if sonuc_kayitlari:
        imlec.executemany(
            f"""
            INSERT INTO {sql_kimlik_kacaga_dayanikli(MALZEME_UYUM_TABLOSU)} (
                Malzeme,
                Baslangic_Tarihi,
                Bitis_Tarihi,
                Fabrika,
                Firin,
                Hat,
                Birinci_Gun_Uretim_Adedi,
                Birinci_Gunden_Sonraki_Gunlerdeki_Uretim_Adedi,
                Gunluk_Cekis_Ton,
                Kol_Sayisi,
                Damla_Sayisi,
                Gunluk_Kapasite_Ton,
                Kapasite_Kullanim_Orani
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            sonuc_kayitlari,
        )

    yer_tutucular = ", ".join("?" for _ in MALZEME_UYUM_HATA_TIPLERI)
    imlec.execute(
        f"DELETE FROM Hatalar_X WHERE Hata_Tipi IN ({yer_tutucular})",
        MALZEME_UYUM_HATA_TIPLERI,
    )
    if hata_kayitlari:
        imlec.executemany(
            """
            INSERT INTO Hatalar_X (
                Hata_Tipi,
                Hata_Aciklamasi,
                Excel_Duzeltme,
                Malzeme,
                Kalıp,
                Detay
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            hata_kayitlari,
        )

    log_kaydet(
        baglanti,
        "MALZEME_UYUM_TABLOSU_OLUSTUR",
        f"{MALZEME_UYUM_TABLOSU} tablosu yeniden olusturuldu. Basarili kayit: {len(sonuc_kayitlari)}, hata: {len(hata_kayitlari)}",
        len(sonuc_kayitlari),
    )

    baglanti.row_factory = onceki_row_factory
    baglanti.commit()
    return len(sonuc_kayitlari), len(hata_kayitlari)


def renk_plani_tablosunu_duzenle(
    baglanti: sqlite3.Connection,
) -> tuple[int, int, int, int, int, int, int, int]:
    """
    Renk_Plani tablosunda istenen kolon duzenlemelerini yapar:
    - CampaignType ve ColorCode kolonlarini siler
    - LocationCode -> Fabrika
    - ResourceCode -> Firin
    - CharacteristicValue degerini Renk_Plani_Map'te arayip Aciklama'ya yazar
    - CharacteristicValue kolonunu siler
    - Aciklama kolonunu Renk olarak yeniden adlandirir
    - StartPeriod -> Başlangıç Tarihi
    - EndPeriod -> Bitiş Tarihi
    - Son tablo adini Renk_Plani_X yapar

    Returns:
        tuple: (
            silinen_kolon_sayisi,
            yeniden_adlandirilan_kolon_sayisi,
            map_eslesen_kayit_sayisi,
            map_eslesmeyen_kayit_sayisi,
            characteristic_silinen_sayisi,
            aciklama_renk_rename_sayisi,
            tarih_kolon_rename_sayisi,
            tablo_rename_sayisi,
        )
    """
    imlec = baglanti.cursor()
    renk_plani_tablo = "Renk_Plani"
    hedef_renk_plani_tablo = "Renk_Plani_X"
    map_tablo = "Renk_Plani_Map"

    tablolar = {
        satir[0]
        for satir in imlec.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    if renk_plani_tablo not in tablolar:
        raise ValueError(f"Tablo bulunamadi: {renk_plani_tablo}")
    if map_tablo not in tablolar:
        raise ValueError(f"Tablo bulunamadi: {map_tablo}")

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]

    silinen_kolon_sayisi = 0
    for sutun_adi in ("CampaignType", "ColorCode"):
        if sutun_adi in sutunlar:
            imlec.execute(
                f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
                f"DROP COLUMN {sql_kimlik_kacaga_dayanikli(sutun_adi)}"
            )
            silinen_kolon_sayisi += 1

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]

    yeniden_adlandirilan_kolon_sayisi = 0
    if "LocationCode" in sutunlar and "Fabrika" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('LocationCode')} "
            f"TO {sql_kimlik_kacaga_dayanikli('Fabrika')}"
        )
        yeniden_adlandirilan_kolon_sayisi += 1

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]
    if "ResourceCode" in sutunlar and "Firin" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('ResourceCode')} "
            f"TO {sql_kimlik_kacaga_dayanikli('Firin')}"
        )
        yeniden_adlandirilan_kolon_sayisi += 1

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]

    hedef_renk_kolonu = "Aciklama"
    if "Aciklama" not in sutunlar and "Renk" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"ADD COLUMN {sql_kimlik_kacaga_dayanikli('Aciklama')} TEXT"
        )
        sutunlar.append("Aciklama")
    if "Renk" in sutunlar:
        hedef_renk_kolonu = "Renk"

    if "CharacteristicValue" in sutunlar:
        imlec.execute(
            f"""UPDATE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)}
            SET {sql_kimlik_kacaga_dayanikli(hedef_renk_kolonu)} = (
                SELECT m.{sql_kimlik_kacaga_dayanikli('Aciklama')}
                FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
                WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')}) =
                      TRIM({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)}.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')})
                LIMIT 1
            )
            WHERE {sql_kimlik_kacaga_dayanikli('CharacteristicValue')} IS NOT NULL
              AND TRIM({sql_kimlik_kacaga_dayanikli('CharacteristicValue')}) != ''"""
        )

        imlec.execute(
            f"""SELECT COUNT(*)
            FROM {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} r
            WHERE r.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')} IS NOT NULL
              AND TRIM(r.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')}) != ''
              AND EXISTS (
                  SELECT 1
                  FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
                  WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')}) =
                        TRIM(r.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')})
              )"""
        )
        map_eslesen_kayit_sayisi = imlec.fetchone()[0]

        imlec.execute(
            f"""SELECT COUNT(*)
            FROM {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} r
            WHERE r.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')} IS NOT NULL
              AND TRIM(r.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')}) != ''
              AND NOT EXISTS (
                  SELECT 1
                  FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
                  WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')}) =
                        TRIM(r.{sql_kimlik_kacaga_dayanikli('CharacteristicValue')})
              )"""
        )
        map_eslesmeyen_kayit_sayisi = imlec.fetchone()[0]
    else:
        map_eslesen_kayit_sayisi = 0
        map_eslesmeyen_kayit_sayisi = 0

    characteristic_silinen_sayisi = 0
    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]
    if "CharacteristicValue" in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"DROP COLUMN {sql_kimlik_kacaga_dayanikli('CharacteristicValue')}"
        )
        characteristic_silinen_sayisi = 1

    aciklama_renk_rename_sayisi = 0
    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]
    if "Aciklama" in sutunlar and "Renk" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('Aciklama')} "
            f"TO {sql_kimlik_kacaga_dayanikli('Renk')}"
        )
        aciklama_renk_rename_sayisi = 1

    tarih_kolon_rename_sayisi = 0
    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]
    if "StartPeriod" in sutunlar and "Başlangıç Tarihi" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('StartPeriod')} "
            f"TO {sql_kimlik_kacaga_dayanikli('Başlangıç Tarihi')}"
        )
        tarih_kolon_rename_sayisi += 1

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(renk_plani_tablo)})"
        )
    ]
    if "EndPeriod" in sutunlar and "Bitiş Tarihi" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('EndPeriod')} "
            f"TO {sql_kimlik_kacaga_dayanikli('Bitiş Tarihi')}"
        )
        tarih_kolon_rename_sayisi += 1

    imlec.execute(f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(hedef_renk_plani_tablo)}")
    imlec.execute(
        f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(renk_plani_tablo)} "
        f"RENAME TO {sql_kimlik_kacaga_dayanikli(hedef_renk_plani_tablo)}"
    )
    tablo_rename_sayisi = 1

    log_kaydet(
        baglanti,
        "RENK_PLANI_DUZENLE",
        (
            f"Renk_Plani duzenlendi: silinen={silinen_kolon_sayisi}, "
            f"rename={yeniden_adlandirilan_kolon_sayisi}, map_eslesen={map_eslesen_kayit_sayisi}, "
            f"map_eslesmeyen={map_eslesmeyen_kayit_sayisi}, tarih_rename={tarih_kolon_rename_sayisi}, "
            f"tablo_rename={tablo_rename_sayisi}"
        ),
        map_eslesen_kayit_sayisi,
    )

    baglanti.commit()
    return (
        silinen_kolon_sayisi,
        yeniden_adlandirilan_kolon_sayisi,
        map_eslesen_kayit_sayisi,
        map_eslesmeyen_kayit_sayisi,
        characteristic_silinen_sayisi,
        aciklama_renk_rename_sayisi,
        tarih_kolon_rename_sayisi,
        tablo_rename_sayisi,
    )


def talep_renk_map_kolonlarini_ekle_ve_doldur(
    baglanti: sqlite3.Connection,
) -> tuple[int, int, int, int, int]:
    """
    talepler_x tablosuna Firindaki_Renk ve FH kolonlarini ekler.
    Renk Tanimi kolonunu Talep_Renk Map tablosundaki Talepteki_Renk
    ile eslestirerek ilgili kolonlari doldurur.

    Returns:
        tuple: (
            eklenen_kolon_sayisi,
            eslesen_kayit_sayisi,
            eslesmeyen_kayit_sayisi,
            silinen_kolon_sayisi,
            yeniden_adlandirilan_kolon_sayisi,
        )
    """
    imlec = baglanti.cursor()
    map_tablo = "Talep_Renk Map"

    tablolar = {
        satir[0]
        for satir in imlec.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    if map_tablo not in tablolar:
        raise ValueError(f"Renk esleme tablosu bulunamadi: {map_tablo}")

    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]

    eklenecek_kolonlar = [("Firindaki_Renk", "TEXT"), ("FH", "TEXT")]
    eklenen_kolon_sayisi = 0
    for sutun_adi, sutun_tipi in eklenecek_kolonlar:
        if sutun_adi in sutunlar:
            continue
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"ADD COLUMN {sql_kimlik_kacaga_dayanikli(sutun_adi)} {sutun_tipi}"
        )
        eklenen_kolon_sayisi += 1

    if eklenen_kolon_sayisi > 0:
        log_kaydet(
            baglanti,
            "KOLON_EKLE",
            "Firindaki_Renk ve FH kolonlari eklendi",
            eklenen_kolon_sayisi,
        )

    hedef_sutun_kumesi = set(sutunlar)
    hedef_malzeme_sutunu = sutun_adi_coz(hedef_sutun_kumesi, ("Malzeme",))
    hedef_renk_tanimi_sutunu = sutun_adi_coz(
        hedef_sutun_kumesi,
        ("Renk Tan\u0131m\u0131", "Renk Tanimi"),
    )
    if hedef_renk_tanimi_sutunu is None:
        raise ValueError(f"{HEDEF_TABLO} tablosunda Renk Tanimi kolonu bulunamadi.")

    # talepler_x tarafinda renk bilgisi bos ise once Malzeme_Data tablosundan doldur.
    # Doldurulan deger mevcut akistaki Talep_Renk Map donusumunde kullanilir.
    if "Malzeme_Data" in tablolar and hedef_malzeme_sutunu is not None:
        malzeme_data_sutunlari = tablo_sutunlarini_getir(baglanti, "Malzeme_Data")
        md_malzeme_sutunu = sutun_adi_coz(malzeme_data_sutunlari, ("Malzeme",))
        md_renk_sutunu = sutun_adi_coz(
            malzeme_data_sutunlari,
            (
                "Renk",
                "Renk(Cam Ambalaj)",
                "Renk (Cam Ambalaj)",
                "Renk Cam Ambalaj",
            ),
        )
        if md_malzeme_sutunu is not None and md_renk_sutunu is not None:
            imlec.execute(
                f"""UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
                SET {sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} = (
                        SELECT md.{sql_kimlik_kacaga_dayanikli(md_renk_sutunu)}
                        FROM {sql_kimlik_kacaga_dayanikli("Malzeme_Data")} md
                        WHERE TRIM(CAST(md.{sql_kimlik_kacaga_dayanikli(md_malzeme_sutunu)} AS TEXT)) =
                              TRIM(CAST({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}.{sql_kimlik_kacaga_dayanikli(hedef_malzeme_sutunu)} AS TEXT))
                          AND md.{sql_kimlik_kacaga_dayanikli(md_renk_sutunu)} IS NOT NULL
                          AND TRIM(CAST(md.{sql_kimlik_kacaga_dayanikli(md_renk_sutunu)} AS TEXT)) != ''
                        LIMIT 1
                    )
                WHERE (
                        {sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} IS NULL
                        OR TRIM(CAST({sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} AS TEXT)) = ''
                    )
                  AND {sql_kimlik_kacaga_dayanikli(hedef_malzeme_sutunu)} IS NOT NULL
                  AND TRIM(CAST({sql_kimlik_kacaga_dayanikli(hedef_malzeme_sutunu)} AS TEXT)) != ''
                  AND EXISTS (
                      SELECT 1
                      FROM {sql_kimlik_kacaga_dayanikli("Malzeme_Data")} md
                      WHERE TRIM(CAST(md.{sql_kimlik_kacaga_dayanikli(md_malzeme_sutunu)} AS TEXT)) =
                            TRIM(CAST({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}.{sql_kimlik_kacaga_dayanikli(hedef_malzeme_sutunu)} AS TEXT))
                        AND md.{sql_kimlik_kacaga_dayanikli(md_renk_sutunu)} IS NOT NULL
                        AND TRIM(CAST(md.{sql_kimlik_kacaga_dayanikli(md_renk_sutunu)} AS TEXT)) != ''
                  )"""
            )
            renk_tanimi_doldurulan_sayisi = imlec.rowcount
            if renk_tanimi_doldurulan_sayisi > 0:
                log_kaydet(
                    baglanti,
                    "TALEP_RENK_DOLDUR",
                    "Bos Renk Tanimi degerleri Malzeme_Data tablosundan dolduruldu",
                    renk_tanimi_doldurulan_sayisi,
                )

    imlec.execute(
        f"""UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        SET {sql_kimlik_kacaga_dayanikli('Firindaki_Renk')} = (
                SELECT m.{sql_kimlik_kacaga_dayanikli('Firindaki_Renk')}
                FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
                WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('Talepteki_Renk')}) =
                      TRIM({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)})
                LIMIT 1
            ),
            {sql_kimlik_kacaga_dayanikli('FH')} = (
                SELECT m.{sql_kimlik_kacaga_dayanikli('FH')}
                FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
                WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('Talepteki_Renk')}) =
                      TRIM({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)})
                LIMIT 1
            )
        WHERE {sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} IS NOT NULL
          AND TRIM({sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)}) != ''"""
    )

    # Talep_Renk Map'te karsiligi olmayan renkleri bos birakma:
    # fırın rengi bulunamazsa talepteki renk adini koru.
    imlec.execute(
        f"""UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        SET {sql_kimlik_kacaga_dayanikli('Firindaki_Renk')} =
            {sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)}
        WHERE (
                {sql_kimlik_kacaga_dayanikli('Firindaki_Renk')} IS NULL
                OR TRIM(CAST({sql_kimlik_kacaga_dayanikli('Firindaki_Renk')} AS TEXT)) = ''
            )
          AND {sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} IS NOT NULL
          AND TRIM(CAST({sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} AS TEXT)) != ''"""
    )
    map_bulunamayan_renk_fallback_sayisi = imlec.rowcount
    if map_bulunamayan_renk_fallback_sayisi > 0:
        log_kaydet(
            baglanti,
            "TALEP_RENK_MAP_FALLBACK",
            "Talep_Renk Map'te karsiligi olmayan renklerde talepteki renk korunarak dolduruldu",
            map_bulunamayan_renk_fallback_sayisi,
        )

    imlec.execute(
        f"""SELECT COUNT(*)
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} t
        WHERE t.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} IS NOT NULL
          AND TRIM(t.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)}) != ''
          AND EXISTS (
              SELECT 1
              FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
              WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('Talepteki_Renk')}) =
                    TRIM(t.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)})
          )"""
    )
    eslesen_kayit_sayisi = imlec.fetchone()[0]

    imlec.execute(
        f"""SELECT COUNT(*)
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} t
        WHERE t.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)} IS NOT NULL
          AND TRIM(t.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)}) != ''
          AND NOT EXISTS (
              SELECT 1
              FROM {sql_kimlik_kacaga_dayanikli(map_tablo)} m
              WHERE TRIM(m.{sql_kimlik_kacaga_dayanikli('Talepteki_Renk')}) =
                    TRIM(t.{sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)})
          )"""
    )
    eslesmeyen_kayit_sayisi = imlec.fetchone()[0]

    log_kaydet(
        baglanti,
        "TALEP_RENK_MAP",
        f"Talep_Renk Map eslemesi yapildi: {eslesen_kayit_sayisi} eslesen, {eslesmeyen_kayit_sayisi} eslesmeyen",
        eslesen_kayit_sayisi,
    )

    # talep tarafindaki eski renk kolonunu kaldir
    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]
    silinen_kolon_sayisi = 0
    if hedef_renk_tanimi_sutunu in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"DROP COLUMN {sql_kimlik_kacaga_dayanikli(hedef_renk_tanimi_sutunu)}"
        )
        silinen_kolon_sayisi = 1
        log_kaydet(
            baglanti,
            "KOLON_SIL",
            "Renk Tanımı kolonu silindi",
            silinen_kolon_sayisi,
        )

    # firindaki renk kolonunu Renk olarak adlandir
    sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)})"
        )
    ]
    yeniden_adlandirilan_kolon_sayisi = 0
    if "Firindaki_Renk" in sutunlar and "Renk" not in sutunlar:
        imlec.execute(
            f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"RENAME COLUMN {sql_kimlik_kacaga_dayanikli('Firindaki_Renk')} "
            f"TO {sql_kimlik_kacaga_dayanikli('Renk')}"
        )
        yeniden_adlandirilan_kolon_sayisi = 1
        log_kaydet(
            baglanti,
            "KOLON_RENAME",
            "Firindaki_Renk kolonu Renk olarak yeniden adlandirildi",
            yeniden_adlandirilan_kolon_sayisi,
        )

    baglanti.commit()
    return (
        eklenen_kolon_sayisi,
        eslesen_kayit_sayisi,
        eslesmeyen_kayit_sayisi,
        silinen_kolon_sayisi,
        yeniden_adlandirilan_kolon_sayisi,
    )


def pd_talepler_olustur(baglanti: sqlite3.Connection) -> int:
    """
    talepler_x tablosundan PD_Talepler tablosunu istenen kolonlarla yeniden olusturur.
    """
    imlec = baglanti.cursor()
    mevcut_sutunlar = tablo_sutunlarini_getir(baglanti, HEDEF_TABLO)
    tarih_sutunlari = sorted(
        [sutun for sutun in mevcut_sutunlar if TARIH_SUTUN_DESENI.match(sutun)],
        key=lambda sutun: tarih_metnini_coz(sutun) or datetime.max,
    )
    if not tarih_sutunlari:
        raise ValueError(
            f"{HEDEF_TABLO} tablosunda DD.MM.YYYY formatinda en az bir tarih sutunu bulunamadi."
        )

    malzeme_sutunu = sutun_adi_coz(mevcut_sutunlar, ("Malzeme",))
    toplam_talep_sutunu = sutun_adi_coz(mevcut_sutunlar, ("Toplam_Talep",))
    baslangic_stoku_sutunu = sutun_adi_coz(
        mevcut_sutunlar,
        ("Başlangıç Stoku", "BaÅŸlangÄ±Ã§ Stoku", "BaÅ£langÃ½Ã§ Stoku", "BaÃ…Å¸langÃ„Â±ÃƒÂ§ Stoku"),
    )
    if baslangic_stoku_sutunu is None:
        for sutun in sorted(mevcut_sutunlar):
            token = normalize_lookup_token(sutun)
            if not token or "stok" not in token:
                continue
            if "baslang" in token or "batlang" in token or (token.startswith("ba") and token.endswith(("stok", "stoku"))):
                baslangic_stoku_sutunu = sutun
                break
    fh_sutunu = sutun_adi_coz(mevcut_sutunlar, ("FH",))

    eksikler = []
    if malzeme_sutunu is None:
        eksikler.append("Malzeme")
    if toplam_talep_sutunu is None:
        eksikler.append("Toplam_Talep")
    if baslangic_stoku_sutunu is None:
        eksikler.append("Baslangic Stoku")
    if fh_sutunu is None:
        eksikler.append("FH")
    if eksikler:
        raise ValueError(
            f"{HEDEF_TABLO} tablosunda PD_Talepler icin gerekli sutunlar eksik: {', '.join(eksikler)}"
        )

    secilecek_sutunlar = [
        malzeme_sutunu,
        *tarih_sutunlari,
        toplam_talep_sutunu,
        baslangic_stoku_sutunu,
        fh_sutunu,
    ]
    secim_ifadesi = ",\n            ".join(
        f"src.{sql_kimlik_kacaga_dayanikli(sutun)} AS {sql_kimlik_kacaga_dayanikli(sutun)}"
        for sutun in secilecek_sutunlar
    )

    imlec.execute(f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(PD_TALEPLER_TABLOSU)}")
    imlec.execute(
        f"""
        CREATE TABLE {sql_kimlik_kacaga_dayanikli(PD_TALEPLER_TABLOSU)} AS
        SELECT
            {secim_ifadesi}
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} AS src
        """
    )
    imlec.execute(f"SELECT COUNT(*) FROM {sql_kimlik_kacaga_dayanikli(PD_TALEPLER_TABLOSU)}")
    kayit_sayisi = imlec.fetchone()[0]
    log_kaydet(
        baglanti,
        "PD_TALEPLER_OLUSTUR",
        f"{PD_TALEPLER_TABLOSU} tablosu yeniden olusturuldu",
        kayit_sayisi,
    )
    baglanti.commit()
    return kayit_sayisi


def planlama_sqlite_olustur(
    baglanti: sqlite3.Connection,
    planlama_yolu: Path,
) -> tuple[int, dict[str, int]]:
    """
    Secili tabloları planlama.sqlite icine sifirdan kopyalar.
    """
    mevcut_tablolar = {
        satir[0]
        for satir in baglanti.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    eksik_tablolar = [tablo for tablo in PLANLAMA_TABLOLARI if tablo not in mevcut_tablolar]
    if eksik_tablolar:
        raise ValueError(
            "planlama.sqlite olusturmak icin eksik tablolar var: "
            + ", ".join(eksik_tablolar)
        )
    aktarilacak_tablolar = list(PLANLAMA_TABLOLARI)
    for tablo_adi in PLANLAMA_OPSIYONEL_TABLOLARI:
        if tablo_adi in mevcut_tablolar:
            aktarilacak_tablolar.append(tablo_adi)

    silinen_dosya_sayisi = sqlite_veritabanini_temizle(planlama_yolu)
    imlec = baglanti.cursor()
    imlec.execute("ATTACH DATABASE ? AS planlama", (str(planlama_yolu),))

    tablo_kayit_sayilari: dict[str, int] = {}
    try:
        for tablo_adi in aktarilacak_tablolar:
            kacisli_tablo = sql_kimlik_kacaga_dayanikli(tablo_adi)
            imlec.execute(
                f"CREATE TABLE planlama.{kacisli_tablo} AS "
                f"SELECT * FROM main.{kacisli_tablo}"
            )
            imlec.execute(f"SELECT COUNT(*) FROM planlama.{kacisli_tablo}")
            tablo_kayit_sayilari[tablo_adi] = imlec.fetchone()[0]
        baglanti.commit()
    finally:
        imlec.execute("DETACH DATABASE planlama")

    log_kaydet(
        baglanti,
        "PLANLAMA_SQLITE_OLUSTUR",
        f"{planlama_yolu.name} icine {len(tablo_kayit_sayilari)} tablo aktarildi",
        len(tablo_kayit_sayilari),
    )
    baglanti.commit()
    return silinen_dosya_sayisi, tablo_kayit_sayilari



def rotasi_eksik_malzemeleri_tespit_et(baglanti: sqlite3.Connection) -> int:
    """
    talepler_x tablosunda olup Malzeme_Fabrika_Firin_Hat_Tarih_Uyumu tablosunda
    hiç rotası bulunmayan malzemeleri tespit eder ve Hatalar_X tablosuna ekler.
    """
    imlec = baglanti.cursor()
    rota_yok_kosulu = f"""
        Malzeme NOT IN (
            SELECT DISTINCT Malzeme
            FROM {sql_kimlik_kacaga_dayanikli(MALZEME_UYUM_TABLOSU)}
        )
    """
    rota_yok_kosulu_t = f"""
        t.Malzeme NOT IN (
            SELECT DISTINCT Malzeme
            FROM {sql_kimlik_kacaga_dayanikli(MALZEME_UYUM_TABLOSU)}
        )
    """
    rota_silinebilir_kosulu = rota_yok_kosulu
    if tablo_var_mi(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU):
        snapshot_sutunlari = tablo_sutunlarini_getir(baglanti, HAM_TALEP_SNAPSHOT_TABLOSU)
        snapshot_malzeme_sutunu = sutun_adi_coz(snapshot_sutunlari, ("Malzeme",))
        snapshot_stok_sutunu = baslangic_stoku_sutununu_coz(snapshot_sutunlari)
        if snapshot_malzeme_sutunu and snapshot_stok_sutunu:
            rota_silinebilir_kosulu = f"""
                {rota_yok_kosulu}
                AND COALESCE((
                    SELECT CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_stok_sutunu)} AS REAL)
                    FROM {sql_kimlik_kacaga_dayanikli(HAM_TALEP_SNAPSHOT_TABLOSU)} s
                    WHERE CAST(s.{sql_kimlik_kacaga_dayanikli(snapshot_malzeme_sutunu)} AS TEXT) =
                          CAST({sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}.Malzeme AS TEXT)
                ), 0) <= 0
            """

    # Rota bulunmayan malzemeleri Hatalar_X tablosuna ekle
    imlec.execute(
        f"""
        INSERT INTO Hatalar_X (
            Hata_Tipi,
            Hata_Aciklamasi,
            Excel_Duzeltme,
            Malzeme,
            Kalıp,
            Detay
        )
        SELECT 
            'Rota_Eksik' AS Hata_Tipi,
            'Malzeme için Malzeme_Fabrika_Firin_Hat_Tarih_Uyumu tablosunda üretim rotası oluşturulamadı.' AS Hata_Aciklamasi,
            'Excel: Öncelikle Renk_Plani sayfasında malzemenin rengi için uygun Fabrika-Fırın planı olup olmadığını, ardından Parametre ve Kalıp_Bilgisi sayfaları içinde rotası olup olmadığını kontrol edin.' AS Excel_Duzeltme,
            t.Malzeme,
            t."Kalıp Numarası" AS Kalıp,
            'Renk=' || COALESCE(CAST(t.Renk AS TEXT), '') || 
            ' | Sebep=' || (
                CASE 
                    WHEN NOT EXISTS (SELECT 1 FROM Ayarlar_X a WHERE CAST(a.Kalıp AS TEXT) = CAST(t."Kalıp Numarası" AS TEXT)) 
                    THEN 'Kalıp/Parametre tanımı Ayarlar_X içinde bulunamadı.'
                    WHEN NOT EXISTS (SELECT 1 FROM Renk_Plani_X rp WHERE TRIM(CAST(rp.Renk AS TEXT)) = TRIM(CAST(t.Renk AS TEXT))) 
                    THEN 'Renk Planı (' || TRIM(CAST(t.Renk AS TEXT)) || ') Renk_Plani_X içinde bulunamadı.'
                    ELSE 'Uyuşmazlık! Kalıp şuralarda: (' || 
                         (SELECT GROUP_CONCAT(DISTINCT Fabrika || '-' || Firin) FROM Ayarlar_X a WHERE CAST(a.Kalıp AS TEXT) = CAST(t."Kalıp Numarası" AS TEXT)) ||
                         ') ama Renk Planı şuralarda: (' || 
                         (SELECT GROUP_CONCAT(DISTINCT Fabrika || '-' || Firin) FROM Renk_Plani_X rp WHERE TRIM(CAST(rp.Renk AS TEXT)) = TRIM(CAST(t.Renk AS TEXT))) || ')'
                END
            ) AS Detay        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} t
        WHERE {rota_yok_kosulu_t}
        """
    )

    # Rota bulunamayan malzemeleri talepler_x tablosundan sil
    talepler_x_silinme_logu_kaydet(
        baglanti,
        "ROTASI_EKSIK_SIL",
        "Malzeme icin uretim rotasi olusturulamadigi icin kayit silindi.",
        rota_silinebilir_kosulu,
    )
    imlec.execute(
        f"""
        DELETE FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        WHERE {rota_silinebilir_kosulu}
        """
    )
    silinen = imlec.rowcount

    if silinen > 0:
        log_kaydet(baglanti, "ROTASI_EKSIK_SIL", f"Üretim rotası bulunamayan {silinen} malzeme talepler_x'ten silindi", silinen)
        baglanti.commit()

    return silinen


def log_kaydet(baglanti: sqlite3.Connection, islem: str, aciklama: str, kayit_sayisi: int = None) -> None:
    """
    Yapılan her işlemi loglar_x tablosuna kaydeder.
    """
    imlec = baglanti.cursor()

    imlec.execute(
        "INSERT INTO loglar_x (Zaman, Islem, Aciklama, Kayit_Sayisi) VALUES (?, ?, ?, ?)",
        (simdi_metni(), islem, aciklama, kayit_sayisi)
    )


def loglar_x_olustur(baglanti: sqlite3.Connection) -> None:
    """
    İşlem loglarını tutacak loglar_x tablosunu oluşturur.
    """
    imlec = baglanti.cursor()
    imlec.execute("DROP TABLE IF EXISTS loglar_x")
    imlec.execute(
        """CREATE TABLE loglar_x (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Zaman TEXT,
            Islem TEXT,
            Aciklama TEXT,
            Kayit_Sayisi INTEGER
        )"""
    )


def talepler_x_olustur(baglanti: sqlite3.Connection) -> tuple[int, int, int]:
    """
    Talepler tablosundan gereksiz sutunlari atarak talepler_x olusturur.
    Sonrasinda ADT sutunlarinin adini ay sonu tarihe cevirir ve degerleri x1000 yapar.
    """
    imlec = baglanti.cursor()
    tablolar = {
        satir[0]
        for satir in imlec.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    if KAYNAK_TABLO not in tablolar:
        raise ValueError(f"Kaynak tablo bulunamadi: {KAYNAK_TABLO}")

    kaynak_sutunlar = [
        satir[1]
        for satir in imlec.execute(
            f"PRAGMA table_info({sql_kimlik_kacaga_dayanikli(KAYNAK_TABLO)})"
        )
    ]
    tutulacaklar = [s for s in kaynak_sutunlar if s not in SILINECEK_SUTUNLAR]
    if not tutulacaklar:
        raise ValueError("talepler_x olusturmak icin sutun kalmadi.")

    secim = ", ".join(sql_kimlik_kacaga_dayanikli(s) for s in tutulacaklar)
    imlec.execute(f"DROP TABLE IF EXISTS {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}")
    imlec.execute(
        f"CREATE TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"AS SELECT {secim} FROM {sql_kimlik_kacaga_dayanikli(KAYNAK_TABLO)}"
    )

    log_kaydet(baglanti, "TABLO_OLUSTUR", f"{KAYNAK_TABLO} tablosundan {HEDEF_TABLO} olusturuldu", len(tutulacaklar))

    # SADE_Malzemesi kolonunu ekle (TEXT tipinde, not yazilabilir)
    imlec.execute(
        f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"ADD COLUMN SADE_Malzemesi TEXT"
    )

    log_kaydet(baglanti, "KOLON_EKLE", "SADE_Malzemesi kolonu eklendi", 1)

    yeniden_adlandirilan = adt_sutunlarini_ay_sonu_tarihine_cevir(baglanti)
    carpilan_sutun_sayisi = adt_sutunlarini_bin_ile_carp(baglanti, yeniden_adlandirilan)
    
    log_kaydet(baglanti, "ADT_YENIDEN_ADLANDIR", f"ADT sutunlari ay sonu tarih formatina cevrildi", len(yeniden_adlandirilan))
    log_kaydet(baglanti, "ADT_CARPAN", f"ADT sutunlari 1000 ile carpildi", carpilan_sutun_sayisi)
    
    return len(tutulacaklar), len(yeniden_adlandirilan), carpilan_sutun_sayisi


def tanim_kolonunu_doldur(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    talepler_x tablosunda Tanım kolonu boş/null olan kayıtları
    Malzeme_Data tablosundan Hiyerarşi değeri ile doldurur.

    Returns:
        tuple: (doldurulan_sayisi, bulunamayan_sayisi)
    """
    imlec = baglanti.cursor()

    # Tanım'i boş/null olanları Malzeme_Data'dan Hiyerarşi ile doldur
    imlec.execute(
        f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"SET Tanım = ("
        f"    SELECT md.Hiyerarşi FROM Malzeme_Data md "
        f"    WHERE md.Malzeme = {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}.Malzeme"
        f") "
        f"WHERE Tanım IS NULL OR Tanım = ''"
    )
    doldurulan = imlec.rowcount
    
    if doldurulan > 0:
        log_kaydet(baglanti, "TANIM_DOLDUR", "Tanım kolonu Malzeme_Data'dan dolduruldu", doldurulan)

    # Hala boş/null olanlar (Malzeme_Data'da bulunamayanlar)
    imlec.execute(
        f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"SET Tanım = 'Bulunamadı' "
        f"WHERE Tanım IS NULL OR Tanım = ''"
    )
    bulunamayan = imlec.rowcount
    
    if bulunamayan > 0:
        log_kaydet(baglanti, "TANIM_BULUNAMADI", "Tanım kolonu için Malzeme_Data'da karşılık bulunamadı", bulunamayan)

    baglanti.commit()
    return doldurulan, bulunamayan


def sade_nerede_bul(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    talepler_x tablosuna SADE_Nerede kolonu ekler ve SADE malzeme kodlarının
    nerede bulunduğunu kaydeder.
    SADE malzemesi Malzeme_Data'da olanları talepler_x'e yeni kayıt olarak ekler.
    
    Mantık:
    1. SADE_Malzemesi kolonunda malzeme kodu olanları al (null ve 'SADE bulunamadı' hariç)
    2. Her SADE malzeme kodunu kontrol et:
       - talepler_x'te Malzeme kolonunda varsa → "talepler_x tablosunda"
       - Malzeme_Data'da Malzeme kolonunda varsa → "Malzeme_Data'da" + talepler_x'e ekle
       - İkisinde de yoksa → "SADE Malzemeye Ait Kayıt yok"
    
    Returns:
        tuple: (guncellenen_kayit_sayisi, eklenen_yeni_kayit_sayisi)
    """
    imlec = baglanti.cursor()
    
    # SADE_Nerede kolonunu ekle
    imlec.execute(
        f"ALTER TABLE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"ADD COLUMN SADE_Nerede TEXT"
    )
    
    log_kaydet(baglanti, "KOLON_EKLE", "SADE_Nerede kolonu eklendi", 1)
    
    # Distinct SADE malzeme kodlarını al
    imlec.execute(
        f"SELECT DISTINCT SADE_Malzemesi FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"WHERE SADE_Malzemesi IS NOT NULL AND SADE_Malzemesi != 'SADE bulunamadı'"
    )
    sade_kodlari = imlec.fetchall()
    
    guncellenen = 0
    talepler_x_sayisi = 0
    malzeme_data_sayisi = 0
    kayit_yok_sayisi = 0
    eklenen_kayit_sayisi = 0
    
    for (sade_kodu,) in sade_kodlari:
        # talepler_x'te var mı kontrol et
        imlec.execute(
            f"SELECT 1 FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} WHERE Malzeme = ? LIMIT 1",
            (sade_kodu,)
        )
        if imlec.fetchone():
            # talepler_x'te var - bu SADE malzemesine sahip TÜM kayıtları güncelle
            imlec.execute(
                f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
                f"SET SADE_Nerede = 'talepler_x tablosunda' "
                f"WHERE SADE_Malzemesi = ?",
                (sade_kodu,)
            )
            talepler_x_sayisi += 1
        else:
            # Malzeme_Data'da var mı kontrol et
            imlec.execute(
                "SELECT 1 FROM Malzeme_Data WHERE Malzeme = ? LIMIT 1",
                (sade_kodu,)
            )
            md_kaydi = imlec.fetchone()
            if md_kaydi:
                # Malzeme_Data'da var - talepler_x'e ekle (SADE_Malzemesi ve SADE_Nerede boş)
                imlec.execute(
                    "SELECT * FROM Malzeme_Data WHERE Malzeme = ?",
                    (sade_kodu,)
                )
                md_satir = imlec.fetchone()
                # md_satir: (Kalıp, Malzeme, Malzeme Tanımı, Renk, Ambalaj Tipi, İç Adet, Net Ağırlık, Hiyerarşi)
                kalip = md_satir[0]
                malzeme_kodu = md_satir[1]
                malzeme_tanimi = md_satir[7]
                renk = md_satir[3]
                ambalaj_tipi = md_satir[4]
                
                # talepler_x'e yeni kayıt ekle (SADE_Malzemesi ve SADE_Nerede boş bırakılıyor)
                # Kolonlar: Malzeme, Kalıp Numarası, Renk Tanımı, Tanım, 10 adet tarih kolonu
                imlec.execute(
                    f"""INSERT INTO {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} 
                    (Malzeme, "Kalıp Numarası", "Renk Tanımı", Tanım)
                    VALUES (?, ?, ?, ?)""",
                    (malzeme_kodu, kalip, renk, malzeme_tanimi)
                )
                eklenen_kayit_sayisi += imlec.rowcount
                
                # Bu SADE malzemesine sahip TÜM kayıtları güncelle
                imlec.execute(
                    f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
                    f"SET SADE_Nerede = 'Malzeme_Data''da' "
                    f"WHERE SADE_Malzemesi = ?",
                    (sade_kodu,)
                )
                
                malzeme_data_sayisi += 1
            else:
                imlec.execute(
                    f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
                    f"SET SADE_Nerede = 'SADE Malzemeye Ait Kayıt yok' "
                    f"WHERE SADE_Malzemesi = ?",
                    (sade_kodu,)
                )
                kayit_yok_sayisi += 1
        
        guncellenen += imlec.rowcount
    
    # Detaylı loglar
    log_kaydet(baglanti, "SADE_NEREDE_TALEPLER_X", "talepler_x tablosunda bulunan SADE kodları", talepler_x_sayisi)
    log_kaydet(baglanti, "SADE_NEREDE_MALZEME_DATA", "Malzeme_Data'da bulunan ve eklenen SADE kodları", malzeme_data_sayisi)
    log_kaydet(baglanti, "SADE_NEREDE_KAYIT_YOK", "Kayıt bulunamayan SADE kodları", kayit_yok_sayisi)
    log_kaydet(baglanti, "SADE_NEREDE_EKLENEN", "talepler_x'e yeni eklenen SADE kayıtları", eklenen_kayit_sayisi)
    toplam_kod = talepler_x_sayisi + malzeme_data_sayisi + kayit_yok_sayisi
    log_kaydet(baglanti, "SADE_NEREDE_OZET", f"talepler_x: {talepler_x_sayisi} | Malzeme_Data: {malzeme_data_sayisi} | Kayıt yok: {kayit_yok_sayisi} | Eklenen: {eklenen_kayit_sayisi}", toplam_kod)
    
    baglanti.commit()
    return guncellenen, eklenen_kayit_sayisi


def stok_x_olustur(baglanti: sqlite3.Connection) -> int:
    """
    Stok tablosundan Depo yeri ve Üretim yeri kolonlarını hariç tutarak,
    sadece Malzeme bazında toplam stokları hesaplar ve stok_x tablosuna kaydeder.
    
    Returns:
        int: stok_x'e eklenen kayıt sayısı
    """
    imlec = baglanti.cursor()
    
    # stok_x tablosunu oluştur
    imlec.execute("DROP TABLE IF EXISTS stok_x")
    imlec.execute(
        """CREATE TABLE stok_x (
            Malzeme INTEGER,
            "Toplam Stok" INTEGER
        )"""
    )
    
    # Malzeme bazında toplam stokları hesapla ve ekle
    imlec.execute(
        """INSERT INTO stok_x (Malzeme, "Toplam Stok")
        SELECT Malzeme, SUM("Toplam Stok Adet") as "Toplam Stok"
        FROM Stok
        GROUP BY Malzeme"""
    )
    
    eklenen = imlec.rowcount
    
    log_kaydet(baglanti, "STOK_X_OLUSTUR", "stok_x tablosu oluşturuldu (Malzeme bazında toplam stok)", eklenen)
    
    baglanti.commit()
    return eklenen


def sade_malzemesi_bul(baglanti: sqlite3.Connection) -> tuple[int, int]:
    """
    Tanım'i SADE olmayan kayıtlar için ürün ağacı hiyerarşisinde
    recursive olarak SADE malzeme kodunu bulup SADE_Malzemesi kolonunu günceller.
    SADE bulunamayanları Sadesi_Bulunamayanlar_X tablosuna yazar ve
    talepler_x'ten siler.

    Algoritma:
    1. talepler_x'te Tanım != 'SADE' olanların Malzeme değerlerini al
    2. Urun_Agaci'nda recursive olarak tüm bileşenleri gez
    3. Her bileşeni Malzeme_Data'da ara, Hiyerarşi == 'SADE' olanı bul
    4. Bulunan SADE malzeme kodunu talepler_x.SADE_Malzemesi'ye yaz
       SADE bulunamazsa 'SADE bulunamadı' notu yaz ve Sadesi_Bulunamayanlar_X'e ekle
       ve talepler_x'ten sil

    Returns:
        tuple: (guncellenen_kayit_sayisi, bulunamayan_sayisi)
    """
    imlec = baglanti.cursor()

    # Sadesi_Bulunamayanlar_X tablosunu oluştur
    imlec.execute("DROP TABLE IF EXISTS Sadesi_Bulunamayanlar_X")
    imlec.execute(
        f"""CREATE TABLE Sadesi_Bulunamayanlar_X AS
        SELECT Malzeme, Tanım, SADE_Malzemesi
        FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)}
        WHERE 0 = 1"""
    )
    
    log_kaydet(baglanti, "TABLO_OLUSTUR", "Sadesi_Bulunamayanlar_X tablosu oluşturuldu", 0)

    # Recursive CTE ile ürün ağacını gez ve SADE malzeme bul
    sorgu = """
    WITH RECURSIVE
    -- Tanımı SADE olmayan malzemeleri al
    Hedef_Malzemeler AS (
        SELECT DISTINCT Malzeme
        FROM talepler_x
        WHERE Tanım IS NOT NULL AND Tanım != 'SADE' AND Tanım != ''
    ),
    -- Ürün ağacını recursive olarak gez (tüm kademeler)
    Urun_Agaci_Recursive AS (
        -- İlk kademe: malzemenin direkt bileşeni
        SELECT 
            ua.Malzeme AS Ana_Malzeme,
            ua.Bileşen,
            1 AS Kademe
        FROM Urun_Agaci ua
        INNER JOIN Hedef_Malzemeler hm ON ua.Malzeme = hm.Malzeme
        
        UNION ALL
        
        -- Sonraki kademeler: bileşenin bileşeni
        SELECT 
            uar.Ana_Malzeme,
            ua.Bileşen,
            uar.Kademe + 1
        FROM Urun_Agaci_Recursive uar
        INNER JOIN Urun_Agaci ua ON uar.Bileşen = ua.Malzeme
        WHERE uar.Kademe < 20  -- Sonsuz döngüyü önlemek için kademe limiti
    ),
    -- Tüm bileşenleri Malzeme_Data ile eşleştir ve SADE olanları filtrele
    Sade_Adaylar AS (
        SELECT 
            uar.Ana_Malzeme,
            uar.Bileşen AS Sade_Malzeme,
            md.Hiyerarşi,
            uar.Kademe
        FROM Urun_Agaci_Recursive uar
        INNER JOIN Malzeme_Data md ON uar.Bileşen = md.Malzeme
        WHERE md.Hiyerarşi = 'SADE'
    ),
    -- Her ana malzeme için en üst kademede (en yakın) SADE malzemeyi seç
    En_Yakin_Sade AS (
        SELECT 
            Ana_Malzeme,
            Sade_Malzeme,
            ROW_NUMBER() OVER (PARTITION BY Ana_Malzeme ORDER BY Kademe ASC) AS rn
        FROM Sade_Adaylar
    )
    SELECT Ana_Malzeme, Sade_Malzeme
    FROM En_Yakin_Sade
    WHERE rn = 1
    """

    imlec.execute(sorgu)
    sonuclar = imlec.fetchall()

    # SADE bulunanları güncelle
    guncellenen = 0
    for ana_malzeme, sade_malzeme in sonuclar:
        imlec.execute(
            f"UPDATE {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"SET SADE_Malzemesi = ? "
            f"WHERE Malzeme = ?",
            (str(sade_malzeme), ana_malzeme)
        )
        guncellenen += imlec.rowcount
    
    if guncellenen > 0:
        log_kaydet(baglanti, "SADE_GUNCELLE", "SADE malzeme kodları güncellendi", guncellenen)

    # SADE bulunamayanları tespit et
    imlec.execute(
        f"SELECT DISTINCT Malzeme FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
        f"WHERE Tanım IS NOT NULL AND Tanım != 'SADE' AND Tanım != '' "
        f"AND SADE_Malzemesi IS NULL"
    )
    bulunamayanlar = imlec.fetchall()
    bulunamayan_malzemeler = tuple(malzeme_kodu for malzeme_kodu, in bulunamayanlar)
    stoklu_malzeme_kumesi = stok_tablosundan_pozitif_stoklu_malzemeler(baglanti)
    silinecek_malzemeler = tuple(
        malzeme_kodu
        for malzeme_kodu in bulunamayan_malzemeler
        if normalize_malzeme_anahtari(malzeme_kodu) not in stoklu_malzeme_kumesi
    )

    if silinecek_malzemeler:
        yer_tutucular = ",".join("?" * len(silinecek_malzemeler))
        talepler_x_silinme_logu_kaydet(
            baglanti,
            "SADE_BULUNAMADI",
            "Urun agacinda SADE karsiligi bulunamadigi icin kayit silindi.",
            f"Malzeme IN ({yer_tutucular})",
            silinecek_malzemeler,
        )

    bulunamayan_sayisi = 0
    stoklu_korunan_sayisi = 0
    for (malzeme_kodu,) in bulunamayanlar:
        # Sadesi_Bulunamayanlar_X'e ekle
        imlec.execute(
            f"INSERT INTO Sadesi_Bulunamayanlar_X (Malzeme, Tanım, SADE_Malzemesi) "
            f"SELECT Malzeme, Tanım, SADE_Malzemesi "
            f"FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"WHERE Malzeme = ?",
            (malzeme_kodu,)
        )

        if normalize_malzeme_anahtari(malzeme_kodu) in stoklu_malzeme_kumesi:
            stoklu_korunan_sayisi += 1
            continue

        # Stoku olmayanlar talepler_x'ten silinir.
        imlec.execute(
            f"DELETE FROM {sql_kimlik_kacaga_dayanikli(HEDEF_TABLO)} "
            f"WHERE Malzeme = ?",
            (malzeme_kodu,)
        )
        bulunamayan_sayisi += imlec.rowcount

    if bulunamayan_sayisi > 0:
        log_kaydet(baglanti, "SADE_BULUNAMADI", "SADE malzeme kodu bulunamadı", bulunamayan_sayisi)
        log_kaydet(baglanti, "BULUNAMAYANLAR_TABLO", "Sadesi_Bulunamayanlar_X tablosuna eklendi ve talepler_x'ten silindi", bulunamayan_sayisi)
    if stoklu_korunan_sayisi > 0:
        log_kaydet(
            baglanti,
            "SADE_BULUNAMADI_STOKLU_KORUNDU",
            "Stoklu malzemeler SADE bulunamasa da talepler_x'te korundu.",
            stoklu_korunan_sayisi,
        )

    baglanti.commit()
    return guncellenen, bulunamayan_sayisi


def sqlite_veritabanini_temizle(sqlite_yolu: Path) -> int:
    """
    SQLite veritabanini tamamen sifirlamak icin ana dosyayi ve
    olasi yan dosyalari (-wal, -shm, -journal) siler.
    """
    import time
    silinecekler = [
        sqlite_yolu,
        Path(str(sqlite_yolu) + "-wal"),
        Path(str(sqlite_yolu) + "-shm"),
        Path(str(sqlite_yolu) + "-journal"),
    ]
    silinen = 0
    for yol in silinecekler:
        if yol.exists():
            try:
                yol.unlink()
                silinen += 1
            except PermissionError:
                # Dosya kullanımda, 1 saniye bekle ve tekrar dene
                time.sleep(1)
                try:
                    yol.unlink()
                    silinen += 1
                except PermissionError:
                    # Hala kullanımda, uyarı ver ve devam et
                    print(f"UYARI: {yol} dosyasi kullanımda, silinemedi.")
    return silinen


def main(
    excel_yolu: str | Path | None = None,
    sqlite_yolu: str | Path | None = None,
    planlama_yolu: str | Path | None = None,
) -> None:
    """ETL akisini calistirir ve sonucu etl.sqlite dosyasina yazar."""
    excel_yolu = Path(excel_yolu) if excel_yolu is not None else Path("data.xlsx")
    sqlite_yolu = Path(sqlite_yolu) if sqlite_yolu is not None else Path("etl.sqlite")
    planlama_yolu = (
        Path(planlama_yolu) if planlama_yolu is not None else Path(PLANLAMA_SQLITE_DOSYASI)
    )

    emit_progress(2, "startup", "ETL ortami hazirlaniyor")
    if not excel_yolu.exists():
        raise FileNotFoundError(f"Excel dosyasi bulunamadi: {excel_yolu.resolve()}")

    emit_progress(8, "cleanup_sqlite", "Eski ETL ciktilari temizleniyor")
    silinen_dosya_sayisi = sqlite_veritabanini_temizle(sqlite_yolu)

    emit_progress(15, "load_excel", "Excel sayfalari okunuyor")
    sayfalar = pd.read_excel(excel_yolu, sheet_name=None)
    if not sayfalar:
        raise ValueError("data.xlsx icinde hic sayfa bulunamadi.")

    with sqlite3.connect(sqlite_yolu) as baglanti:
        emit_progress(26, "load_sheets", "Sayfalar veritabanina yukleniyor")
        # Log tablosunu oluştur
        loglar_x_olustur(baglanti)
        talepler_x_silinme_logu_olustur(baglanti)
        talepler_x_adim_ozeti_olustur(baglanti)
        talepler_x_toplulastirma_logu_olustur(baglanti)
        
        # Excel sayfalarını yükle
        for sayfa_adi, df in sayfalar.items():
            df.to_sql(sayfa_adi, baglanti, if_exists="replace", index=False)
            log_kaydet(baglanti, "SAYFA_YUKLE", f"{sayfa_adi} sayfası yüklendi", len(df))

        urun_agaci_kendi_esit_silinen = urun_agaci_malzeme_bilesen_ayni_kayitlari_sil(
            baglanti
        )

        emit_progress(33, "source_prune", "Kaynak tablolar hedef malzeme ve kaliba gore daraltiliyor")
        hedef_malzeme_sayisi, hedef_kalip_sayisi, kaynak_daraltma_gecisleri = (
            kaynak_tablolari_hedefe_gore_daralt(baglanti)
        )

        # Renk_Plani tablosunu map tablosuna gore duzenle
        (
            renk_plani_silinen_kolon_sayisi,
            renk_plani_rename_sayisi,
            renk_plani_map_eslesen_sayisi,
            renk_plani_map_eslesmeyen_sayisi,
            renk_plani_characteristic_silinen_sayisi,
            renk_plani_aciklama_renk_rename_sayisi,
            renk_plani_tarih_rename_sayisi,
            renk_plani_tablo_rename_sayisi,
        ) = renk_plani_tablosunu_duzenle(baglanti)
        emit_progress(40, "prep_support_tables", "Yardimci tablolar ve temel donusumler hazirlaniyor")

        damla_sayisi_x_kayit_sayisi = damla_sayisi_x_olustur(baglanti)
        playground_x_kayit_sayisi = playground_x_olustur(baglanti)
        kapasiteler_x_kayit_sayisi, kapasiteler_x_firin_eksik_sayisi = kapasiteler_x_olustur(baglanti)
        aylik_teorik_kapasite_x_kayit_sayisi = aylik_teorik_kapasite_x_olustur(baglanti)
        talepler_ilk_kayit_sayisi = tablo_kayit_sayisi(baglanti, KAYNAK_TABLO)

        tutulacak_sutun_sayisi, yeniden_adlandirilan_sayisi, carpilan_sayi = talepler_x_olustur(
            baglanti
        )
        talepler_x_ilk_kayit_sayisi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "TALEPLERDEN_TALEPLER_X",
            talepler_ilk_kayit_sayisi,
            talepler_x_ilk_kayit_sayisi,
            "Talepler tablosu talepler_x'e kopyalandi; bu adimda satir silinmedi.",
        )

        emit_progress(58, "demand_enrichment", "Talep temizleme ve zenginlestirme adimlari calisiyor")
        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        doldurulan_sayisi, tanim_bulunamayan_sayisi = tanim_kolonunu_doldur(baglanti)
        guncellenen_kayit_sayisi, sade_bulunamayan_sayisi = sade_malzemesi_bul(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "SADE_BULUNAMAYANLAR",
            onceki_sayi,
            sonraki_sayi,
            "Urun agacinda SADE karsiligi bulunamayan malzemeler cikarildi.",
        )

        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        sade_nerede_sayisi, eklenen_kayit_sayisi = sade_nerede_bul(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "SADE_KAYDI_EKLEME",
            onceki_sayi,
            sonraki_sayi,
            "Malzeme_Data'da bulunan SADE malzemeler talepler_x'e eklendi.",
        )
        stok_x_kayit_sayisi = stok_x_olustur(baglanti)
        
        # En son: tarihli kolonlardaki NULL ve sayisal olmayan degerleri 0 yap
        tarihli_kolon_sayisi, temizlenen_kayit_sayisi = tarihli_kolonlari_temizle(baglanti)
        
        # Tarihli kolonları temizledikten sonra aynı malzemeleri birleştir
        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        eski_kayit_sayisi, yeni_kayit_sayisi = talepler_x_toplula(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "TOPLULASTIRMA",
            onceki_sayi,
            sonraki_sayi,
            "Ayni Malzeme'ye ait birden fazla satir tek satirda birlestirildi.",
        )

        # Kalıp Numaralarını Kalıp_Bilgisi tablosunda kontrol et
        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        toplam_kalip, kalip_eslesen, kalip_eksik, kalip_eksik_kayit = kalip_bilgisi_kontrol_et(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "KALIP_BILGISI_ELEME",
            onceki_sayi,
            sonraki_sayi,
            "Kalip_Bilgisi tablosunda eslesmeyen kaliplarda stoksuz kayitlar cikarildi.",
        )

        # Gramaj kolonunu ekle (Parametre tablosundan Kalıp Numarası ile eşleşen ilk değeri al)
        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        gramaj_doldurulan_sayisi, gramaj_eksik_sayisi = gramaj_kolonu_ekle(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "GRAMAJ_ELEME",
            onceki_sayi,
            sonraki_sayi,
            "Parametre tablosunda gramaji bulunamayan kayitlarda stoksuz olanlar cikarildi.",
        )

        ham_snapshot_kayit_sayisi = talepler_x_ham_snapshot_olustur(baglanti)

        # Stokla talepleri karşıla (SADE olmayanlar için)
        (
            tahsis_edilen_kayit_sayisi,
            tahsis_edilen_kolon_sayisi,
            stok_tahsis_haritasi,
        ) = stok_tahsis_et(baglanti)

        # SADE olmayan malzemelerin taleplerini SADE malzemelere aktar
        aktarilan_kayit_sayisi, aktarilan_sutun_sayisi = sadeye_talep_aktar(
            baglanti,
            stok_tahsis_haritasi=stok_tahsis_haritasi,
        )

        # Toplam_Talep = 0 olan kayıtları sil
        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        silinen_kayit_sayisi = sifir_talepleri_sil(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "SIFIR_TALEP_ELEME",
            onceki_sayi,
            sonraki_sayi,
            "Stok tahsisi ve talep aktarimlari sonrasinda Toplam_Talep'i 0 olan kayitlar cikarildi.",
        )

        # SADE_Malzemesi ve SADE_Nerede kolonlarını sil
        emit_progress(74, "quality_and_routes", "Kalite kontrolleri ve rota eslesmeleri yapiliyor")
        silinen_kolon_sayisi = sade_kolonlarini_sil(baglanti)

        # Renk eşleme kolonlarını ekle ve Talep_Renk Map tablosundan doldur
        (
            renk_map_eklenen_kolon_sayisi,
            renk_map_eslesen_kayit_sayisi,
            renk_map_eslesmeyen_kayit_sayisi,
            renk_map_silinen_kolon_sayisi,
            renk_map_rename_kolon_sayisi,
        ) = talep_renk_map_kolonlarini_ekle_ve_doldur(baglanti)

        ayarlar_x_kayit_sayisi = ayarlar_x_olustur(baglanti)
        hatali_kol_sayilari_x_kayit_sayisi = hatali_kol_sayilari_x_olustur(baglanti)
        hatalar_x_kayit_sayisi = hatalar_x_olustur(baglanti)
        hatalar_x_sifir_talep_temizlenen_sayisi = (
            hatalar_x_sifir_talep_silinenleri_temizle(baglanti)
        )
        malzeme_uyum_kayit_sayisi, malzeme_uyum_hata_sayisi = malzeme_fabrika_firin_hat_tarih_uyumu_olustur(baglanti)
        
        # Üretim uyum tablosu oluştuktan sonra, talepler_x'te olup uyum tablosuna giremeyenleri bul
        # Bu işlem aynı zamanda bu malzemeleri talepler_x'ten siler
        onceki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        rotasi_eksik_sayisi = rotasi_eksik_malzemeleri_tespit_et(baglanti)
        sonraki_sayi = tablo_kayit_sayisi(baglanti, HEDEF_TABLO)
        talepler_x_adim_ozeti_kaydet(
            baglanti,
            "ROTA_ELEME",
            onceki_sayi,
            sonraki_sayi,
            "Uretim rotasi olusturulamayan malzemelerde stoksuz olanlar cikarildi.",
        )

        ham_talep_geri_yuklenen_sayisi = (
            talepler_x_ham_talebi_stoklu_hatali_malzemeler_icin_geri_yukle(baglanti)
        )
        hatalar_x_stok_tam_karsilama_temizlenen_sayisi = (
            hatalar_x_stok_tam_karsilayanlari_temizle(baglanti)
        )
        hatalar_x_son_kayit_sayisi = (
            tablo_kayit_sayisi(baglanti, "Hatalar_X")
            if tablo_var_mi(baglanti, "Hatalar_X")
            else 0
        )

        # PD_Talepler tablosunu EN SON (temizlenmiş talepler_x'ten) oluştur
        emit_progress(88, "publish_outputs", "Planlama icin cikti tablolari hazirlaniyor")
        pd_talepler_kayit_sayisi = pd_talepler_olustur(baglanti)
        talepler_x_ham_snapshot_temizle(baglanti)
        
        planlama_silinen_dosya_sayisi, planlama_tablo_kayit_sayilari = planlama_sqlite_olustur(
            baglanti,
            planlama_yolu,
        )

    # Bağlantı kapandıktan sonra kısa bekle
    import time
    emit_progress(98, "finalize", "ETL ozeti tamamlaniyor")
    time.sleep(0.5)

    print(f"{len(sayfalar)} sayfa yuklendi -> {sqlite_yolu.resolve()}")
    print(f"Eski SQLite dosyalarindan {silinen_dosya_sayisi} adet silindi.")
    print(
        "Urun_Agaci erken temizlik: "
        f"Malzeme=Bilesen olan {urun_agaci_kendi_esit_silinen} kayit silindi."
    )
    print(
        f"Kaynak daraltma hedefleri: {hedef_malzeme_sayisi} malzeme, {hedef_kalip_sayisi} kalip."
    )
    for tablo_adi in ("Stok", "Kalıp_Bilgisi", "Parametre", "Malzeme_Data"):
        onceki, sonraki = kaynak_daraltma_gecisleri.get(tablo_adi, (0, 0))
        print(f"Kaynak daraltma - {tablo_adi}: {onceki} -> {sonraki} kayit.")
    print(
        f"Renk_Plani duzenlendi: {renk_plani_silinen_kolon_sayisi} kolon silindi, "
        f"{renk_plani_rename_sayisi} kolon yeniden adlandirildi, "
        f"{renk_plani_map_eslesen_sayisi} kayit maplendi."
    )
    print(
        f"Renk_Plani son adimlar: CharacteristicValue silme={renk_plani_characteristic_silinen_sayisi}, "
        f"Aciklama->Renk rename={renk_plani_aciklama_renk_rename_sayisi}, "
        f"tarih kolon rename={renk_plani_tarih_rename_sayisi}, "
        f"tablo rename(Renk_Plani_X)={renk_plani_tablo_rename_sayisi}."
    )
    if renk_plani_map_eslesmeyen_sayisi > 0:
        print(
            f"UYARI: Renk_Plani icin {renk_plani_map_eslesmeyen_sayisi} kayitta Renk_Plani_Map eslesmesi bulunamadi."
        )
    print(f"Damla_Sayisi_x tablosu olusturuldu: {damla_sayisi_x_kayit_sayisi} kayit.")
    print(f"Playground_x tablosu olusturuldu: {playground_x_kayit_sayisi} tekil kayit.")
    print(f"Kapasiteler_x tablosu olusturuldu: {kapasiteler_x_kayit_sayisi} kayit.")
    print(f"Aylik_Teorik_Kapasite_X tablosu olusturuldu: {aylik_teorik_kapasite_x_kayit_sayisi} kayit.")
    if kapasiteler_x_firin_eksik_sayisi > 0:
        print(f"UYARI: Kapasiteler_x icin {kapasiteler_x_firin_eksik_sayisi} kayitta Firin bilgisi bulunamadi.")
    print(f"Ayarlar_X tablosu olusturuldu: {ayarlar_x_kayit_sayisi} ortak kayit.")
    print(
        "Hatalar_X tablosu guncellendi: "
        f"{hatalar_x_kayit_sayisi} ana kayit olusturuldu, "
        f"sifir talep adimindan temizlenen {hatalar_x_sifir_talep_temizlenen_sayisi}, "
        f"stokla tum talebi karsilayanlardan temizlenen {hatalar_x_stok_tam_karsilama_temizlenen_sayisi}, "
        f"son kayit sayisi {hatalar_x_son_kayit_sayisi}."
    )
    if rotasi_eksik_sayisi > 0:
        print(f"UYARI: {rotasi_eksik_sayisi} malzeme icin hicbir üretim rotası (uyum kaydı) oluşturulamadı. Bu malzemeler Hatalar_X tablosuna 'Rota_Eksik' tipiyle eklendi.")
    print(
        f"{MALZEME_UYUM_TABLOSU} tablosu olusturuldu: "
        f"{malzeme_uyum_kayit_sayisi} kayit, {malzeme_uyum_hata_sayisi} parametre hatasi."
    )
    print(f"{PD_TALEPLER_TABLOSU} tablosu olusturuldu: {pd_talepler_kayit_sayisi} kayit.")
    print(
        "Ham talep snapshot: "
        f"{ham_snapshot_kayit_sayisi} kayit, "
        f"stoklu-hatali geri yuklenen: {ham_talep_geri_yuklenen_sayisi}."
    )
    print(
        f"{planlama_yolu.name} olusturuldu: {planlama_silinen_dosya_sayisi} eski dosya silindi, "
        f"{len(planlama_tablo_kayit_sayilari)} tablo aktarildi."
    )
    for tablo_adi, kayit_sayisi in planlama_tablo_kayit_sayilari.items():
        print(f"  - {tablo_adi}: {kayit_sayisi} kayit")
    print(
        f"{KAYNAK_TABLO} tablosundan {tutulacak_sutun_sayisi} sutunla {HEDEF_TABLO} olusturuldu."
    )
    print(f"ADT adli {yeniden_adlandirilan_sayisi} sutun ay sonu tarih formatina cevrildi.")
    print(f"ADT adli {carpilan_sayi} sutundaki degerler 1000 ile carpildi.")
    print(f"Tarihli kolonlardaki sayisal olmayan degerler temizlendi ({tarihli_kolon_sayisi} kolon, {temizlenen_kayit_sayisi} kayit).")
    print(f"talepler_x toplulaştirildi: {eski_kayit_sayisi} -> {yeni_kayit_sayisi} kayit.")
    print(f"Kalip_Bilgisi kontrolu: Toplam {toplam_kalip} distinct kalip, {kalip_eslesen} eslesen, {kalip_eksik} eslesmeyen ({kalip_eksik_kayit} kayit).")
    if kalip_eksik_kayit > 0:
        print(f"UYARI: {kalip_eksik_kayit} kayit icin Kalip_Bilgisi'nde eslesme bulunamadi (Hatalar_X'e eklendi; stoksuz olanlar talepler_x'ten silindi).")
    print(f"Tanım kolonu {doldurulan_sayisi} kayit icin Malzeme_Data'dan dolduruldu.")

    if tanim_bulunamayan_sayisi > 0:
        print(f"UYARI: {tanim_bulunamayan_sayisi} kayit icin Malzeme_Data'da Hiyerarşi bulunamadi ('Bulunamadı' yazildi).")

    print(f"SADE_Malzemesi kolonu {guncellenen_kayit_sayisi} kayit icin guncellendi.")
    print(f"SADE_Nerede kolonu {sade_nerede_sayisi} kayit icin dolduruldu.")
    print(f"Malzeme_Data'dan talepler_x'e {eklenen_kayit_sayisi} yeni SADE kaydı eklendi.")
    print(f"Stok tablosundan {stok_x_kayit_sayisi} kayitla stok_x olusturuldu.")
    print(f"Gramaj kolonu {gramaj_doldurulan_sayisi} kayit icin Parametre'den dolduruldu.")
    if gramaj_eksik_sayisi > 0:
        print(f"UYARI: {gramaj_eksik_sayisi} kayit icin Parametre'de Gramaj bulunamadi (Hatalar_X'e eklendi; stoksuz olanlar talepler_x'ten silindi).")
    print(f"Başlangıç Stoku ile talep tahsisi yapildi: {tahsis_edilen_kayit_sayisi} kayit, {tahsis_edilen_kolon_sayisi} kolon.")
    print(f"SADE olmayan malzemelerden SADE'ye talep aktarildi: {aktarilan_kayit_sayisi} kayit, {aktarilan_sutun_sayisi} kolon.")
    print(f"Toplam_Talep = 0 olan kayitlar silindi (ham snapshot korumalariyla): {silinen_kayit_sayisi} kayit.")
    print(f"SADE_Malzemesi ve SADE_Nerede kolonlari silindi: {silinen_kolon_sayisi} kolon.")
    print(
        f"Talep_Renk Map eslemesi tamamlandi: {renk_map_eklenen_kolon_sayisi} kolon eklendi, "
        f"{renk_map_eslesen_kayit_sayisi} kayit eslesti."
    )
    print(
        f"Renk kolon duzeni tamamlandi: {renk_map_silinen_kolon_sayisi} kolon silindi, "
        f"{renk_map_rename_kolon_sayisi} kolon yeniden adlandirildi."
    )
    if renk_map_eslesmeyen_kayit_sayisi > 0:
        print(
            f"UYARI: {renk_map_eslesmeyen_kayit_sayisi} kayit icin Talep_Renk Map tablosunda "
            "renk karsiligi bulunamadi."
        )

    if sade_bulunamayan_sayisi > 0:
        print(f"UYARI: {sade_bulunamayan_sayisi} malzeme icin Urun_Agaci'nda SADE karsiligi bulunamadi.")
        print("Bu kayitlarin SADE_Malzemesi kolonuna 'SADE bulunamadı' yazildi.")
    emit_progress(100, "completed", "Veri hazirlama adimi tamamlandi")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run ETL and produce SQLite outputs.")
    parser.add_argument("--excel", default="data.xlsx", help="Input Excel file path")
    parser.add_argument("--sqlite", default="etl.sqlite", help="Output ETL SQLite file path")
    parser.add_argument(
        "--planlama-sqlite",
        default=PLANLAMA_SQLITE_DOSYASI,
        help="Output planning SQLite file path",
    )
    cli_args = parser.parse_args()
    main(cli_args.excel, cli_args.sqlite, cli_args.planlama_sqlite)
