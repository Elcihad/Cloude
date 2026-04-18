"""
SAP Suite v2 — Alan Tanıtma + Otomatik İndirici
Dinamik sayfa grupları, eylem listesi, insan gibi mouse hareketi.

Kurulum:
    pip install PyQt6 pyautogui pyperclip openpyxl pandas opencv-python mss pillow psutil
"""

import sys, json, os, time, subprocess, shutil, threading, math, random, base64, re
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QListWidget, QListWidgetItem,
    QSplitter, QFrame, QStatusBar, QMessageBox, QSpinBox, QComboBox,
    QLineEdit, QGroupBox, QTabWidget, QCheckBox, QTextEdit, QProgressBar,
    QSizePolicy, QScrollArea, QDialog, QDialogButtonBox, QDoubleSpinBox,
    QTreeWidget, QTreeWidgetItem, QMenu, QInputDialog
)
from PyQt6.QtCore import Qt, QRect, QPoint, QSize, pyqtSignal, QTimer, QThread
from PyQt6.QtGui import (
    QPixmap, QImage, QPainter, QPen, QColor, QFont,
    QAction, QCursor, QTextCursor, QShortcut, QKeySequence
)
import cv2, numpy as np

# ── Renkler ───────────────────────────────────────────────────────────────────
C = {
    "bg":      "#0f1117", "panel":  "#181c27", "card":   "#1e2334",
    "accent":  "#00d4aa", "accent2":"#6c63ff", "danger": "#ff4757",
    "warn":    "#ffa502", "text":   "#e8eaf0", "dim":    "#8892a4",
    "border":  "#2a3048",
}

SS = f"""
QMainWindow,QWidget{{background:{C['bg']};color:{C['text']};font-family:'Segoe UI',sans-serif;font-size:13px;}}
QTabWidget::pane{{border:1px solid {C['border']};background:{C['panel']};border-radius:8px;}}
QTabBar::tab{{background:{C['card']};color:{C['dim']};padding:10px 28px;border-radius:6px 6px 0 0;margin-right:2px;}}
QTabBar::tab:selected{{background:{C['accent']};color:#000;font-weight:bold;}}
QGroupBox{{background:{C['card']};border:1px solid {C['border']};border-radius:8px;margin-top:14px;padding:8px;
           font-weight:bold;color:{C['accent']};font-size:11px;letter-spacing:1px;}}
QGroupBox::title{{subcontrol-origin:margin;left:10px;padding:0 6px;}}
QPushButton{{background:{C['card']};color:{C['text']};border:1px solid {C['border']};
             border-radius:6px;padding:7px 16px;font-size:12px;min-height:32px;}}
QPushButton:hover{{background:{C['accent']};color:#000;border-color:{C['accent']};}}
QPushButton#accent{{background:{C['accent']};color:#000;border-color:{C['accent']};font-weight:bold;}}
QPushButton#accent:hover{{background:#00f0c0;}}
QPushButton#danger{{border-color:{C['danger']};color:{C['danger']};}}
QPushButton#danger:hover{{background:{C['danger']};color:#fff;}}
QPushButton#active{{background:{C['accent2']};color:#fff;border-color:{C['accent2']};}}
QPushButton:disabled{{background:{C['border']};color:{C['dim']};}}
QListWidget{{background:{C['panel']};border:1px solid {C['border']};border-radius:6px;color:{C['text']};outline:none;}}
QListWidget::item{{padding:6px 10px;border-bottom:1px solid {C['border']};}}
QListWidget::item:selected{{background:{C['accent2']};color:#fff;}}
QTreeWidget{{background:{C['panel']};border:1px solid {C['border']};border-radius:6px;color:{C['text']};outline:none;}}
QTreeWidget::item{{padding:4px 6px;}}
QTreeWidget::item:selected{{background:{C['accent2']};color:#fff;}}
QSpinBox,QComboBox,QLineEdit,QDoubleSpinBox{{background:{C['panel']};color:{C['text']};border:1px solid {C['border']};
                               border-radius:4px;padding:4px 8px;}}
QSpinBox:focus,QComboBox:focus,QLineEdit:focus,QDoubleSpinBox:focus{{border-color:{C['accent']};}}
QTextEdit{{background:{C['panel']};color:{C['text']};border:1px solid {C['border']};border-radius:6px;
           font-family:'Consolas',monospace;font-size:11px;}}
QCheckBox{{color:{C['text']};spacing:6px;}}
QCheckBox::indicator{{width:16px;height:16px;border:1px solid {C['border']};border-radius:3px;background:{C['panel']};}}
QCheckBox::indicator:checked{{background:{C['accent']};border-color:{C['accent']};}}
QProgressBar{{background:{C['panel']};border:1px solid {C['border']};border-radius:4px;height:16px;text-align:center;color:{C['text']};}}
QProgressBar::chunk{{background:{C['accent']};border-radius:3px;}}
QStatusBar{{background:{C['panel']};color:{C['dim']};border-top:1px solid {C['border']};font-size:11px;}}
QScrollBar:vertical{{background:{C['bg']};width:8px;border-radius:4px;}}
QScrollBar::handle:vertical{{background:{C['border']};border-radius:4px;min-height:30px;}}
QScrollBar::handle:vertical:hover{{background:{C['accent']};}}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{{height:0;}}
QSplitter::handle{{background:{C['border']};width:2px;}}
QFrame#canvas_frame{{background:{C['panel']};border:2px solid {C['border']};border-radius:8px;}}
QMenu{{background:{C['card']};color:{C['text']};border:1px solid {C['border']};border-radius:6px;padding:4px;}}
QMenu::item{{padding:6px 20px;border-radius:4px;}}
QMenu::item:selected{{background:{C['accent2']};color:#fff;}}
"""

# ── Sayfa Tipleri ─────────────────────────────────────────────────────────────
SAYFA_TIPLERI = {
    "giris":      {"etiket": "🔐 Giriş",       "renk": "#6c63ff", "aciklama": "Program başında 1 kez çalışır.\nÖrn: Chrome Aç → Bekle → Mail'den Kod Gir → Oturum aç tıkla"},
    "siparis":    {"etiket": "📋 Sipariş",      "renk": "#00d4aa", "aciklama": "Her sipariş için çalışır (Excel'den numara alır).\nÖrn: URL'ye Git → Kartı Tıkla → Numara Yaz → Export → Dosya Bekle"},
    "dongu_sonu": {"etiket": "🔄 Döngü Sonu",  "renk": "#ffa502", "aciklama": "Her sipariş bittikten sonra çalışır.\nÖrn: Geri dön, sayfayı sıfırla"},
    "cikis":      {"etiket": "🚪 Çıkış",       "renk": "#ff4757", "aciklama": "En sonda 1 kez çalışır.\nÖrn: Tarayıcıyı kapat"},
}

# ── Eylem Türleri ─────────────────────────────────────────────────────────────
EYLEMLER = {
    "sol_tikla":      {"etiket": "🖱  Sol Tıkla",           "icon": "🖱",  "renk": "#00d4aa"},
    "metin_yaz":      {"etiket": "⌨  Metin Yaz",            "icon": "⌨",  "renk": "#ffa502"},
    "enter":          {"etiket": "↵  Enter'a Bas",          "icon": "↵",  "renk": "#6c63ff"},
    "tab":            {"etiket": "⇥  Tab'a Bas",            "icon": "⇥",  "renk": "#52c7ff"},
    "bekle":          {"etiket": "⏳  Bekle (saniye)",       "icon": "⏳", "renk": "#8892a4"},
    "mouse_hareket":  {"etiket": "➡  Mouse Hareketi",       "icon": "➡",  "renk": "#ff6b9d"},
    "excel_numara":   {"etiket": "📊  Excel'den Numara Yaz", "icon": "📊", "renk": "#00b894"},
    "klavye_kisayol": {"etiket": "⌘  Klavye Kısayolu",      "icon": "⌘",  "renk": "#fdcb6e"},
    "chrome_ac":      {"etiket": "🌐  Chrome Aç",            "icon": "🌐", "renk": "#4a90d9"},
    "url_git":        {"etiket": "🔗  URL'ye Git",           "icon": "🔗", "renk": "#0984e3"},
    "mail_kod_gir":   {"etiket": "📧  Mail'den Kod Gir",     "icon": "📧", "renk": "#a29bfe"},
    "escape":         {"etiket": "⎋  Escape'e Bas",         "icon": "⎋",  "renk": "#636e72"},
    "f5":             {"etiket": "F5  F5'e Bas",              "icon": "F5", "renk": "#e17055"},
    "fonksiyon_tusu": {"etiket": "Fn  Fonksiyon Tuşu",         "icon": "Fn", "renk": "#d63031"},
    "dosya_bekle":    {"etiket": "⬇  Dosya İnmesini Bekle", "icon": "⬇",  "renk": "#55efc4"},
    "tus_ve_goruntu": {"etiket": "🔁  Tuşa Bas → Görüntü Bekle", "icon": "🔁", "renk": "#fd79a8"},
}

# Metin kaynakları (Metin Yaz eylemi için)
METIN_KAYNAKLARI = {
    "sabit":          "Sabit Metin",
    "kullanici_adi":  "Config — Kullanıcı Adı",
    "sifre":          "Config — Şifre",
    "excel_numara":   "Excel — Sipariş Numarası",
}

# ── Uygulama Klasörleri ───────────────────────────────────────────────────────
def _belgeler():
    home = Path.home()
    for p in [home/"OneDrive"/"Belgeler", home/"OneDrive"/"Documents",
              home/"Belgeler", home/"Documents"]:
        if p.exists(): return p
    return home / "Documents"

def _windows_downloads():
    return Path.home() / "Downloads"

APP_DIR     = _belgeler() / "SAP_Indirici"
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APP_DIR / "config.json"
LOG_FILE    = APP_DIR / "sap_log.txt"
AKIS_FILE   = APP_DIR / "sap_akis.json"

DEFAULT_CONFIG = {
    "sap_url":  "https://online.tupras.com.tr/sap/bc/ui2/flp?saml2=disabled&sap-client=100&sap-language=TR#Shell-home",
    "username": "", "password": "", "2fa_gerekli": False,
    "excel_file": "", "excel_sheet": 0, "numara_sutun": "C",
    "numara_baslangic_satir": 4,
    "download_folder":         str(APP_DIR/"downloads"),
    "pdf_download_folder":     str(APP_DIR/"downloads_pdf"),
    "ekran_bekleme_timeout":   12,
    "ekran_bekleme_esik":      0.80,
    "delay_between_numbers": 4, "retry_count": 3,
    "akis_json": str(AKIS_FILE),
    "akis_json_excel": "",
    "akis_json_pdf":   "",
    "github_token":    "",
    "github_repo":     "Elcihad/Cloude",
    "github_file":     "sap_suite_v2.py",
    "mail_konu_filtre":     "Doğrulama Kodu",
    "mail_gonderen_filtre": "",
    "mail_klasor_yolu":     "Doğrulama",
    "mail_kod_regex":       r"(?:Kodunuz|Verification Code)[:\s]+(\d{4,8})",
    # İnsan gibi davranma ayarları
    "mouse_hiz_min": 0.05,
    "mouse_hiz_max": 0.15,
    "tiklama_oncesi_bekleme_min": 0.03,
    "tiklama_oncesi_bekleme_max": 0.08,
    "yazi_hiz_min": 0.03,
    "yazi_hiz_max": 0.08,
}


# ── Mail'den 2FA Kodu Oku (Outlook COM + IMAP) ───────────────────────────────
def _kod_bul_regex(metin, regex):
    eslesme = re.search(regex, metin, re.IGNORECASE)
    if not eslesme: return None
    try:    return eslesme.group(1)
    except: return eslesme.group(0)

def _outlook_klasor_bul(inbox, klasor_yolu):
    """
    Outlook Inbox altında alt klasör yolunu traverse eder.
    klasor_yolu: "Tüpraş/Doğrulama" gibi / ile ayrılmış string.
    Bulamazsa None döner.
    """
    if not klasor_yolu or not klasor_yolu.strip():
        return inbox  # Yol verilmemişse doğrudan Inbox

    parcalar = [p.strip() for p in klasor_yolu.strip("/").split("/") if p.strip()]
    mevcut_klasor = inbox
    for parca in parcalar:
        bulundu = False
        try:
            for i in range(1, mevcut_klasor.Folders.Count + 1):
                alt = mevcut_klasor.Folders.Item(i)
                if alt.Name.lower() == parca.lower():
                    mevcut_klasor = alt
                    bulundu = True
                    break
        except Exception:
            return None
        if not bulundu:
            return None  # Klasör bulunamadı
    return mevcut_klasor


def mail_den_kod_oku(cfg, timeout=60):
    """
    Kurulu Outlook'u dener (pywin32 — sifresiz).
    mail_klasor_yolu config ayarıyla alt klasörü destekler.
    Örn: "Tüpraş/Doğrulama"  →  Inbox > Tüpraş > Doğrulama klasörüne bakar.
    Boş bırakılırsa sadece Inbox'a bakar.
    """
    filtre       = cfg.get("mail_konu_filtre", "Doğrulama Kodu")
    gonderen     = cfg.get("mail_gonderen_filtre", "")
    # "Tüpraş Online Doğrulama Kodunuz: 0501" ve "Verification Code: 0501" formatlarını yakalar
    regex        = cfg.get("mail_kod_regex", r"(?:Kodunuz|Verification Code)[:\s]+(\d{4,8})")
    klasor_yolu  = cfg.get("mail_klasor_yolu", "Doğrulama")   # Outlook → Doğrulama klasörü

    # ── Yontem 1: Kurulu Outlook (COM API) ───────────────────────────────────
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")
        inbox   = ns.GetDefaultFolder(6)  # 6 = Inbox

        # Alt klasör yolunu çöz (boşsa Inbox'ın kendisi döner)
        hedef_klasor = _outlook_klasor_bul(inbox, klasor_yolu)
        if hedef_klasor is None:
            # Klasör bulunamadıysa Inbox'a düş
            hedef_klasor = inbox

        bitis = time.time() + timeout
        while time.time() < bitis:
            try:
                mesajlar = hedef_klasor.Items
                mesajlar.Sort("[ReceivedTime]", True)  # en yeni once
                for mail in mesajlar:
                    try:
                        konu         = str(mail.Subject or "")
                        gonderen_addr= str(mail.SenderEmailAddress or "")
                        govde        = str(mail.Body or "")
                        # Gonderen filtresi
                        if gonderen and gonderen.lower() not in gonderen_addr.lower():
                            continue
                        # Konu filtresi
                        if filtre and filtre.lower() not in konu.lower():
                            continue
                        kod = _kod_bul_regex(govde + "\n" + konu, regex)
                        if kod:
                            return kod
                        break  # en yeni mailde yok, bekle
                    except: continue
            except: pass
            time.sleep(5)
        return None  # timeout doldu
    except ImportError:
        pass  # pywin32 yok
    except Exception:
        pass  # Outlook acik degil vs.

    return None  # Outlook bulunamadi veya kod gelmedi


def github_push(cfg, dosya_yolu=None):
    """Mevcut sap_suite_v2.py dosyasini GitHub'a pushlar."""
    import urllib.request, base64, json as _json
    token = cfg.get("github_token", "").strip()
    repo  = cfg.get("github_repo",  "Elcihad/Cloude").strip()
    fname = cfg.get("github_file",  "sap_suite_v2.py").strip()
    if not token:
        return False, "GitHub token ayarlanmamis!"
    # Kaynak dosya: bu scriptin kendisi
    if dosya_yolu is None:
        dosya_yolu = __file__
    try:
        with open(dosya_yolu, 'rb') as f:
            icerik = f.read()
    except Exception as e:
        return False, f"Dosya okunamadi: {e}"
    b64 = base64.b64encode(icerik).decode('ascii')
    api_url = f"https://api.github.com/repos/{repo}/contents/{fname}"
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'SAP-Suite-v2'
    }
    # Mevcut SHA'yi al
    try:
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as r:
            mevcut = _json.loads(r.read())
        sha = mevcut.get("sha", "")
        if not sha:
            return False, f"SHA bos. Repo:{repo} Dosya:{fname}"
    except urllib.error.HTTPError as e:
        hata = e.read().decode()[:200]
        return False, f"SHA HTTP {e.code}\nRepo:{repo}\nDosya:{fname}\n{hata}"
    except Exception as e:
        return False, f"SHA alinamadi: {e}\nRepo:{repo} Dosya:{fname}"
    # PUT ile guncelle
    payload = _json.dumps({
        "message": "SAP Suite v2 guncellendi (otomatik)",
        "content": b64,
        "sha": sha
    }).encode('utf-8')
    try:
        req = urllib.request.Request(api_url, data=payload, method='PUT', headers=headers)
        with urllib.request.urlopen(req, timeout=30) as r:
            _json.loads(r.read())
        return True, "GitHub'a basariyla yuklendi!"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode()[:200]}"
    except Exception as e:
        return False, f"Hata: {e}"


def load_config():
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, indent=2, ensure_ascii=False), encoding="utf-8")
    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        # Bozuk config — sessizce yutma, yedekle ve kullanıcıya görünür yap
        try:
            yedek = CONFIG_FILE.with_suffix(f".bozuk_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            shutil.copy(str(CONFIG_FILE), str(yedek))
            print(f"[load_config] Bozuk config yedeklendi: {yedek}  ({e})")
        except Exception:
            pass
        cfg = {}
    for k, v in DEFAULT_CONFIG.items(): cfg.setdefault(k, v)
    save_config(cfg); return cfg

def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

# ── İnsan Gibi Mouse Yardımcı ─────────────────────────────────────────────────
def insan_gibi_hareket(x1, y1, x2, y2, sure=0.5):
    """Bezier eğrisi ile insan gibi mouse hareketi."""
    try:
        import pyautogui
        adim = max(10, int(sure * 30))
        # Rastgele kontrol noktaları (titreme/eğri)
        cx1 = x1 + random.randint(-80, 80)
        cy1 = y1 + random.randint(-80, 80)
        cx2 = x2 + random.randint(-60, 60)
        cy2 = y2 + random.randint(-60, 60)
        for i in range(adim + 1):
            t = i / adim
            # Kübik Bezier
            bx = (1-t)**3*x1 + 3*(1-t)**2*t*cx1 + 3*(1-t)*t**2*cx2 + t**3*x2
            by = (1-t)**3*y1 + 3*(1-t)**2*t*cy1 + 3*(1-t)*t**2*cy2 + t**3*y2
            # Hız değişimi (başlangıç ve sonda yavaş, ortada hızlı)
            bx += random.uniform(-1.5, 1.5)
            by += random.uniform(-1.5, 1.5)
            pyautogui.moveTo(int(bx), int(by), duration=0)
            time.sleep(sure / adim * (0.5 + abs(math.sin(math.pi * t))))
    except Exception:
        pass

def insan_gibi_tikla(x, y, cfg):
    try:
        import pyautogui
        mevcut_x, mevcut_y = pyautogui.position()
        sure = random.uniform(cfg.get("mouse_hiz_min",0.05), cfg.get("mouse_hiz_max",0.15))
        insan_gibi_hareket(mevcut_x, mevcut_y, x, y, sure)
        time.sleep(random.uniform(cfg.get("tiklama_oncesi_bekleme_min",0.03),cfg.get("tiklama_oncesi_bekleme_max",0.08)))
        pyautogui.click(x, y)
    except Exception:
        pass

def sayfa_goruntu_bekle(goruntu_b64, timeout=12, esik=0.80, log_cb=None):
    """
    Ekranda şablon görüntüyü arar, bulunca True döner.
    log_cb: verilirse debug/hata mesajlarını loga yazar (None ise sessiz).
    """
    def _log(msg, level="INFO"):
        if log_cb:
            try: log_cb(msg, level)
            except Exception: pass

    if not goruntu_b64:
        _log("  📷  Görüntü tanımsız → bekleme atlanıyor", "WARN")
        return True   # tanımsızsa normal akışı kırma

    try:
        import base64
        arr = np.frombuffer(base64.b64decode(goruntu_b64), dtype=np.uint8)
        sablon_orig = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
        if sablon_orig is None:
            _log("  📷  Şablon decode edilemedi (base64 bozuk olabilir)", "ERROR")
            return False   # YALAN söyleme — gerçekten başarısız
        import mss
        with mss.mss() as sct:
            _ilk = cv2.cvtColor(np.array(sct.grab(sct.monitors[1])), cv2.COLOR_BGRA2GRAY)
        eh0, ew0 = _ilk.shape[:2]
        sh0, sw0 = sablon_orig.shape[:2]
        if sw0 > ew0 or sh0 > eh0:
            oran = min(ew0 / sw0, eh0 / sh0) * 0.95
            sablon = cv2.resize(sablon_orig, (int(sw0 * oran), int(sh0 * oran)))
        else:
            sablon = sablon_orig

        en_yuksek_skor = 0.0
        bitis = time.time() + timeout
        while time.time() < bitis:
            with mss.mss() as sct:
                ekran = cv2.cvtColor(np.array(sct.grab(sct.monitors[1])), cv2.COLOR_BGRA2GRAY)
            try:
                _, max_val, _, _ = cv2.minMaxLoc(
                    cv2.matchTemplate(ekran, sablon, cv2.TM_CCOEFF_NORMED)
                )
                if max_val > en_yuksek_skor:
                    en_yuksek_skor = max_val
                if max_val >= esik: return True
            except cv2.error as ce:
                _log(f"  📷  OpenCV hata: {ce}", "ERROR")
            time.sleep(0.3)

        # Timeout — en yüksek skor ne kadardı, kullanıcıya göster (eşiği ayarlaması için)
        _log(
            f"  📷  Görüntü bulunamadı — en yüksek benzerlik: {en_yuksek_skor:.2f} (eşik: {esik})",
            "WARN"
        )
    except Exception as e:
        _log(f"  📷  Bekleme hatası: {e}", "ERROR")
    return False

def _bolge_goruntu_bekle(sablon_b64, bolge_rect, timeout=5, esik=0.80, log_cb=None):
    """
    Ekranın sadece bolge_rect=[x,y,w,h] bölgesini yakalayıp sablon_b64 ile karşılaştırır.
    Eşleşme bulunursa ekran koordinatında (mx, my) merkez döndürür, bulunamazsa None.
    """
    def _log(msg, level="INFO"):
        if log_cb:
            try: log_cb(msg, level)
            except Exception: pass

    if not sablon_b64 or not bolge_rect: return None
    try:
        arr    = np.frombuffer(base64.b64decode(sablon_b64), dtype=np.uint8)
        sablon = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
        if sablon is None:
            _log("  🎯  Bölge şablonu decode edilemedi", "ERROR")
            return None
        import mss
        bx, by, bw, bh = bolge_rect
        mon   = {"left": bx, "top": by, "width": bw, "height": bh}
        en_yuksek = 0.0
        bitis = time.time() + timeout
        while time.time() < bitis:
            with mss.mss() as sct:
                shot  = sct.grab(mon)
                bolge = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2GRAY)
            sh, sw = sablon.shape[:2]; gh, gw = bolge.shape[:2]
            if sw > gw or sh > gh:
                oran   = min(gw / sw, gh / sh) * 0.95
                sablon = cv2.resize(sablon, (int(sw * oran), int(sh * oran)))
                sh, sw = sablon.shape[:2]
            try:
                sonuc = cv2.matchTemplate(bolge, sablon, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(sonuc)
                if max_val > en_yuksek: en_yuksek = max_val
                if max_val >= esik:
                    # Eşleşme merkezi → ekran koordinatına çevir
                    mx = bx + max_loc[0] + sw // 2
                    my = by + max_loc[1] + sh // 2
                    return (mx, my)
            except cv2.error as ce:
                _log(f"  🎯  OpenCV hata: {ce}", "ERROR")
            time.sleep(0.15)
        _log(
            f"  🎯  Bölgede bulunamadı — en yüksek benzerlik: {en_yuksek:.2f} (eşik: {esik})",
            "WARN"
        )
    except Exception as e:
        _log(f"  🎯  Bölge bekleme hatası: {e}", "ERROR")
    return None

def insan_gibi_yaz(metin, cfg):
    """Metni panoya kopyalayıp Ctrl+V ile yapıştır.
    pyautogui.typewrite() sadece ASCII destekler; bu yöntem Türkçe karakter,
    nokta, boşluk gibi her karakteri doğru iletir."""
    try:
        import pyautogui, pyperclip
        pyperclip.copy(str(metin))
        # İnsan gibi küçük bir bekleme (yapıştırmadan önce)
        time.sleep(random.uniform(
            cfg.get("tiklama_oncesi_bekleme_min", 0.05),
            cfg.get("tiklama_oncesi_bekleme_max", 0.25)
        ))
        pyautogui.hotkey("ctrl", "v")
    except Exception:
        pass

# ── Canvas (Alan Çizim) ───────────────────────────────────────────────────────
class CanvasWidget(QLabel):
    alan_eklendi = pyqtSignal(dict)
    alan_secildi = pyqtSignal(int)
    alan_tasindi = pyqtSignal(int, list, list)  # idx, yeni_rect, yeni_merkez

    def __init__(self):
        super().__init__()
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumSize(500, 350)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setCursor(QCursor(Qt.CursorShape.CrossCursor))
        self._pixmap_orig = None
        self._display_pm  = None
        self._scale = 1.0
        self._offset = QPoint(0, 0)
        self.alanlar = []
        self._cizim_basl = None
        self._gecici_rect = None
        self._cizim_modu = True
        self._secili_idx = -1
        # Taşıma
        self._tasima_aktif = False
        self._tasima_basl_mouse = None
        self._tasima_basl_rect  = None

    def goruntu_yukle(self, cv_img):
        self._pixmap_orig = cv_img.copy()
        h, w = cv_img.shape[:2]
        rgb = cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGB)
        rgb = np.ascontiguousarray(rgb)          # bellek düzenini garanti et
        self._rgb_buf = rgb                      # GC'den korumak için referans tut
        qimg = QImage(rgb.data, w, h, w * 3, QImage.Format.Format_RGB888)
        self._display_pm = QPixmap.fromImage(qimg)
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._display_pm is None:
            p = QPainter(self)
            p.fillRect(self.rect(), QColor(24, 28, 39))
            p.setPen(QColor(42, 48, 72))
            p.setFont(QFont("Segoe UI", 12))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Görüntü yükle veya ekran görüntüsü al")
            p.end()
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        ww, wh = self.width(), self.height()
        pm_w, pm_h = self._display_pm.width(), self._display_pm.height()
        scale = min(ww/pm_w, wh/pm_h, 1.0)
        self._scale = scale
        dw, dh = int(pm_w*scale), int(pm_h*scale)
        ox, oy = (ww-dw)//2, (wh-dh)//2
        self._offset = QPoint(ox, oy)
        p.drawPixmap(ox, oy, dw, dh, self._display_pm)
        for i, alan in enumerate(self.alanlar):
            r = alan["rect"]
            x, y, w, h = ox+int(r[0]*scale), oy+int(r[1]*scale), int(r[2]*scale), int(r[3]*scale)
            renk_hex = EYLEMLER.get(alan.get("eylem", "sol_tikla"), {}).get("renk", "#aaaaaa")
            renk = QColor(renk_hex)
            secili = (i == self._secili_idx)
            p.setPen(QPen(renk, 3 if secili else 2))
            p.setBrush(QColor(renk.red(), renk.green(), renk.blue(), 40 if secili else 20))
            p.drawRect(x, y, w, h)
            if secili:
                p.setBrush(renk); p.setPen(Qt.PenStyle.NoPen)
                for cx, cy in [(x,y),(x+w,y),(x,y+h),(x+w,y+h)]:
                    p.drawRect(cx-4, cy-4, 8, 8)
            icon = EYLEMLER.get(alan.get("eylem","sol_tikla"), {}).get("icon", "?")
            etiket = f" {icon} {alan['isim']} "
            p.setPen(QPen(renk)); p.setFont(QFont("Consolas", 9, QFont.Weight.Bold))
            p.setBrush(QColor(15, 17, 23, 200))
            fm = p.fontMetrics(); tw = fm.horizontalAdvance(etiket); th = fm.height()
            p.drawRect(x, y-th-2, tw+4, th+4)
            p.setPen(renk); p.drawText(x+2, y-4, etiket)
        if self._gecici_rect:
            p.setPen(QPen(QColor("#00d4aa"), 2, Qt.PenStyle.DashLine))
            p.setBrush(QColor(0, 212, 170, 30))
            p.drawRect(self._gecici_rect)
        p.end()

    def _w2g(self, pos):
        if self._display_pm is None: return None
        ox, oy = self._offset.x(), self._offset.y()
        pm_w, pm_h = self._display_pm.width(), self._display_pm.height()
        dw, dh = int(pm_w*self._scale), int(pm_h*self._scale)
        if 0 <= pos.x()-ox <= dw and 0 <= pos.y()-oy <= dh:
            return QPoint(int((pos.x()-ox)/self._scale), int((pos.y()-oy)/self._scale))
        return None

    def mousePressEvent(self, ev):
        if ev.button() != Qt.MouseButton.LeftButton or self._display_pm is None: return
        gp = self._w2g(ev.pos())
        if not gp: return
        if self._cizim_modu:
            self._cizim_basl = gp
        else:
            for i, alan in enumerate(self.alanlar):
                r = alan["rect"]
                if r[0] <= gp.x() <= r[0]+r[2] and r[1] <= gp.y() <= r[1]+r[3]:
                    self._secili_idx = i
                    self.alan_secildi.emit(i)
                    self._tasima_aktif = True
                    self._tasima_basl_mouse = gp
                    self._tasima_basl_rect  = list(alan["rect"])
                    self.setCursor(QCursor(Qt.CursorShape.SizeAllCursor))
                    self.update(); return
            self._secili_idx = -1
            self._tasima_aktif = False
            self.update()

    def mouseMoveEvent(self, ev):
        if self._display_pm is None: return
        gp = self._w2g(ev.pos())
        if not gp: return
        if self._cizim_modu and self._cizim_basl is not None:
            ox, oy = self._offset.x(), self._offset.y()
            x1 = ox+int(self._cizim_basl.x()*self._scale)
            y1 = oy+int(self._cizim_basl.y()*self._scale)
            self._gecici_rect = QRect(
                min(x1, ev.pos().x()), min(y1, ev.pos().y()),
                abs(ev.pos().x()-x1), abs(ev.pos().y()-y1)
            )
            self.update()
        elif not self._cizim_modu and self._tasima_aktif and self._secili_idx >= 0:
            dx = gp.x() - self._tasima_basl_mouse.x()
            dy = gp.y() - self._tasima_basl_mouse.y()
            br = self._tasima_basl_rect
            yeni_x = max(0, br[0] + dx)
            yeni_y = max(0, br[1] + dy)
            self.alanlar[self._secili_idx]["rect"]   = [yeni_x, yeni_y, br[2], br[3]]
            self.alanlar[self._secili_idx]["merkez"] = [yeni_x + br[2]//2, yeni_y + br[3]//2]
            self.update()
        elif not self._cizim_modu:
            uzerinde = any(
                a["rect"][0] <= gp.x() <= a["rect"][0]+a["rect"][2] and
                a["rect"][1] <= gp.y() <= a["rect"][1]+a["rect"][3]
                for a in self.alanlar
            )
            self.setCursor(QCursor(Qt.CursorShape.SizeAllCursor if uzerinde else Qt.CursorShape.ArrowCursor))

    def mouseReleaseEvent(self, ev):
        if ev.button() != Qt.MouseButton.LeftButton: return
        if self._cizim_modu and self._cizim_basl is not None:
            gp = self._w2g(ev.pos()); self._gecici_rect = None
            if gp and self._cizim_basl:
                x1, y1 = self._cizim_basl.x(), self._cizim_basl.y()
                x2, y2 = gp.x(), gp.y()
                w, h = abs(x2-x1), abs(y2-y1)
                if w > 5 and h > 5:
                    alan = {
                        "id": len(self.alanlar),
                        "eylem": "sol_tikla",
                        "rect": [min(x1,x2), min(y1,y2), w, h],
                        "isim": f"Alan_{len(self.alanlar)+1}",
                        "merkez": [min(x1,x2)+w//2, min(y1,y2)+h//2],
                        "params": {}
                    }
                    self.alanlar.append(alan)
                    self.alan_eklendi.emit(alan)
            self._cizim_basl = None; self.update()
        elif not self._cizim_modu and self._tasima_aktif:
            if self._secili_idx >= 0:
                alan = self.alanlar[self._secili_idx]
                self.alan_tasindi.emit(self._secili_idx, list(alan["rect"]), list(alan["merkez"]))
            self._tasima_aktif = False
            self._tasima_basl_mouse = None
            self._tasima_basl_rect  = None
            self.setCursor(QCursor(Qt.CursorShape.ArrowCursor))
            self.update()

    def alan_sil(self, idx):
        if 0 <= idx < len(self.alanlar):
            self.alanlar.pop(idx); self._secili_idx = -1; self.update()

    def hepsini_temizle(self): self.alanlar.clear(); self._secili_idx = -1; self.update()
    def mod_degistir(self, cizim):
        self._cizim_modu = cizim
        self._tasima_aktif = False
        self.setCursor(QCursor(Qt.CursorShape.CrossCursor if cizim else Qt.CursorShape.ArrowCursor))
    def secili_idx(self): return self._secili_idx
    def sec(self, idx): self._secili_idx = idx; self.update()

# ── Tek Eylem Satırı Widget ───────────────────────────────────────────────────
class EylemSatiriWidget(QWidget):
    """Bir eylem + parametrelerini gösteren, silinebilir satır."""
    silindi = pyqtSignal(object)   # self

    def __init__(self, eylem_verisi=None, parent=None):
        super().__init__(parent)
        # eylem_verisi: {"eylem": "sol_tikla", "params": {...}}
        self._veri = eylem_verisi or {"eylem": "sol_tikla", "params": {}}
        self._kur()

    def _kur(self):
        lay = QHBoxLayout(self); lay.setContentsMargins(4, 2, 4, 2); lay.setSpacing(6)

        # Sıra göstergesi (dışarıdan set edilir)
        self.sirano_lbl = QLabel("1."); self.sirano_lbl.setFixedWidth(20)
        self.sirano_lbl.setStyleSheet(f"color:{C['dim']};font-size:11px;")
        lay.addWidget(self.sirano_lbl)

        # Eylem seçici
        self.eylem_combo = QComboBox(); self.eylem_combo.setFixedWidth(170)
        for key, meta in EYLEMLER.items():
            self.eylem_combo.addItem(meta["etiket"], key)
        mevcut = self._veri.get("eylem","sol_tikla")
        idx = list(EYLEMLER.keys()).index(mevcut) if mevcut in EYLEMLER else 0
        self.eylem_combo.setCurrentIndex(idx)
        lay.addWidget(self.eylem_combo)

        # Parametre alanı (dinamik)
        self.param_container = QWidget(); self.param_lay = QHBoxLayout(self.param_container)
        self.param_lay.setContentsMargins(0,0,0,0); self.param_lay.setSpacing(4)
        lay.addWidget(self.param_container, 1)

        # Sil butonu
        btn_sil = QPushButton("✕"); btn_sil.setFixedSize(24, 24)
        btn_sil.setObjectName("danger"); btn_sil.clicked.connect(lambda: self.silindi.emit(self))
        lay.addWidget(btn_sil)

        self.eylem_combo.currentIndexChanged.connect(self._eylem_degisti)
        self._eylem_degisti()

    def _temizle_params(self):
        for i in reversed(range(self.param_lay.count())):
            item = self.param_lay.itemAt(i)
            if item and item.widget(): item.widget().deleteLater()

    def _eylem_degisti(self):
        self._temizle_params()
        eylem = self.eylem_combo.currentData()
        params = self._veri.get("params", {})

        if eylem == "metin_yaz":
            self.kaynak_combo = QComboBox(); self.kaynak_combo.setFixedWidth(160)
            for key, label in METIN_KAYNAKLARI.items():
                self.kaynak_combo.addItem(label, key)
            mevcut_k = params.get("kaynak","sabit")
            kidx = list(METIN_KAYNAKLARI.keys()).index(mevcut_k) if mevcut_k in METIN_KAYNAKLARI else 0
            self.kaynak_combo.setCurrentIndex(kidx)
            self.param_lay.addWidget(self.kaynak_combo)
            self.sabit_edit = QLineEdit(params.get("sabit_metin",""))
            self.sabit_edit.setPlaceholderText("Sabit metin...")
            self.sabit_edit.setVisible(mevcut_k == "sabit")
            self.param_lay.addWidget(self.sabit_edit)
            self.kaynak_combo.currentIndexChanged.connect(
                lambda: self.sabit_edit.setVisible(self.kaynak_combo.currentData()=="sabit")
            )

        elif eylem == "bekle":
            self.param_lay.addWidget(QLabel("sn:"))
            self.bekle_spin = QDoubleSpinBox(); self.bekle_spin.setRange(0.1,60.0)
            self.bekle_spin.setSingleStep(0.5); self.bekle_spin.setFixedWidth(70)
            self.bekle_spin.setValue(params.get("sure",1.0))
            self.param_lay.addWidget(self.bekle_spin)

        elif eylem == "mouse_hareket":
            self.param_lay.addWidget(QLabel("X:"))
            self.mx_spin = QSpinBox(); self.mx_spin.setRange(0,9999); self.mx_spin.setFixedWidth(60)
            self.mx_spin.setValue(params.get("hedef_x",0)); self.param_lay.addWidget(self.mx_spin)
            self.param_lay.addWidget(QLabel("Y:"))
            self.my_spin = QSpinBox(); self.my_spin.setRange(0,9999); self.my_spin.setFixedWidth(60)
            self.my_spin.setValue(params.get("hedef_y",0)); self.param_lay.addWidget(self.my_spin)
            self.param_lay.addWidget(QLabel("hız:"))
            self.mhiz_spin = QDoubleSpinBox(); self.mhiz_spin.setRange(0.1,5.0); self.mhiz_spin.setSingleStep(0.1)
            self.mhiz_spin.setFixedWidth(60); self.mhiz_spin.setValue(params.get("hiz",0.5))
            self.param_lay.addWidget(self.mhiz_spin)

        elif eylem == "fonksiyon_tusu":
            self.fn_edit = QLineEdit(params.get("tus","f5"))
            self.fn_edit.setFixedWidth(60)
            self.fn_edit.setPlaceholderText("f1-f12")
            self.param_lay.addWidget(self.fn_edit)
        elif eylem == "klavye_kisayol":
            self.kisayol_edit = QLineEdit(params.get("tuslar","")); self.kisayol_edit.setFixedWidth(120)
            self.kisayol_edit.setPlaceholderText("ctrl+a")
            self.param_lay.addWidget(self.kisayol_edit)

        elif eylem == "url_git":
            self.url_edit_e = QLineEdit(params.get("url","")); self.url_edit_e.setFixedWidth(300)
            self.url_edit_e.setPlaceholderText("https://... (bos=config URL)")
            self.param_lay.addWidget(self.url_edit_e)

        elif eylem == "chrome_ac":
            self.chrome_url_edit = QLineEdit(params.get("url","")); self.chrome_url_edit.setFixedWidth(300)
            self.chrome_url_edit.setPlaceholderText("https://... (bos=config URL)")
            self.param_lay.addWidget(self.chrome_url_edit)

        elif eylem == "dosya_bekle":
            self.param_lay.addWidget(QLabel("timeout(sn):"))
            self.dosya_timeout = QSpinBox(); self.dosya_timeout.setRange(5,300)
            self.dosya_timeout.setValue(params.get("timeout",60)); self.dosya_timeout.setFixedWidth(70)
            self.param_lay.addWidget(self.dosya_timeout)

        elif eylem == "tus_ve_goruntu":
            # Tuş seçici
            self.param_lay.addWidget(QLabel("Tuş:"))
            self.tvg_tus = QLineEdit(params.get("tus","f5")); self.tvg_tus.setFixedWidth(50)
            self.tvg_tus.setPlaceholderText("f5")
            self.param_lay.addWidget(self.tvg_tus)
            # Bekleme arası
            self.param_lay.addWidget(QLabel("Tekrar(sn):"))
            self.tvg_tekrar = QDoubleSpinBox(); self.tvg_tekrar.setRange(0.5,30.0)
            self.tvg_tekrar.setSingleStep(0.5); self.tvg_tekrar.setFixedWidth(60)
            self.tvg_tekrar.setValue(params.get("tekrar_sure",3.0))
            self.param_lay.addWidget(self.tvg_tekrar)
            # Max deneme
            self.param_lay.addWidget(QLabel("Maks:"))
            self.tvg_maks = QSpinBox(); self.tvg_maks.setRange(1,30)
            self.tvg_maks.setValue(params.get("maks_deneme",10)); self.tvg_maks.setFixedWidth(50)
            self.param_lay.addWidget(self.tvg_maks)
            # Bölge + görüntü durumu
            self.tvg_goruntu_b64  = params.get("goruntu_b64","")
            self.tvg_bolge_b64    = params.get("bolge_b64","")
            self.tvg_bolge_rect   = params.get("bolge_rect", None)  # [x,y,w,h] ekran koordinatı
            # Durum etiketi: bölge mi tam ekran mı
            _bolge_ok = bool(self.tvg_bolge_b64 and self.tvg_bolge_rect)
            _tam_ok   = bool(self.tvg_goruntu_b64)
            if _bolge_ok:
                _lbl_txt = "🎯 Bölge Seçildi"; _lbl_clr = "#00d4aa"
            elif _tam_ok:
                _lbl_txt = "🖼 Tam Ekran";     _lbl_clr = "#ffa502"
            else:
                _lbl_txt = "Görüntü Yok";       _lbl_clr = "#ff4757"
            self.tvg_goruntu_lbl = QLabel(_lbl_txt)
            self.tvg_goruntu_lbl.setStyleSheet(f"color:{_lbl_clr};font-size:10px;")
            self.param_lay.addWidget(self.tvg_goruntu_lbl)
            # Bölge seç butonu (öncelikli — tavsiye edilen yöntem)
            btn_bolge = QPushButton("🎯"); btn_bolge.setFixedSize(28,24)
            btn_bolge.setToolTip("Ekran görüntüsü al → kritik bölgeyi çiz → sadece o bölgeyi izle (hızlı ve kesin)")
            btn_bolge.clicked.connect(self._tvg_bolge_sec)
            self.param_lay.addWidget(btn_bolge)
            # Tam ekran butonu (yedek)
            btn_goruntu = QPushButton("📷"); btn_goruntu.setFixedSize(28,24)
            btn_goruntu.setToolTip("Tüm ekranı şablon olarak kaydet (yavaş, tavsiye edilmez)")
            btn_goruntu.clicked.connect(self._tvg_goruntu_sec)
            self.param_lay.addWidget(btn_goruntu)
            # Eşleşince tıkla checkbox
            self.tvg_tikla_chk = QCheckBox("Tıkla")
            self.tvg_tikla_chk.setChecked(params.get("eslesince_tikla", False))
            self.tvg_tikla_chk.setToolTip("Görüntü eşleşince bölgenin merkezine sol tıkla")
            self.param_lay.addWidget(self.tvg_tikla_chk)

        else:
            lbl = QLabel("—"); lbl.setStyleSheet(f"color:{C['dim']};")
            self.param_lay.addWidget(lbl)

        self.param_lay.addStretch()

    def get_veri(self):
        """Bu satırın {"eylem":..., "params":{...}} sözlüğünü döndür."""
        eylem = self.eylem_combo.currentData()
        params = {}
        if eylem == "metin_yaz":
            params["kaynak"] = self.kaynak_combo.currentData()
            params["sabit_metin"] = self.sabit_edit.text()
        elif eylem == "bekle":
            params["sure"] = self.bekle_spin.value()
        elif eylem == "mouse_hareket":
            params["hedef_x"] = self.mx_spin.value()
            params["hedef_y"] = self.my_spin.value()
            params["hiz"] = self.mhiz_spin.value()
        elif eylem == "klavye_kisayol":
            params["tuslar"] = self.kisayol_edit.text().strip()
        elif eylem == "fonksiyon_tusu":
            params["tus"] = self.fn_edit.text().strip().lower() if hasattr(self,"fn_edit") else "f5"
        elif eylem == "url_git":
            params["url"] = self.url_edit_e.text().strip()
        elif eylem == "chrome_ac":
            params["url"] = self.chrome_url_edit.text().strip()
        elif eylem == "dosya_bekle":
            params["timeout"] = self.dosya_timeout.value()
        elif eylem == "tus_ve_goruntu":
            params["tus"]          = self.tvg_tus.text().strip().lower() if hasattr(self,"tvg_tus") else "f5"
            params["tekrar_sure"]  = self.tvg_tekrar.value() if hasattr(self,"tvg_tekrar") else 3.0
            params["maks_deneme"]  = self.tvg_maks.value() if hasattr(self,"tvg_maks") else 10
            params["goruntu_b64"]  = self.tvg_goruntu_b64 if hasattr(self,"tvg_goruntu_b64") else ""
            params["bolge_b64"]    = self.tvg_bolge_b64   if hasattr(self,"tvg_bolge_b64")   else ""
            params["bolge_rect"]   = self.tvg_bolge_rect  if hasattr(self,"tvg_bolge_rect")  else None
            params["eslesince_tikla"] = self.tvg_tikla_chk.isChecked() if hasattr(self,"tvg_tikla_chk") else False
        return {"eylem": eylem, "params": params}

    def _tvg_goruntu_sec(self):
        """📷 Tam ekran şablon — menü aç (Dosyadan Seç / Ekran Al)."""
        menu = QMenu(self)
        menu.addAction("📂  Dosyadan Seç", self._tvg_dosyadan_sec)
        menu.addAction("🖥  Tüm Ekranı Al (yedek)", self._tvg_ekrandan_sec)
        menu.exec(self.cursor().pos())

    def _tvg_bolge_sec(self):
        """🎯 Bölge seç — ekran görüntüsü al, BolgeSecimDlg ile kritik bölgeyi çiz."""
        ust_pencere = self.window()
        # Tüm dialog ve ana pencereleri gizle
        gizlenen_dialoglar = []
        for w in QApplication.topLevelWidgets():
            if isinstance(w, QDialog) and w.isVisible():
                w.hide(); gizlenen_dialoglar.append(w)
        # Ana QMainWindow'u da bul ve minimize et (Chrome görünsün)
        ana_pencere = None
        for w in QApplication.topLevelWidgets():
            if isinstance(w, QMainWindow):
                ana_pencere = w; w.showMinimized(); break
        if ana_pencere is None:
            ust_pencere.hide()
        # 700ms sonra ekran görüntüsü al ve BolgeSecimDlg'yi aç
        QTimer.singleShot(
            700,
            lambda: self._tvg_bolge_yakala(ana_pencere, ust_pencere, gizlenen_dialoglar)
        )

    def _tvg_bolge_yakala(self, ana_pencere, ust_pencere, gizlenen_dialoglar):
        """Ekran görüntüsü al, sonra BolgeSecimDlg'yi aç."""
        tam_ekran = None
        try:
            import mss
            with mss.mss() as sct:
                shot = sct.grab(sct.monitors[1])
                tam_ekran = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)
        except Exception as e:
            QMessageBox.warning(ust_pencere, "Hata", f"Ekran alınamadı:\n{e}")

        # Pencereleri geri getir (BolgeSecimDlg bunların üstünde açılacak)
        if ana_pencere is not None:
            ana_pencere.showNormal()
        else:
            ust_pencere.show()
        for d in gizlenen_dialoglar:
            d.show()

        if tam_ekran is None:
            return

        # Bölge seçim dialogu aç
        dlg = BolgeSecimDlg(tam_ekran, ust_pencere)
        if dlg.exec():
            rect, bolge_img = dlg.secili_bolge()
            if bolge_img is not None and bolge_img.size > 0:
                # Sadece seçilen bölgeyi şablon olarak kaydet
                ok, buf = cv2.imencode(".png", bolge_img)
                if ok:
                    self.tvg_bolge_b64  = base64.b64encode(buf.tobytes()).decode("utf-8")
                    self.tvg_bolge_rect = rect  # [x, y, w, h] ekran koordinatı
                    # Tam ekranı da yedek olarak tut (fallback için)
                    ok2, buf2 = cv2.imencode(".png", tam_ekran)
                    if ok2:
                        self.tvg_goruntu_b64 = base64.b64encode(buf2.tobytes()).decode("utf-8")
                    self.tvg_goruntu_lbl.setText("🎯 Bölge Seçildi")
                    self.tvg_goruntu_lbl.setStyleSheet("color:#00d4aa;font-size:10px;")

    def _tvg_dosyadan_sec(self):
        yol,_ = QFileDialog.getOpenFileName(self,"Görüntü Seç","","Görüntüler (*.png *.jpg *.jpeg *.bmp)")
        if not yol: return
        img = cv2.imread(yol)
        if img is None:
            QMessageBox.warning(self, "Hata", "Görüntü okunamadı."); return
        ok, buf = cv2.imencode(".png", img)
        if ok:
            self.tvg_goruntu_b64 = base64.b64encode(buf.tobytes()).decode("utf-8")
            self.tvg_bolge_b64   = ""      # dosya seçilince bölge modunu sıfırla
            self.tvg_bolge_rect  = None
            self.tvg_goruntu_lbl.setText("🖼 Tam Ekran")
            self.tvg_goruntu_lbl.setStyleSheet("color:#ffa502;font-size:10px;")

    def _tvg_ekrandan_sec(self):
        """Tüm ekranı (yedek mod) şablon olarak kaydet — pencereleri gizle."""
        ust_pencere = self.window()
        gizlenen_dialoglar = []
        for w in QApplication.topLevelWidgets():
            if isinstance(w, QDialog) and w.isVisible():
                w.hide(); gizlenen_dialoglar.append(w)
        ana_pencere = None
        for w in QApplication.topLevelWidgets():
            if isinstance(w, QMainWindow):
                ana_pencere = w; w.showMinimized(); break
        if ana_pencere is None:
            ust_pencere.hide()
        QTimer.singleShot(
            700,
            lambda: self._tvg_ekran_yakala(ana_pencere, ust_pencere, gizlenen_dialoglar)
        )

    def _tvg_ekran_yakala(self, ana_pencere, ust_pencere, gizlenen_dialoglar):
        try:
            import mss
            with mss.mss() as sct:
                shot = sct.grab(sct.monitors[1])
                img = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)
            ok, buf = cv2.imencode(".png", img)
            if ok:
                self.tvg_goruntu_b64 = base64.b64encode(buf.tobytes()).decode("utf-8")
                self.tvg_bolge_b64   = ""     # tam ekran modunda bölgeyi sıfırla
                self.tvg_bolge_rect  = None
                self.tvg_goruntu_lbl.setText("🖼 Tam Ekran")
                self.tvg_goruntu_lbl.setStyleSheet("color:#ffa502;font-size:10px;")
            else:
                QMessageBox.warning(ust_pencere, "Hata", "Görüntü kodlanamadı")
        except Exception as e:
            QMessageBox.warning(ust_pencere, "Hata", f"Ekran görüntüsü alınamadı:\n{e}")
        finally:
            if ana_pencere is not None:
                ana_pencere.showNormal()
            else:
                ust_pencere.show()
            for d in gizlenen_dialoglar:
                d.show()

    def set_sirano(self, n):
        self.sirano_lbl.setText(f"{n}.")


# ── Bölge Seçim Diyalogu (tus_ve_goruntu için) ───────────────────────────────
class BolgeSecimDlg(QDialog):
    """
    Tam ekran görüntüsü üzerinde kullanıcının dikdörtgen çizmesini sağlar.
    Sadece seçilen bölge izleme şablonu olarak kaydedilir.
    """
    def __init__(self, cv_img, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Kritik Bölgeyi Çiz — İzlenecek Alan")
        self.setStyleSheet(SS)
        self._cv_img   = cv_img
        self._rect     = None
        self._basl     = None
        self._scale    = 1.0
        self._offset_x = 0
        self._offset_y = 0
        self._kur()

    def _kur(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8,8,8,8); lay.setSpacing(6)

        bilgi = QLabel("🎯  İzlenecek kritik bölgeyi çiz  (örn: sayfa başlığı, buton, durum yazısı)")
        bilgi.setStyleSheet(f"color:{C['accent']};font-size:11px;font-weight:bold;")
        lay.addWidget(bilgi)

        self._lbl = QLabel()
        self._lbl.setMinimumSize(900, 540)
        self._lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._lbl.setCursor(QCursor(Qt.CursorShape.CrossCursor))
        self._lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl.mousePressEvent   = self._mouse_press
        self._lbl.mouseMoveEvent    = self._mouse_move
        self._lbl.mouseReleaseEvent = self._mouse_release
        lay.addWidget(self._lbl)

        self._durum_lbl = QLabel("Sol tık basılı tut → sürükle → bırak")
        self._durum_lbl.setStyleSheet(f"color:{C['dim']};font-size:10px;")
        lay.addWidget(self._durum_lbl)

        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("✓  Bu Bölgeyi Kullan")
        bb.button(QDialogButtonBox.StandardButton.Ok).setObjectName("accent")
        bb.accepted.connect(self._kabul)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

        self._pixmap_guncelle()

    def _pixmap_guncelle(self, gecici_rect=None):
        img = self._cv_img.copy()
        if self._rect:
            x,y,w,h = self._rect
            cv2.rectangle(img, (x,y), (x+w,y+h), (0,212,170), 3)
        if gecici_rect:
            x1,y1,x2,y2 = gecici_rect
            cv2.rectangle(img, (min(x1,x2),min(y1,y2)), (max(x1,x2),max(y1,y2)), (255,165,0), 2)
        h, w = img.shape[:2]
        lw, lh = self._lbl.width() or 900, self._lbl.height() or 540
        scale = min(lw/w, lh/h, 1.0)
        self._scale    = scale
        self._offset_x = int((lw - w*scale)//2)
        self._offset_y = int((lh - h*scale)//2)
        dw, dh = int(w*scale), int(h*scale)
        rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        rgb = np.ascontiguousarray(rgb)
        qimg = QImage(rgb.data, w, h, w*3, QImage.Format.Format_RGB888)
        pm = QPixmap.fromImage(qimg).scaled(
            dw, dh, Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        self._lbl.setPixmap(pm)

    def _lbl2img(self, pos):
        ix = int((pos.x() - self._offset_x) / self._scale)
        iy = int((pos.y() - self._offset_y) / self._scale)
        h, w = self._cv_img.shape[:2]
        return max(0, min(ix, w-1)), max(0, min(iy, h-1))

    def _mouse_press(self, ev):
        if ev.button() == Qt.MouseButton.LeftButton:
            self._basl = self._lbl2img(ev.pos())

    def _mouse_move(self, ev):
        if self._basl:
            bitis = self._lbl2img(ev.pos())
            self._pixmap_guncelle(gecici_rect=(*self._basl, *bitis))

    def _mouse_release(self, ev):
        if ev.button() == Qt.MouseButton.LeftButton and self._basl:
            bitis = self._lbl2img(ev.pos())
            x1,y1 = self._basl; x2,y2 = bitis
            x,y,w,h = min(x1,x2), min(y1,y2), abs(x2-x1), abs(y2-y1)
            if w > 10 and h > 10:
                self._rect = [x, y, w, h]
                self._durum_lbl.setText(f"✓ Bölge: X={x} Y={y}  {w}×{h}px  — OK ile onayla")
                self._durum_lbl.setStyleSheet(f"color:{C['accent']};font-size:10px;font-weight:bold;")
            self._basl = None
            self._pixmap_guncelle()

    def _kabul(self):
        if not self._rect:
            QMessageBox.warning(self, "Uyarı", "Önce bir bölge çizin!"); return
        self.accept()

    def secili_bolge(self):
        if not self._rect: return None, None
        x,y,w,h = self._rect
        return self._rect, self._cv_img[y:y+h, x:x+w]


# ── Eylem Zinciri Diyalogu ────────────────────────────────────────────────────
class EylemDialog(QDialog):
    """
    Bir alana birden fazla eylem atamanın sağlandığı diyalog.
    Örnek:  1. Sol Tıkla
            2. Metin Yaz → Config: Kullanıcı Adı
            3. Tab'a Bas
    """
    def __init__(self, alan, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Eylem Zinciri Ayarları")
        self.setMinimumWidth(680)
        self.setMinimumHeight(340)
        self.setStyleSheet(SS)
        self.alan = alan
        self._satirlar = []   # EylemSatiriWidget listesi
        self._kur()

    def _kur(self):
        lay = QVBoxLayout(self); lay.setSpacing(6)

        # Alan ismi
        isim_lay = QHBoxLayout()
        isim_lay.addWidget(QLabel("Alan İsmi:"))
        self.isim_edit = QLineEdit(self.alan.get("isim",""))
        isim_lay.addWidget(self.isim_edit); lay.addLayout(isim_lay)

        # Eylem listesi başlığı
        baslik = QLabel("Eylem Zinciri  (sırayla çalışır):")
        baslik.setStyleSheet(f"color:{C['accent']};font-weight:bold;font-size:11px;")
        lay.addWidget(baslik)

        # Scrollable eylem listesi
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._liste_widget = QWidget()
        self._liste_lay = QVBoxLayout(self._liste_widget)
        self._liste_lay.setContentsMargins(0,0,0,0); self._liste_lay.setSpacing(2)
        self._liste_lay.addStretch()
        scroll.setWidget(self._liste_widget)
        lay.addWidget(scroll)

        # Eylem ekle butonu
        btn_ekle = QPushButton("➕  Eylem Ekle"); btn_ekle.setFixedHeight(28)
        btn_ekle.clicked.connect(lambda: self._satir_ekle())
        lay.addWidget(btn_ekle)

        # Mevcut eylemleri yükle
        # Geriye dönük uyumluluk: eski format tek eylem, yeni format eylem_zinciri listesi
        mevcut = self.alan.get("eylem_zinciri", None)
        if mevcut is None:
            # Eski format: tek eylem varsa onu zincire çevir
            tek = self.alan.get("eylem","sol_tikla")
            mevcut = [{"eylem": tek, "params": self.alan.get("params",{})}]
        for e in mevcut:
            self._satir_ekle(e)

        # Hiç satır yoksa boş bir tane ekle
        if not self._satirlar:
            self._satir_ekle()

        # OK / İptal
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _satir_ekle(self, veri=None):
        satir = EylemSatiriWidget(veri, self)
        satir.silindi.connect(self._satir_sil)
        # Stretch'in hemen önüne ekle
        idx = self._liste_lay.count() - 1
        self._liste_lay.insertWidget(idx, satir)
        self._satirlar.append(satir)
        self._siralari_guncelle()

    def _satir_sil(self, satir):
        if len(self._satirlar) <= 1:
            QMessageBox.information(self, "Bilgi", "En az bir eylem olmalı."); return
        self._satirlar.remove(satir)
        self._liste_lay.removeWidget(satir)
        satir.deleteLater()
        self._siralari_guncelle()

    def _siralari_guncelle(self):
        for i, s in enumerate(self._satirlar):
            s.set_sirano(i+1)

    def get_sonuc(self):
        zincir = [s.get_veri() for s in self._satirlar]
        # İlk eylemin ikonu canvas etiketi için
        ilk = zincir[0] if zincir else {"eylem":"sol_tikla","params":{}}
        return {
            "isim": self.isim_edit.text().strip() or self.alan.get("isim","Alan"),
            "eylem": ilk["eylem"],        # canvas rengi / ikonu için
            "params": ilk["params"],      # geriye dönük uyumluluk
            "eylem_zinciri": zincir,
        }

# ── Sayfa Ekleme Diyalogu ─────────────────────────────────────────────────────
class SayfaEkleDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sayfa Grubu Ekle")
        self.setMinimumWidth(380)
        self.setStyleSheet(SS)
        lay = QVBoxLayout(self)
        lay.setSpacing(10)

        # İsim
        lay.addWidget(QLabel("Sayfa Adı:"))
        self._isim_edit = QLineEdit()
        self._isim_edit.setPlaceholderText("Örn: Sipariş Ara, Geri Dön...")
        lay.addWidget(self._isim_edit)

        # Tip seçimi
        lay.addWidget(QLabel("Sayfa Tipi:"))
        self._tip_combo = QComboBox()
        for key, meta in SAYFA_TIPLERI.items():
            self._tip_combo.addItem(meta["etiket"], key)
        lay.addWidget(self._tip_combo)

        # Açıklama etiketi
        self._aciklama_lbl = QLabel()
        self._aciklama_lbl.setStyleSheet(f"color:{C['dim']};font-size:10px;")
        self._aciklama_lbl.setWordWrap(True)
        lay.addWidget(self._aciklama_lbl)
        self._tip_combo.currentIndexChanged.connect(self._tip_degisti)
        self._tip_degisti()

        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._kabul); bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _tip_degisti(self):
        key = self._tip_combo.currentData()
        meta = SAYFA_TIPLERI.get(key, {})
        self._aciklama_lbl.setText(meta.get("aciklama",""))

    def _kabul(self):
        if not self._isim_edit.text().strip():
            QMessageBox.warning(self, "Uyarı", "Sayfa adı boş olamaz!"); return
        self.accept()

    def isim(self): return self._isim_edit.text().strip()
    def tip(self):  return self._tip_combo.currentData()


# ── Akış Ağacı Widget ─────────────────────────────────────────────────────────
class AkisAgaci(QTreeWidget):
    """
    Ağaç yapısı:
      📄 Sayfa Grubu
        └─ 🖱 Alan_1 (sol_tikla)
        └─ ⌨ kullanici_input (metin_yaz)
    """
    degisti = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._goruntu_yukle_cb = lambda img: None
        self._get_canvas_img_cb = lambda: None
        self._ekran_goruntu_cb = lambda: None
        self._ekran_goruntu_hedef = None
        self.setHeaderLabels(["Adım", "Eylem", "Parametre"])
        self.setColumnWidth(0, 180)
        self.setColumnWidth(1, 130)
        self.setColumnWidth(2, 180)
        self.setDragDropMode(QTreeWidget.DragDropMode.InternalMove)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._sag_tik_menu)

    def sayfa_ekle(self, isim=None, goruntu_b64=None, sayfa_tipi=None):
        if isim is None:
            # Tip + isim diyalogu
            dlg = SayfaEkleDialog(self)
            if not dlg.exec(): return None
            isim = dlg.isim()
            sayfa_tipi = dlg.tip()
        sayfa_tipi = sayfa_tipi or "giris"
        tip_meta = SAYFA_TIPLERI.get(sayfa_tipi, SAYFA_TIPLERI["giris"])
        item = QTreeWidgetItem(self, [""])
        item.setData(0, Qt.ItemDataRole.UserRole, {
            "tip": "sayfa", "isim": isim.strip(),
            "sayfa_tipi": sayfa_tipi,
            "goruntu_b64": goruntu_b64 or ""
        })
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsDropEnabled)
        item.setExpanded(True)
        font = QFont(); font.setBold(True); item.setFont(0, font)
        self._sayfa_item_guncelle(item)
        self.degisti.emit()
        return item

    def _sayfa_item_guncelle(self, item):
        """Item metnini ve rengini sayfa tipine göre ayarla."""
        veri = item.data(0, Qt.ItemDataRole.UserRole) or {}
        isim = veri.get("isim", "")
        sayfa_tipi = veri.get("sayfa_tipi", "giris")
        goruntu_b64 = veri.get("goruntu_b64", "")
        tip_meta = SAYFA_TIPLERI.get(sayfa_tipi, SAYFA_TIPLERI["giris"])
        ic_goruntu = "🖼 " if goruntu_b64 else ""
        item.setText(0, f"{tip_meta['etiket']}  {ic_goruntu}{isim}")
        item.setForeground(0, QColor(tip_meta["renk"]))

    def adim_ekle(self, alan_verisi, sayfa_item=None):
        if sayfa_item is None:
            if self.topLevelItemCount() == 0:
                sayfa_item = self.sayfa_ekle("Sayfa 1")
            else:
                sayfa_item = self.topLevelItem(self.topLevelItemCount()-1)

        # Geriye dönük uyumluluk: eski format → zincire çevir
        zincir = alan_verisi.get("eylem_zinciri", None)
        if zincir is None:
            zincir = [{"eylem": alan_verisi.get("eylem","sol_tikla"),
                       "params": alan_verisi.get("params",{})}]

        # Sütun 1: ikonları yan yana
        eylem_ozet = "  ".join(EYLEMLER.get(e["eylem"],{}).get("icon","?") for e in zincir)
        # Sütun 2: parametreler özeti
        param_ozet = self._zincir_param_str(zincir)

        child = QTreeWidgetItem(sayfa_item, [
            f"  {alan_verisi.get('isim','?')}",
            eylem_ozet,
            param_ozet
        ])
        data = {"tip": "adim", **alan_verisi, "eylem_zinciri": zincir}
        child.setData(0, Qt.ItemDataRole.UserRole, data)
        child.setFlags((child.flags() | Qt.ItemFlag.ItemIsDragEnabled) & ~Qt.ItemFlag.ItemIsDropEnabled)
        ilk_renk = EYLEMLER.get(zincir[0]["eylem"] if zincir else "sol_tikla", {}).get("renk","#aaa")
        child.setForeground(1, QColor(ilk_renk))
        self.degisti.emit()
        return child

    def _param_str(self, eylem, params):
        if eylem == "metin_yaz":
            k = params.get("kaynak","sabit")
            return METIN_KAYNAKLARI.get(k,"?") + (f': "{params.get("sabit_metin","")}"' if k=="sabit" else "")
        elif eylem == "bekle": return f'{params.get("sure",1.0)} sn'
        elif eylem == "mouse_hareket": return f'→({params.get("hedef_x",0)},{params.get("hedef_y",0)}) {params.get("hiz",0.5)}sn'
        elif eylem == "fonksiyon_tusu": return params.get("tus","f5").upper()
        elif eylem == "klavye_kisayol": return params.get("tuslar","")
        return ""

    def _zincir_param_str(self, zincir):
        parcalar = []
        for e in zincir:
            s = self._param_str(e["eylem"], e.get("params",{}))
            if s: parcalar.append(s)
        return " · ".join(parcalar) if parcalar else ""

    def _sag_tik_menu(self, pos):
        item = self.itemAt(pos)
        menu = QMenu(self)
        if item is None:
            menu.addAction("➕  Sayfa Grubu Ekle", self.sayfa_ekle)
        else:
            veri = item.data(0, Qt.ItemDataRole.UserRole) or {}
            if veri.get("tip") == "sayfa":
                menu.addAction("➕  Sayfa Grubu Ekle", self.sayfa_ekle)
                menu.addAction("✏  Yeniden Adlandır", lambda: self._sayfa_yeniden_adlandir(item))
                # Tip değiştir alt menüsü
                tip_menu = menu.addMenu("🔄  Tip Değiştir")
                for tkey, tmeta in SAYFA_TIPLERI.items():
                    akt = "✓ " if veri.get("sayfa_tipi") == tkey else "   "
                    tip_menu.addAction(f"{akt}{tmeta['etiket']}", lambda tk=tkey: self._sayfa_tip_degistir(item, tk))
                menu.addSeparator()
                menu.addAction("📷  Görüntü Ata (Dosya)", lambda: self._sayfa_goruntu_dosya(item))
                menu.addAction("📷  Görüntü Ata (Canvas'tan)", lambda: self.sayfa_canvas_goruntu_ata(item, self._get_canvas_img_cb()))
                menu.addAction("🖥  Ekran Görüntüsü Al", lambda: self._sayfa_ekran_goruntu_al(item))
                has_img = bool(veri.get("goruntu_b64",""))
                if has_img:
                    menu.addAction("🗑  Görüntüyü Kaldır", lambda: self._sayfa_goruntu_kaldir(item))
                menu.addSeparator()
                menu.addAction("🗑  Sayfayı Sil", lambda: self._sayfa_sil(item))
            elif veri.get("tip") == "adim":
                menu.addAction("✏  Düzenle", lambda: self._adim_duzenle(item))
                menu.addAction("⬆  Yukarı Taşı", lambda: self._adim_tasi(item, -1))
                menu.addAction("⬇  Aşağı Taşı", lambda: self._adim_tasi(item, 1))
                menu.addSeparator()
                menu.addAction("🗑  Adımı Sil", lambda: self._adim_sil(item))
        menu.exec(self.viewport().mapToGlobal(pos))

    # ── Görüntü işleme ────────────────────────────────────────────────────────
    def _cv_to_b64(self, cv_img):
        ok, buf = cv2.imencode(".png", cv_img)
        if not ok: return ""
        return base64.b64encode(buf.tobytes()).decode("utf-8")

    def b64_to_cv(self, b64str):
        try:
            raw = base64.b64decode(b64str)
            arr = np.frombuffer(raw, dtype=np.uint8)
            return cv2.imdecode(arr, cv2.IMREAD_COLOR)
        except Exception:
            return None

    def _sayfa_goruntu_guncelle(self, item, goruntu_b64):
        veri = item.data(0, Qt.ItemDataRole.UserRole)
        veri["goruntu_b64"] = goruntu_b64
        item.setData(0, Qt.ItemDataRole.UserRole, veri)
        self._sayfa_item_guncelle(item)
        self.degisti.emit()

    def _sayfa_goruntu_dosya(self, item):
        yol, _ = QFileDialog.getOpenFileName(
            self, "Sayfa Görüntüsü Seç", "",
            "Görüntüler (*.png *.jpg *.jpeg *.bmp)"
        )
        if not yol: return
        cv_img = cv2.imread(yol)
        if cv_img is None:
            QMessageBox.warning(self, "Hata", "Görüntü okunamadı."); return
        b64 = self._cv_to_b64(cv_img)
        self._sayfa_goruntu_guncelle(item, b64)
        self._goruntu_yukle_cb(cv_img)   # canvas'a da yükle

    def _sayfa_ekran_goruntu_al(self, item):
        """Ana pencereyi gizle, ekran görüntüsü al, sayfaya ata."""
        self._ekran_goruntu_hedef = item
        self._ekran_goruntu_cb()   # AlanTanitmaSekmesi'nin _ekran_al_icin_sayfa metodunu tetikle

    def ekran_goruntu_cb_kaydet(self, cb):
        """AlanTanitmaSekmesi'nden ekran görüntüsü alma callback'ini kaydet."""
        self._ekran_goruntu_cb = cb

    def _sayfa_goruntu_kaldir(self, item):
        self._sayfa_goruntu_guncelle(item, "")

    def goruntu_yukle_cb_kaydet(self, cb):
        """AlanTanitmaSekmesi'nden canvas yükleme callback'i alır."""
        self._goruntu_yukle_cb = cb

    def get_canvas_img_cb_kaydet(self, cb):
        """Canvas'taki mevcut cv_img'i döndüren callback'i kaydet."""
        self._get_canvas_img_cb = cb

    def sayfa_canvas_goruntu_ata(self, item, cv_img):
        """Canvas'taki mevcut görüntüyü bu sayfaya ata (dışarıdan çağrılır)."""
        if cv_img is None:
            QMessageBox.warning(self, "Uyarı", "Canvas'ta görüntü yok!"); return
        b64 = self._cv_to_b64(cv_img)
        self._sayfa_goruntu_guncelle(item, b64)

    def _sayfa_tip_degistir(self, item, yeni_tip):
        veri = item.data(0, Qt.ItemDataRole.UserRole)
        veri["sayfa_tipi"] = yeni_tip
        item.setData(0, Qt.ItemDataRole.UserRole, veri)
        self._sayfa_item_guncelle(item)
        self.degisti.emit()

    def _sayfa_yeniden_adlandir(self, item):
        mevcut = item.data(0, Qt.ItemDataRole.UserRole).get("isim","")
        isim, ok = QInputDialog.getText(self, "Yeniden Adlandır", "Yeni ad:", text=mevcut)
        if ok and isim.strip():
            veri = item.data(0, Qt.ItemDataRole.UserRole)
            veri["isim"] = isim.strip()
            item.setData(0, Qt.ItemDataRole.UserRole, veri)
            self._sayfa_item_guncelle(item)
            self.degisti.emit()

    def _sayfa_sil(self, item):
        if QMessageBox.question(self, "Sil", "Bu sayfa ve tüm adımlar silinsin mi?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            idx = self.indexOfTopLevelItem(item)
            self.takeTopLevelItem(idx)
            self.degisti.emit()

    def _adim_sil(self, item):
        parent = item.parent()
        if parent:
            parent.removeChild(item)
            self.degisti.emit()

    def _adim_duzenle(self, item):
        veri = item.data(0, Qt.ItemDataRole.UserRole) or {}
        dlg = EylemDialog(veri, self)
        if dlg.exec():
            sonuc = dlg.get_sonuc()
            veri.update(sonuc)
            item.setData(0, Qt.ItemDataRole.UserRole, veri)
            zincir = sonuc.get("eylem_zinciri", [])
            eylem_ozet = "  ".join(EYLEMLER.get(e["eylem"],{}).get("icon","?") for e in zincir)
            item.setText(0, f"  {sonuc['isim']}")
            item.setText(1, eylem_ozet)
            item.setText(2, self._zincir_param_str(zincir))
            ilk_renk = EYLEMLER.get(zincir[0]["eylem"] if zincir else "sol_tikla",{}).get("renk","#aaa")
            item.setForeground(1, QColor(ilk_renk))
            self.degisti.emit()

    def _adim_tasi(self, item, yon):
        parent = item.parent()
        if not parent: return
        idx = parent.indexOfChild(item)
        yeni = idx + yon
        if 0 <= yeni < parent.childCount():
            parent.takeChild(idx)
            parent.insertChild(yeni, item)
            self.setCurrentItem(item)
            self.degisti.emit()

    def akis_verisi(self):
        """JSON'a yazılacak yapıyı döndür."""
        sayfalar = []
        for i in range(self.topLevelItemCount()):
            sayfa_item = self.topLevelItem(i)
            sayfa_veri = sayfa_item.data(0, Qt.ItemDataRole.UserRole) or {}
            adimlar = []
            for j in range(sayfa_item.childCount()):
                adim_item = sayfa_item.child(j)
                adim_veri = adim_item.data(0, Qt.ItemDataRole.UserRole) or {}
                zincir = adim_veri.get("eylem_zinciri", None)
                if zincir is None:
                    zincir = [{"eylem": adim_veri.get("eylem","sol_tikla"),
                               "params": adim_veri.get("params",{})}]
                adimlar.append({
                    "isim":          adim_veri.get("isim",""),
                    "merkez":        adim_veri.get("merkez",[0,0]),
                    "rect":          adim_veri.get("rect",[0,0,0,0]),
                    "eylem_zinciri": zincir,
                })
            sayfalar.append({
                "isim": sayfa_veri.get("isim",""),
                "sayfa_tipi": sayfa_veri.get("sayfa_tipi","giris"),
                "goruntu_b64": sayfa_veri.get("goruntu_b64",""),
                "adimlar": adimlar
            })
        return sayfalar

    def akis_yukle(self, sayfalar):
        """JSON'dan akışı yükle."""
        self.clear()
        for sayfa in sayfalar:
            sayfa_item = self.sayfa_ekle(
                sayfa.get("isim","Sayfa"),
                sayfa.get("goruntu_b64",""),
                sayfa.get("sayfa_tipi","giris")
            )
            for adim in sayfa.get("adimlar", []):
                self.adim_ekle(adim, sayfa_item)

# ── Sekme 1: Alan Tanıtma ─────────────────────────────────────────────────────
class AlanTanitmaSekmesi(QWidget):
    def __init__(self, cfg):
        super().__init__()
        self.cfg = cfg
        self._cv_img = None
        self._aktif_sayfa_item = None
        self._kur()

    def _kur(self):
        ana = QHBoxLayout(self)
        ana.setContentsMargins(8, 8, 8, 8); ana.setSpacing(8)

        # ── Sol panel ──────────────────────────────────────────────────────
        sol = QWidget(); sol.setFixedWidth(220)
        sv = QVBoxLayout(sol); sv.setContentsMargins(0,0,0,0); sv.setSpacing(6)

        gb_kaynak = QGroupBox("Görüntü Kaynağı")
        vk = QVBoxLayout(gb_kaynak); vk.setSpacing(4)
        btn_dosya = QPushButton("📂  Dosya Aç"); btn_dosya.clicked.connect(self._dosya_ac)
        vk.addWidget(btn_dosya)
        btn_ekran = QPushButton("🖥  Ekran Görüntüsü"); btn_ekran.clicked.connect(self._ekran_al)
        vk.addWidget(btn_ekran)
        sv.addWidget(gb_kaynak)

        gb_mod = QGroupBox("Çizim Modu")
        vm = QVBoxLayout(gb_mod); vm.setSpacing(4)
        self.btn_ciz = QPushButton("✏  Alan Çiz"); self.btn_ciz.setObjectName("active")
        self.btn_ciz.clicked.connect(lambda: self._mod(True)); vm.addWidget(self.btn_ciz)
        self.btn_sec = QPushButton("🔍  Alan Seç")
        self.btn_sec.clicked.connect(lambda: self._mod(False)); vm.addWidget(self.btn_sec)
        sv.addWidget(gb_mod)

        # Sayfa hedefi seç
        gb_sayfa = QGroupBox("Alan Eklenecek Sayfa")
        vp = QVBoxLayout(gb_sayfa); vp.setSpacing(4)
        lbl_hint = QLabel("Ağaçta sayfaya sağ tık →\nSayfayı seç veya yeni ekle")
        lbl_hint.setStyleSheet(f"color:{C['dim']};font-size:10px;")
        lbl_hint.setWordWrap(True); vp.addWidget(lbl_hint)
        self.aktif_sayfa_lbl = QLabel("Son sayfaya ekle")
        self.aktif_sayfa_lbl.setStyleSheet(f"color:{C['accent']};font-size:11px;font-weight:bold;")
        self.aktif_sayfa_lbl.setWordWrap(True); vp.addWidget(self.aktif_sayfa_lbl)
        sv.addWidget(gb_sayfa)

        gb_ipucu = QGroupBox("Kullanım")
        vi = QVBoxLayout(gb_ipucu)
        ipucu = QLabel(
            "1. Görüntü aç\n"
            "2. Sağda sayfa grubu oluştur\n"
            "3. Sayfayı seç (ağaçta tıkla)\n"
            "4. Canvas'ta alan çiz\n"
            "5. Açılan diyalogdan eylem ata\n"
            "6. JSON Kaydet"
        )
        ipucu.setStyleSheet(f"color:{C['dim']};font-size:10px;"); ipucu.setWordWrap(True)
        vi.addWidget(ipucu); sv.addWidget(gb_ipucu)
        sv.addStretch(); ana.addWidget(sol)

        # ── Orta Canvas ────────────────────────────────────────────────────
        orta = QWidget(); ov = QVBoxLayout(orta); ov.setContentsMargins(0,0,0,0); ov.setSpacing(4)
        self.canvas = CanvasWidget()
        self.canvas.alan_eklendi.connect(self._alan_eklendi)
        self.canvas.alan_secildi.connect(self._canvas_alan_secildi)
        self.canvas.alan_tasindi.connect(self._canvas_alan_tasindi)
        frame = QFrame(); frame.setObjectName("canvas_frame")
        fl = QVBoxLayout(frame); fl.setContentsMargins(4,4,4,4); fl.addWidget(self.canvas)
        ov.addWidget(frame)
        alt = QHBoxLayout(); alt.setSpacing(6)
        btn_temizle = QPushButton("🗑  Canvas Temizle"); btn_temizle.setObjectName("danger")
        btn_temizle.clicked.connect(self._canvas_temizle); alt.addWidget(btn_temizle)
        alt.addStretch()
        btn_kaydet = QPushButton("💾  JSON Kaydet"); btn_kaydet.setObjectName("accent")
        btn_kaydet.clicked.connect(self._json_kaydet); alt.addWidget(btn_kaydet)
        btn_yukle = QPushButton("📥  JSON Yükle"); btn_yukle.clicked.connect(self._json_yukle)
        alt.addWidget(btn_yukle)
        btn_github = QPushButton("🐙  GitHub'a Gönder"); btn_github.clicked.connect(self._github_push)
        alt.addWidget(btn_github)
        ov.addLayout(alt); ana.addWidget(orta)

        # ── Sağ panel: Excel / PDF sekmeleri ──────────────────────────────
        sag = QWidget(); sag.setFixedWidth(480)
        sagv = QVBoxLayout(sag); sagv.setContentsMargins(0,0,0,0); sagv.setSpacing(4)

        self.akis_tab = QTabWidget()

        def _agac_sekme_olustur(tab_label, agac_attr):
            w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(4,4,4,4); wv.setSpacing(4)
            tb = QHBoxLayout(); tb.setSpacing(4)
            bs = QPushButton("➕ Sayfa"); bs.setFixedHeight(28)
            bs.clicked.connect(self._yeni_sayfa); tb.addWidget(bs)
            be = QPushButton("⚡ Eylem Ekle"); be.setFixedHeight(28)
            be.setToolTip("Alan çizmeden doğrudan sayfaya eylem ekle (F5, Enter, Bekle vb.)")
            be.clicked.connect(self._koordinatsiz_eylem_ekle); tb.addWidget(be)
            bu = QPushButton("⬆"); bu.setFixedSize(28,28)
            bu.clicked.connect(self._adim_yukari); tb.addWidget(bu)
            bd = QPushButton("⬇"); bd.setFixedSize(28,28)
            bd.clicked.connect(self._adim_asagi); tb.addWidget(bd)
            tb.addStretch()
            bx = QPushButton("🗑"); bx.setFixedSize(28,28); bx.setObjectName("danger")
            bx.clicked.connect(self._secili_sil); tb.addWidget(bx)
            wv.addLayout(tb)
            agac = AkisAgaci()
            agac.itemClicked.connect(self._agac_tikla)
            agac.goruntu_yukle_cb_kaydet(self._canvas_goruntu_yukle)
            agac.get_canvas_img_cb_kaydet(lambda: self._cv_img)
            agac.ekran_goruntu_cb_kaydet(self._ekran_al_icin_sayfa)
            setattr(self, agac_attr, agac)
            wv.addWidget(agac)
            self.akis_tab.addTab(w, tab_label)

        _agac_sekme_olustur("📊  Excel İndir", "akis_agaci")
        _agac_sekme_olustur("📄  PDF İndir",   "akis_agaci_pdf")

        self.akis_tab.currentChanged.connect(self._sekme_degisti)
        sagv.addWidget(self.akis_tab)

        # Koordinat bilgisi
        gb_koor = QGroupBox("Seçili Alan Bilgisi")
        vk2 = QVBoxLayout(gb_koor)
        self.koor_lbl = QLabel("—"); self.koor_lbl.setFont(QFont("Consolas",9))
        self.koor_lbl.setStyleSheet(f"color:{C['dim']};"); self.koor_lbl.setWordWrap(True)
        vk2.addWidget(self.koor_lbl); sagv.addWidget(gb_koor)

        ana.addWidget(sag)

    # ── Event handlers ────────────────────────────────────────────────────────
    def _mod(self, cizim):
        self.canvas.mod_degistir(cizim)
        self.btn_ciz.setObjectName("active" if cizim else "")
        self.btn_ciz.setStyle(self.btn_ciz.style())
        self.btn_sec.setObjectName("active" if not cizim else "")
        self.btn_sec.setStyle(self.btn_sec.style())

    def _ekran_al_icin_sayfa(self):
        """Sağ tık menüsünden tetiklenir — pencere gizlenir, ekran görüntüsü alınır."""
        self.window().hide()
        QTimer.singleShot(600, self._ekran_yakala_icin_sayfa)

    def _ekran_yakala_icin_sayfa(self):
        hedef_agac = self._aktif_agac()          # hangi sekme aktifse onu kullan
        hedef = hedef_agac._ekran_goruntu_hedef
        try:
            import mss
            with mss.mss() as sct:
                mon = sct.monitors[1]; shot = sct.grab(mon)
                cv_img = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)
            self._cv_img = cv_img
            self.canvas.goruntu_yukle(cv_img)
            if hedef is not None:
                b64 = hedef_agac._cv_to_b64(cv_img)
                hedef_agac._sayfa_goruntu_guncelle(hedef, b64)
                isim = hedef.data(0, Qt.ItemDataRole.UserRole).get("isim","")
                self.aktif_sayfa_lbl.setText(f"→ {isim} 🖼")
        except Exception as e:
            QMessageBox.warning(self, "Hata", str(e))
        finally:
            hedef_agac._ekran_goruntu_hedef = None
            self.window().show()

    def _aktif_sayfa_goruntu_guncelle(self):
        """Mevcut canvas görüntüsünü aktif sayfaya kaydet."""
        if self._aktif_sayfa_item is not None and self._cv_img is not None:
            b64 = self._aktif_agac()._cv_to_b64(self._cv_img)
            self._aktif_agac()._sayfa_goruntu_guncelle(self._aktif_sayfa_item, b64)
            isim = self._aktif_sayfa_item.data(0, Qt.ItemDataRole.UserRole).get("isim","")
            self.aktif_sayfa_lbl.setText(f"→ {isim} 🖼")

    def _dosya_ac(self):
        yol,_ = QFileDialog.getOpenFileName(self,"Görüntü Aç","","Görüntüler (*.png *.jpg *.jpeg *.bmp)")
        if not yol: return
        img = cv2.imread(yol)
        if img is None:
            QMessageBox.warning(self, "Hata", f"Görüntü okunamadı:\n{yol}"); return
        self._cv_img = img
        self.canvas.goruntu_yukle(self._cv_img)
        self._aktif_sayfa_goruntu_guncelle()

    def _ekran_al(self):
        self.window().hide(); QTimer.singleShot(600, self._ekran_yakala)

    def _ekran_yakala(self):
        try:
            import mss
            with mss.mss() as sct:
                mon = sct.monitors[1]; shot = sct.grab(mon)
                self._cv_img = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)
            self.canvas.goruntu_yukle(self._cv_img)
            self._aktif_sayfa_goruntu_guncelle()
        except Exception as e:
            QMessageBox.warning(self,"Hata",str(e))
        self.window().show()

    def _canvas_goruntu_yukle(self, cv_img):
        """Dışarıdan canvas'a görüntü yükle (AkisAgaci callback'i)."""
        if cv_img is not None:
            self._cv_img = cv_img
            self.canvas.goruntu_yukle(cv_img)

    def _yeni_sayfa(self):
        # Canvas'taki görüntüyü yeni sayfaya otomatik ata
        b64 = ""
        if self._cv_img is not None:
            b64 = self._aktif_agac()._cv_to_b64(self._cv_img)
        item = self._aktif_agac().sayfa_ekle(goruntu_b64=b64)
        if item:
            self._aktif_sayfa_item = item
            isim = item.data(0, Qt.ItemDataRole.UserRole).get("isim","")
            suffix = " 🖼" if b64 else ""
            self.aktif_sayfa_lbl.setText(f"→ {isim}{suffix}")

    def _agac_tikla(self, item, _):
        veri = item.data(0, Qt.ItemDataRole.UserRole) or {}
        if veri.get("tip") == "sayfa":
            self._aktif_sayfa_item = item
            isim = veri.get("isim","")
            b64 = veri.get("goruntu_b64","")
            suffix = " 🖼" if b64 else ""
            self.aktif_sayfa_lbl.setText(f"→ {isim}{suffix}")
            # Canvas'a sayfanın görüntüsünü yükle
            if b64:
                cv_img = self._aktif_agac().b64_to_cv(b64)
                if cv_img is not None:
                    self._cv_img = cv_img
                    self.canvas.goruntu_yukle(cv_img)
                    self.canvas.hepsini_temizle()
                    # O sayfanın adımlarını canvas'a tekrar çiz
                    for j in range(item.childCount()):
                        adim_item = item.child(j)
                        adim_veri = adim_item.data(0, Qt.ItemDataRole.UserRole) or {}
                        self.canvas.alanlar.append(adim_veri)
                    self.canvas.update()
        elif veri.get("tip") == "adim":
            r = veri.get("rect",[0,0,0,0])
            m = veri.get("merkez",[0,0])
            self.koor_lbl.setText(
                f"İsim: {veri.get('isim','')}\n"
                f"Eylem: {EYLEMLER.get(veri.get('eylem',''),{}).get('etiket','')}\n"
                f"X:{r[0]} Y:{r[1]}  G:{r[2]} Y:{r[3]}\n"
                f"Merkez: ({m[0]},{m[1]})"
            )

    def _koordinatsiz_eylem_ekle(self):
        """Alan çizmeden doğrudan sayfaya eylem ekle (F5, Enter, Bekle vb.)"""
        agac = self._aktif_agac()
        if self._aktif_sayfa_item is None:
            if agac.topLevelItemCount() == 0:
                QMessageBox.warning(self, "Uyarı", "Önce bir sayfa grubu oluşturun!"); return
            self._aktif_sayfa_item = agac.topLevelItem(agac.topLevelItemCount() - 1)
        alan = {
            "id": -1,
            "eylem": "f5",
            "rect": [0, 0, 0, 0],
            "isim": "Eylem",
            "merkez": [0, 0],
            "params": {}
        }
        dlg = EylemDialog(alan, self)
        if dlg.exec():
            sonuc = dlg.get_sonuc()
            alan.update(sonuc)
            agac.adim_ekle(alan, self._aktif_sayfa_item)

    def _alan_eklendi(self, alan):
        # Eylem diyalogu aç
        dlg = EylemDialog(alan, self)
        if dlg.exec():
            sonuc = dlg.get_sonuc()
            alan.update(sonuc)
        # Ağaca ekle
        self._aktif_agac().adim_ekle(alan, self._aktif_sayfa_item)
        # Canvas'ta ismi güncelle
        idx = len(self.canvas.alanlar) - 1
        if idx >= 0:
            self.canvas.alanlar[idx].update(alan)
            self.canvas.sec(idx); self.canvas.update()

    def _canvas_alan_secildi(self, idx):
        pass  # ağaçta highlight yapılabilir, şimdilik boş

    def _canvas_alan_tasindi(self, canvas_idx, yeni_rect, yeni_merkez):
        """Canvas'ta alan taşındığında ağaçtaki koordinatları güncelle."""
        if self._aktif_sayfa_item is None: return
        agac = self._aktif_agac()
        if 0 <= canvas_idx < self._aktif_sayfa_item.childCount():
            child = self._aktif_sayfa_item.child(canvas_idx)
            veri = child.data(0, Qt.ItemDataRole.UserRole) or {}
            veri["rect"]   = yeni_rect
            veri["merkez"] = yeni_merkez
            child.setData(0, Qt.ItemDataRole.UserRole, veri)
            agac.degisti.emit()
        self.koor_lbl.setText(
            f"Taşındı → X:{yeni_rect[0]} Y:{yeni_rect[1]}  G:{yeni_rect[2]} Y:{yeni_rect[3]}\n"
            f"Merkez: ({yeni_merkez[0]},{yeni_merkez[1]})"
        )


    def _aktif_agac(self):
        idx = self.akis_tab.currentIndex() if hasattr(self, "akis_tab") else 0
        return self.akis_agaci_pdf if idx == 1 else self.akis_agaci

    def _sekme_degisti(self, idx):
        agac = self._aktif_agac()
        agac.goruntu_yukle_cb_kaydet(self._canvas_goruntu_yukle)
        agac.get_canvas_img_cb_kaydet(lambda: self._cv_img)
        agac.ekran_goruntu_cb_kaydet(self._ekran_al_icin_sayfa)

    def _canvas_temizle(self):
        if QMessageBox.question(self,"Emin misin?","Canvas'taki çizimler silinsin mi?",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            self.canvas.hepsini_temizle()

    def _adim_yukari(self):
        agac = self._aktif_agac()
        item = agac.currentItem()
        if item and item.parent():
            agac._adim_tasi(item, -1)

    def _adim_asagi(self):
        agac = self._aktif_agac()
        item = agac.currentItem()
        if item and item.parent():
            agac._adim_tasi(item, 1)

    def _secili_sil(self):
        agac = self._aktif_agac()
        item = agac.currentItem()
        if not item: return
        veri = item.data(0,Qt.ItemDataRole.UserRole) or {}
        if veri.get("tip") == "sayfa":
            agac._sayfa_sil(item)
        elif veri.get("tip") == "adim":
            agac._adim_sil(item)

    def _json_kaydet(self):
        agac  = self._aktif_agac()
        mod   = "pdf" if self.akis_tab.currentIndex() == 1 else "excel"
        etiket = "PDF" if mod == "pdf" else "Excel"
        sayfalar = agac.akis_verisi()
        if not sayfalar:
            QMessageBox.information(self,"Bilgi",f"{etiket} akisinda kaydedilecek sayfa yok!"); return
        cfg_key = "akis_json_pdf" if mod == "pdf" else "akis_json_excel"
        varsayilan = self.cfg.get(cfg_key, str(APP_DIR/f"sap_akis_{mod}.json"))
        yol,_ = QFileDialog.getSaveFileName(self,f"{etiket} Akis JSON Kaydet",varsayilan,"JSON (*.json)")
        if yol:
            try:
                Path(yol).write_text(
                    json.dumps({"sayfalar":sayfalar,"sayi":len(sayfalar)}, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )
                self.cfg[cfg_key] = yol; save_config(self.cfg)
                QMessageBox.information(self,"OK",f"{etiket} akisi kaydedildi:\n{yol}")
            except Exception as e:
                QMessageBox.critical(self, "Kayıt Hatası", f"JSON yazılamadı:\n{e}")

    def _json_yukle(self):
        mod    = "pdf" if self.akis_tab.currentIndex() == 1 else "excel"
        etiket = "PDF" if mod == "pdf" else "Excel"
        cfg_key = "akis_json_pdf" if mod == "pdf" else "akis_json_excel"
        yol,_ = QFileDialog.getOpenFileName(self,f"{etiket} Akis JSON Yukle","","JSON (*.json)")
        if not yol: return
        try:
            veri = json.loads(Path(yol).read_text(encoding="utf-8"))
            agac = self._aktif_agac()
            agac.akis_yukle(veri.get("sayfalar",[]))
            self.cfg[cfg_key] = yol; save_config(self.cfg)

            # İlk sayfayı otomatik seç ve canvas'a yükle
            self._aktif_sayfa_item = None
            self.canvas.hepsini_temizle()
            if agac.topLevelItemCount() > 0:
                ilk_sayfa = agac.topLevelItem(0)
                self._aktif_sayfa_item = ilk_sayfa
                agac.setCurrentItem(ilk_sayfa)
                sayfa_veri = ilk_sayfa.data(0, Qt.ItemDataRole.UserRole) or {}
                isim = sayfa_veri.get("isim","")
                b64  = sayfa_veri.get("goruntu_b64","")
                suffix = " 🖼" if b64 else ""
                self.aktif_sayfa_lbl.setText(f"→ {isim}{suffix}")
                if b64:
                    cv_img = agac.b64_to_cv(b64)
                    if cv_img is not None:
                        self._cv_img = cv_img
                        self.canvas.goruntu_yukle(cv_img)
                        for j in range(ilk_sayfa.childCount()):
                            adim_veri = ilk_sayfa.child(j).data(0, Qt.ItemDataRole.UserRole) or {}
                            self.canvas.alanlar.append(adim_veri)
                        self.canvas.update()

            QMessageBox.information(self,"OK",f"{etiket} akisi yuklendi.")
        except Exception as e:
            import traceback
            QMessageBox.critical(self,"Hata", f"{e}\n\n{traceback.format_exc()}")

    def _github_push(self):
        try:
            cfg = load_config()
            token = cfg.get("github_token","").strip()
            if not token:
                QMessageBox.warning(self,"GitHub","Token girilmemiş!\nAyarlar → GitHub Entegrasyonu → Token alanını doldurun.")
                return
            ok, mesaj = github_push(cfg)
            if ok:
                QMessageBox.information(self, "GitHub ✓", mesaj)
            else:
                QMessageBox.critical(self, "GitHub Hata", mesaj)
        except Exception as e:
            QMessageBox.critical(self, "GitHub Hata", f"Beklenmeyen hata:\n{e}")

# ── Gözcü (Watchdog) — Uzun çalışmada sistemi izler ───────────────────────────
class Gozcu(QThread):
    """
    Saatlerce çalışan worker'ı izleyen ajan.
    Her 30 saniyede bir kontrol eder:
      🧠 RAM: süreç RAM'i belirli bir eşiği aştı mı?
      💀 Donma: worker 5 dakikadır aktivite vermiyor mu?
      💥 Ardışık hata: üst üste 5 sipariş fail oldu mu?
      🌐 Chrome: chrome.exe process'i yaşıyor mu?

    Seviyeler:
      INFO   → log
      WARN   → log + overlay uyarısı (kırmızı)
      KRITIK → log + overlay + worker.dur()
    """
    uyari_signal = pyqtSignal(str, str)   # (seviye, mesaj)
    kritik_durdur = pyqtSignal(str)       # (sebep)

    # Eşikler (config'ten de alınabilir ama default makul)
    RAM_ESIK_MB           = 1500          # 1.5 GB → WARN
    RAM_KRITIK_MB         = 2500          # 2.5 GB → KRITIK
    DONMA_SN              = 300           # 5 dakika sessizlik → WARN
    DONMA_KRITIK_SN       = 600           # 10 dakika sessizlik → KRITIK (durdur)
    ARDISIK_HATA_ESIK     = 5             # 5 fail üst üste → KRITIK
    KONTROL_ARALIGI_SN    = 30

    def __init__(self, worker, parent=None):
        super().__init__(parent)
        self._worker = worker
        self._dur_event = threading.Event()
        self._psutil = None
        self._son_ram_uyarisi = 0          # spam önleme: aynı uyarıyı sık gönderme
        self._son_chrome_uyarisi = 0
        try:
            import psutil
            self._psutil = psutil
        except ImportError:
            pass

    def dur(self): self._dur_event.set()

    def run(self):
        # İlk kontrol için 10 sn bekle (worker ısınsın)
        if self._dur_event.wait(10): return
        while not self._dur_event.is_set():
            try:
                self._kontrol_et()
            except Exception as e:
                # Gözcü kendisi crash olmasın, sadece log
                self.uyari_signal.emit("INFO", f"Gözcü hatası (önemsiz): {e}")
            # Bir sonraki kontrole kadar bekle — dur() çağırılırsa hemen çık
            if self._dur_event.wait(self.KONTROL_ARALIGI_SN): break

    def _kontrol_et(self):
        if self._worker is None or not self._worker.isRunning():
            return

        simdi = time.time()

        # ── 1) Donma tespiti ─────────────────────────────────────────────────
        son_aktivite = getattr(self._worker, "_son_aktivite", simdi)
        sessizlik = simdi - son_aktivite
        if sessizlik > self.DONMA_KRITIK_SN:
            self.kritik_durdur.emit(
                f"Worker {int(sessizlik/60)} dakikadır sessiz — donmuş olabilir, durduruluyor"
            )
            return
        elif sessizlik > self.DONMA_SN:
            self.uyari_signal.emit("WARN",
                f"💀  Worker {int(sessizlik)}sn sessiz — Chrome donmuş olabilir")

        # ── 2) Ardışık hata ──────────────────────────────────────────────────
        ardisik = getattr(self._worker, "_ardisik_hata", 0)
        if ardisik >= self.ARDISIK_HATA_ESIK:
            self.kritik_durdur.emit(
                f"{ardisik} sipariş üst üste başarısız — SAP veya Chrome sorunlu, durduruluyor"
            )
            return

        # ── 3) RAM kontrolü (psutil varsa) ───────────────────────────────────
        if self._psutil:
            try:
                proc = self._psutil.Process(os.getpid())
                ram_mb = proc.memory_info().rss / (1024 * 1024)
                if ram_mb > self.RAM_KRITIK_MB:
                    self.kritik_durdur.emit(
                        f"RAM {ram_mb:.0f} MB'a ulaştı ({self.RAM_KRITIK_MB}+) — bellek sızıntısı, durduruluyor"
                    )
                    return
                elif ram_mb > self.RAM_ESIK_MB:
                    # 5 dakikada bir uyarı (spam önleme)
                    if simdi - self._son_ram_uyarisi > 300:
                        self.uyari_signal.emit("WARN",
                            f"🧠  RAM: {ram_mb:.0f} MB (eşik {self.RAM_ESIK_MB})")
                        self._son_ram_uyarisi = simdi
            except Exception:
                pass

            # ── 4) Chrome process kontrolü ───────────────────────────────────
            chrome_var = False
            try:
                for p in self._psutil.process_iter(["name"]):
                    try:
                        ad = (p.info.get("name") or "").lower()
                        if "chrome" in ad:
                            chrome_var = True; break
                    except (self._psutil.NoSuchProcess, self._psutil.AccessDenied):
                        continue
            except Exception:
                chrome_var = True  # kontrol başarısızsa varsayım: yaşıyor
            if not chrome_var:
                # 2 dakikada bir uyar (Chrome her an açılabilir)
                if simdi - self._son_chrome_uyarisi > 120:
                    self.uyari_signal.emit("WARN",
                        "🌐  Chrome process bulunamadı — tarayıcı kapanmış olabilir")
                    self._son_chrome_uyarisi = simdi



# ── Floating Overlay (Chrome önde iken üstte açılan bilgi çubuğu) ─────────────
class FloatingOverlay(QWidget):
    """
    Program arka planda çalışırken (Chrome önde), ekranın üst-ortasında
    her zaman görünür olan, odak çalmayan şeffaf bilgi çubuğu.
    Sürüklenebilir, × ile gizlenebilir.
    """

    def __init__(self, toplam: int, parent=None):
        flags = (
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool |
            Qt.WindowType.FramelessWindowHint
        )
        super().__init__(parent, flags)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setWindowFlag(Qt.WindowType.WindowDoesNotAcceptFocus, True)

        self._toplam     = toplam
        self._tamamlandi = 0
        self._baslama    = time.time()
        self._surukle    = None          # sürükleme başlangıç noktası

        self._kur()
        self._konumlandir()

    # ── UI kurulumu ──────────────────────────────────────────────────────────
    def _kur(self):
        self.setFixedSize(400, 96)
        ana = QVBoxLayout(self)
        ana.setContentsMargins(14, 8, 14, 8)
        ana.setSpacing(5)

        # ── Başlık satırı ──────────────────────────────────────────────────
        ust = QHBoxLayout()
        ic  = QLabel("⚙  SAP İndirici — Çalışıyor")
        ic.setStyleSheet(
            f"color:{C['accent']};font-weight:bold;font-size:11px;background:transparent;"
        )
        ust.addWidget(ic)
        ust.addStretch()
        btn_x = QPushButton("×")
        btn_x.setFixedSize(20, 20)
        btn_x.setStyleSheet(
            f"QPushButton{{background:transparent;color:{C['dim']};border:none;"
            f"font-size:16px;font-weight:bold;padding:0;}}"
            f"QPushButton:hover{{color:{C['danger']};}}"
        )
        btn_x.clicked.connect(self.hide)
        ust.addWidget(btn_x)
        ana.addLayout(ust)

        # ── İlerleme satırı ────────────────────────────────────────────────
        self.ilerleme_lbl = QLabel("⏳  Başlatılıyor…")
        self.ilerleme_lbl.setStyleSheet(
            f"color:{C['text']};font-size:13px;font-weight:bold;background:transparent;"
        )
        ana.addWidget(self.ilerleme_lbl)

        # ── ETA satırı ─────────────────────────────────────────────────────
        self.eta_lbl = QLabel("⏱  Süre hesaplanıyor…")
        self.eta_lbl.setStyleSheet(
            f"color:{C['dim']};font-size:11px;background:transparent;"
        )
        ana.addWidget(self.eta_lbl)

    # ── Ekranın üst-ortasına yerleştir ──────────────────────────────────────
    def _konumlandir(self):
        try:
            ekran = QApplication.primaryScreen().availableGeometry()
        except Exception:
            ekran = QApplication.primaryScreen().geometry()
        self.move((ekran.width() - self.width()) // 2, 16)

    # ── Arka planı elle çiz (yarı-saydam yuvarlak kutu) ──────────────────────
    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setBrush(QColor(18, 22, 34, 224))           # koyu, hafif şeffaf
        p.setPen(QPen(QColor(C["accent"]), 1.5))
        p.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 10, 10)
        p.end()

    # ── Sürükleme ─────────────────────────────────────────────────────────────
    def mousePressEvent(self, ev):
        if ev.button() == Qt.MouseButton.LeftButton:
            self._surukle = ev.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, ev):
        if self._surukle and ev.buttons() == Qt.MouseButton.LeftButton:
            self.move(ev.globalPosition().toPoint() - self._surukle)

    def mouseReleaseEvent(self, ev):
        self._surukle = None

    # ── Sinyal bağlantısı: ilerleme güncellemesi ─────────────────────────────
    def guncelle(self, tamamlandi: int, toplam: int):
        self._tamamlandi = tamamlandi
        self._toplam     = toplam
        kalan = max(0, toplam - tamamlandi)    # negatif olmasını engelle

        self.ilerleme_lbl.setText(
            f"📥  {tamamlandi} / {toplam} tamamlandı  —  {kalan} dosya kaldı"
        )

        gecen = time.time() - self._baslama
        if tamamlandi > 0 and kalan > 0 and gecen > 0:
            ort_sure      = gecen / tamamlandi          # sipariş başına ortalama süre
            kalan_sure    = kalan * ort_sure
            bitis_ts      = time.time() + kalan_sure
            try:
                bitis_str = datetime.fromtimestamp(bitis_ts).strftime("%H:%M")
            except (ValueError, OSError, OverflowError):
                bitis_str = "—"
            kalan_dk      = int(kalan_sure // 60)
            kalan_sn      = int(kalan_sure % 60)
            self.eta_lbl.setText(
                f"⏱  Tahmini bitiş: {bitis_str}  "
                f"(yaklaşık {kalan_dk} dk {kalan_sn} sn)"
            )
        elif kalan == 0:
            self.eta_lbl.setText("✅  Son sipariş işleniyor…")
        # tamamlandi==0 durumunda "Başlatılıyor…" mesajı kalır (ilk açılış)

        if not self.isVisible():
            self.show()

    # ── Gözcü uyarısı göster (4sn sonra kendini sıfırlar) ────────────────────
    def uyari_goster(self, mesaj: str, seviye: str = "WARN"):
        """Gözcü'den gelen uyarıları overlay'de kırmızı olarak göster."""
        renk = C['danger'] if seviye in ("WARN", "KRITIK") else C['warn']
        self.eta_lbl.setText(f"⚠  {mesaj}")
        self.eta_lbl.setStyleSheet(
            f"color:{renk};font-size:11px;font-weight:bold;background:transparent;"
        )
        # 6 saniye sonra normal renge dön
        QTimer.singleShot(6000, self._uyari_sifirla)

    def _uyari_sifirla(self):
        try:
            self.eta_lbl.setStyleSheet(
                f"color:{C['dim']};font-size:11px;background:transparent;"
            )
        except Exception:
            pass


# ── İndirici Worker ───────────────────────────────────────────────────────────
class IndiriciWorker(QThread):
    log_signal   = pyqtSignal(str)
    bitti_signal = pyqtSignal(dict)
    ilerleme     = pyqtSignal(int, int)

    def __init__(self, cfg, sayfalar, numbers):
        super().__init__()
        self.cfg = cfg
        self.sayfalar = sayfalar  # akış JSON'dan gelen sayfa listesi
        self.numbers = numbers
        self._dur_event = threading.Event()   # thread-safe dur flag
        self._aktif_numara = "dosya"
        self._son_dosya = None
        # ── Gözcü takip alanları ─────────────────────────────────────────────
        self._son_aktivite = time.time()       # donma tespiti için
        self._ardisik_hata = 0                  # ardışık fail sayacı

    def dur(self): self._dur_event.set()

    @property
    def _dur(self): return self._dur_event.is_set()

    def _aktivite_bildir(self):
        """Gözcü'ye 'yaşıyorum' sinyali."""
        self._son_aktivite = time.time()

    def log(self, msg, level="INFO"):
        icons = {"INFO":"·","OK":"✓","WARN":"⚠","ERROR":"✗"}
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {icons.get(level,'·')} {msg}"
        self.log_signal.emit(line)
        self._son_aktivite = time.time()   # Gözcü için heartbeat
        try:
            with open(LOG_FILE,"a",encoding="utf-8") as f: f.write(line+"\n")
        except: pass

    def _adim_calistir(self, adim, excel_numara=None):
        """Bir adımın eylem zincirini baştan sona çalıştır."""
        zincir = adim.get("eylem_zinciri", None)
        if zincir is None:
            # Geriye dönük uyumluluk
            zincir = [{"eylem": adim.get("eylem","sol_tikla"),
                       "params": adim.get("params",{})}]
        merkez = adim.get("merkez",[0,0])
        isim   = adim.get("isim","?")
        basarili = True
        for e in zincir:
            if not self._tek_eylem_calistir(e["eylem"], e.get("params",{}),
                                             merkez, isim, excel_numara):
                basarili = False
        return basarili

    def _tek_eylem_calistir(self, eylem, params, merkez, isim, excel_numara=None):
        """Tek bir eylemi çalıştır."""
        try:
            import pyautogui, pyperclip
            x, y = merkez[0], merkez[1]
            cfg  = self.cfg

            if eylem == "sol_tikla":
                self.log(f"  🖱  Tıkla → {isim} ({x},{y})")
                insan_gibi_tikla(x, y, cfg)

            elif eylem == "metin_yaz":
                kaynak = params.get("kaynak","sabit")
                if kaynak == "sabit":           metin = params.get("sabit_metin","")
                elif kaynak == "kullanici_adi": metin = cfg.get("username","")
                elif kaynak == "sifre":         metin = cfg.get("password","")
                elif kaynak == "excel_numara":  metin = str(excel_numara) if excel_numara is not None else ""
                else: metin = ""
                goster = "*" * len(metin) if kaynak == "sifre" else metin
                self.log(f"  ⌨  Yaz → {isim}: {goster}")
                import pyperclip
                pyautogui.click(x, y);         time.sleep(0.2)
                pyautogui.hotkey("ctrl", "a"); time.sleep(0.1)
                pyperclip.copy(str(metin))
                pyautogui.hotkey("ctrl", "v"); time.sleep(0.2)

            elif eylem == "enter":
                self.log(f"  ↵  Enter → {isim}")
                time.sleep(random.uniform(0.1, 0.3))
                pyautogui.press("enter")

            elif eylem == "tab":
                self.log(f"  ⇥  Tab → {isim}")
                time.sleep(random.uniform(0.05, 0.2))
                pyautogui.press("tab")

            elif eylem == "bekle":
                sure = float(params.get("sure",1.0))
                sure += random.uniform(-sure*0.1, sure*0.1)
                self.log(f"  ⏳  Bekle {sure:.1f}sn")
                time.sleep(sure)

            elif eylem == "mouse_hareket":
                hx = int(params.get("hedef_x",0)); hy = int(params.get("hedef_y",0))
                hiz = float(params.get("hiz",0.5))
                self.log(f"  ➡  Mouse → ({hx},{hy})")
                cx, cy = pyautogui.position()
                insan_gibi_hareket(cx, cy, hx, hy, hiz)

            elif eylem == "excel_numara":
                metin = str(excel_numara) if excel_numara is not None else ""
                self._aktif_numara = metin   # dosya_bekle icin sakla
                self.log(f"  📊  Excel numara → {isim}: {metin}")
                pyautogui.click(x, y);          time.sleep(0.3)
                pyautogui.hotkey("ctrl", "a");  time.sleep(0.1)
                pyperclip.copy(metin)
                pyautogui.hotkey("ctrl", "v");  time.sleep(0.3)
                pyautogui.press("enter");       time.sleep(0.3)

            elif eylem == "klavye_kisayol":
                tuslar = params.get("tuslar","")
                if tuslar:
                    self.log(f"  ⌘  Kısayol → {tuslar}")
                    time.sleep(random.uniform(0.1, 0.3))
                    pyautogui.hotkey(*[t.strip() for t in tuslar.split("+")])

            elif eylem == "chrome_ac":
                url = params.get("url","").strip() or cfg.get("sap_url","")
                self.log(f"  🌐  Chrome açılıyor → {url[:60]}")
                yollar = [
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                ]
                for yol in yollar:
                    if os.path.exists(yol):
                        subprocess.Popen([yol, "--start-maximized", url]); break
                else:
                    os.startfile(url)

            elif eylem == "url_git":
                url = params.get("url","").strip() or cfg.get("sap_url","")
                self.log(f"  🔗  URL → {url[:60]}")
                import pyperclip
                pyautogui.hotkey("ctrl","l"); time.sleep(0.4)
                pyperclip.copy(url); pyautogui.hotkey("ctrl","v")
                time.sleep(0.3); pyautogui.press("enter")

            elif eylem == "escape":
                self.log(f"  ⎋  Escape → {isim}")
                pyautogui.press("escape")

            elif eylem == "f5":
                self.log(f"  F5  F5 → {isim}")
                time.sleep(random.uniform(0.1, 0.3))
                pyautogui.press("f5")

            elif eylem == "fonksiyon_tusu":
                tus = params.get("tus", "f5").lower()
                self.log(f"  Fn  {tus.upper()} → {isim}")
                time.sleep(random.uniform(0.1, 0.3))
                pyautogui.press(tus)

            elif eylem == "mail_kod_gir":
                self.log("  📧  Mail'den kod okunuyor...")
                kod = mail_den_kod_oku(cfg, timeout=60)
                if kod:
                    self.log(f"  📧  Kod alındı: {kod}","OK")
                    pyautogui.click(x, y); time.sleep(0.3)
                    pyautogui.hotkey("ctrl","a"); time.sleep(0.1)
                    pyperclip.copy(str(kod))             # typewrite() yerine — ASCII sorunu yok
                    pyautogui.hotkey("ctrl","v"); time.sleep(0.2)
                else:
                    self.log("  📧  Mail kodu 60sn'de gelmedi!","WARN")

            elif eylem == "dosya_bekle":
                timeout=int(params.get("timeout",60))
                dl_dir_local=Path(cfg["download_folder"]); dl_dir_local.mkdir(parents=True,exist_ok=True)
                numara=getattr(self,"_aktif_numara","dosya")
                indir_klas=_windows_downloads()
                self.log(f"  ⬇  [{numara}] İzleniyor: {indir_klas}")
                try:
                    onceki={f.name:f.stat().st_size for f in indir_klas.iterdir() if f.is_file()}
                except Exception:
                    onceki={}
                bulundu=False
                for _ in range(timeout):
                    if self._dur: break
                    time.sleep(1)
                    try:
                        simdi={f.name:f.stat().st_size for f in indir_klas.iterdir() if f.is_file()}
                    except Exception:
                        continue
                    yeniler=[indir_klas/ad for ad,boyut in simdi.items()
                             if ad not in onceki and not ad.endswith((".crdownload",".tmp",".part")) and boyut>0]
                    if not yeniler: continue
                    # stat() dosya silinmişse/kilitliyse OSError fırlatabilir → filtrele
                    yeniler = [p for p in yeniler if p.exists()]
                    if not yeniler: continue
                    try:
                        kaynak=max(yeniler,key=lambda p:p.stat().st_mtime)
                    except OSError:
                        continue
                    uzanti=kaynak.suffix.lower()
                    self.log(f"  ⬇  Bulunan: {kaynak.name}")
                    hedef=dl_dir_local/f"{numara}{uzanti}"; n=2
                    while hedef.exists(): hedef=dl_dir_local/f"{numara}_{n}{uzanti}"; n+=1
                    try:
                        shutil.move(str(kaynak),str(hedef))
                        self.log(f"  ⬇  Tasindi → {hedef.name}","OK")
                        self._son_dosya=str(hedef); bulundu=True
                    except Exception as mv_e:
                        self.log(f"  ⬇  Tasima hatasi: {mv_e}","ERROR")
                    break
                if not bulundu and not self._dur:
                    self.log(f"  ⬇  {timeout}sn icinde dosya gelmedi!","ERROR")

            elif eylem == "tus_ve_goruntu":
                tus          = params.get("tus","f5").lower()
                tekrar_sure  = float(params.get("tekrar_sure", 3.0))
                maks_deneme  = int(params.get("maks_deneme", 10))
                goruntu_b64  = params.get("goruntu_b64","")
                bolge_b64    = params.get("bolge_b64","")
                bolge_rect   = params.get("bolge_rect", None)   # [x,y,w,h]
                esik         = float(cfg.get("ekran_bekleme_esik", 0.80))
                eslesince_tikla = bool(params.get("eslesince_tikla", False))

                # Kullanılacak şablon ve izleme modu
                _sablon_b64  = bolge_b64 if bolge_b64 else goruntu_b64
                _bolge_modu  = bool(bolge_b64 and bolge_rect)

                if not _sablon_b64:
                    self.log(f"  🔁  [{tus.upper()}] Görüntü tanımlanmamış, sadece tuşa basılıyor","WARN")
                    pyautogui.press(tus)
                else:
                    mod_str = f"🎯 bölge {bolge_rect}" if _bolge_modu else "🖼 tam ekran"
                    self.log(
                        f"  🔁  [{tus.upper()}] Başladı → maks {maks_deneme} deneme, "
                        f"her denemede {tekrar_sure}sn bekle, mod: {mod_str} (eşik {esik:.2f})"
                    )

                    def _goruntu_eslesdi():
                        """Eşleşme bulunursa (x,y) konum döndür, bulunamazsa None."""
                        if _bolge_modu:
                            return _bolge_goruntu_bekle(
                                _sablon_b64, bolge_rect,
                                timeout=tekrar_sure, esik=esik, log_cb=self.log
                            )
                        else:
                            ok = sayfa_goruntu_bekle(
                                _sablon_b64,
                                timeout=tekrar_sure, esik=esik, log_cb=self.log
                            )
                            return (0, 0) if ok else None

                    def _eslesince_isle(konum):
                        if eslesince_tikla and konum and konum != (0, 0):
                            tx, ty = konum
                            self.log(f"  🖱  Eşleşme konumuna tıklanıyor ({tx},{ty})")
                            time.sleep(random.uniform(0.1, 0.2))
                            insan_gibi_tikla(tx, ty, cfg)

                    bulundu = False
                    for deneme in range(1, maks_deneme + 1):
                        if self._dur: break
                        # ── 1. Adım: Önce bekle, görüntü geldi mi? ─────────
                        self.log(f"  🔁  [{tus.upper()}] Deneme {deneme}/{maks_deneme} — önce bekle")
                        konum = _goruntu_eslesdi()
                        if konum is not None:
                            self.log(f"  🔁  Görüntü geldi, devam ediliyor","OK")
                            bulundu = True
                            _eslesince_isle(konum)
                            break
                        # ── 2. Adım: Görüntü gelmedi → tuşa bas → tekrar bekle
                        self.log(f"  🔁  Görüntü gelmedi → {tus.upper()} basılıyor")
                        pyautogui.press(tus)
                        time.sleep(0.5)   # tuştan sonra sayfa değişimi için bekle
                        konum = _goruntu_eslesdi()
                        if konum is not None:
                            self.log(f"  🔁  {tus.upper()} sonrası görüntü geldi","OK")
                            bulundu = True
                            _eslesince_isle(konum)
                            break
                        else:
                            self.log(f"  🔁  {tus.upper()} sonrası da görüntü gelmedi, tekrar deneniyor")
                    if not bulundu and not self._dur:
                        self.log(f"  🔁  {maks_deneme} denemede görüntü gelmedi!","ERROR")

            return True
        except Exception as e:
            import traceback
            self.log(f"  Eylem hatası ({eylem}/{isim}): {e}", "ERROR")
            self.log(traceback.format_exc(), "ERROR")
            return False

    def run(self):
        res = {"total":len(self.numbers),"success":0,"failed":0,"files":[],"errors":[]}
        try:
            self._run_safe(res)
        except Exception as e:
            import traceback
            self.log(f"KRITIK HATA: {e}","ERROR")
            self.log(traceback.format_exc(),"ERROR")
        finally:
            # bitti_signal HER DURUMDA emit edilmeli — aksi halde UI butonları kilitli kalır
            self.bitti_signal.emit(res)

    def _run_safe(self, res):
        import pyautogui
        pyautogui.FAILSAFE = False
        cfg = self.cfg
        dl_dir = Path(cfg["download_folder"]); dl_dir.mkdir(parents=True, exist_ok=True)

        # Chrome aç ve URL'ye git artık akış JSON'dan yönetiliyor
        self.log("Akış başlatılıyor...")

        # ── Giriş + 2FA sayfaları (akış JSON'dan) ───────────────────────────
        # Sipariş döngüsü dışındaki sayfaları çalıştır (son sayfa hariç)
        # Sayfa isimleriyle mantık: "giriş", "2fa", "ana sayfa", "sipariş" vs.
        # Şimdilik tüm sayfaları sipariş döngüsünden önce çalıştır
        # "sipariş" kelimesi geçen sayfayı döngüde çalıştır.

        giris_sayfalar    = []
        siparis_sayfalar  = []
        dongu_sonu_sayfalar = []
        cikis_sayfalar    = []

        for sayfa in self.sayfalar:
            tip = sayfa.get("sayfa_tipi", "")
            # Tip yoksa eski isimlere bak (geriye dönük uyumluluk)
            if not tip:
                isim_lower = sayfa.get("isim","").lower()
                if any(k in isim_lower for k in ["sipari","numara","excel","indirme","download"]):
                    tip = "siparis"
                else:
                    tip = "giris"
            if tip == "siparis":        siparis_sayfalar.append(sayfa)
            elif tip == "dongu_sonu":   dongu_sonu_sayfalar.append(sayfa)
            elif tip == "cikis":        cikis_sayfalar.append(sayfa)
            else:                       giris_sayfalar.append(sayfa)

        # ── Giriş adımları (1 kez) ───────────────────────────────────────────
        for sayfa in giris_sayfalar:
            if self._dur: break
            isim_g=sayfa.get("isim",""); goruntu_g=sayfa.get("goruntu_b64","")
            if goruntu_g:
                self.log(f"── 🔐 Sayfa bekleniyor: {isim_g} ──")
                sayfa_goruntu_bekle(goruntu_g,self.cfg.get("ekran_bekleme_timeout",12),self.cfg.get("ekran_bekleme_esik",0.80))
            else:
                self.log(f"── 🔐 Giris: {isim_g} ──")
            for adim in sayfa.get("adimlar",[]):
                if self._dur: break
                self._adim_calistir(adim)
                time.sleep(random.uniform(0.1, 0.3))

        if self._dur:
            self.log("DURDURULDU","WARN"); return

        # 2FA kontrolü
        if cfg.get("2fa_gerekli", False):
            self.log("2FA — Outlook'tan kod bekleniyor...","WARN")
            kod = mail_den_kod_oku(cfg, timeout=60)
            if kod:
                self.log(f"2FA kodu alindi: {kod}","OK")
                _2fa_girildi = False
                for sayfa in giris_sayfalar:
                    if _2fa_girildi: break
                    for adim in sayfa.get("adimlar",[]):
                        if "2fa" in adim.get("isim","").lower():
                            m = adim.get("merkez",[0,0])
                            pyautogui.click(m[0], m[1]); time.sleep(0.3)
                            pyautogui.hotkey("ctrl","a"); time.sleep(0.1)
                            import pyperclip
                            pyperclip.copy(str(kod))
                            pyautogui.hotkey("ctrl","v"); time.sleep(0.2)
                            pyautogui.press("enter")
                            self.log("2FA kodu girildi","OK")
                            time.sleep(3)
                            _2fa_girildi = True
                            break
            else:
                self.log("2FA kodu 60sn icinde gelmedi — manuel girin","WARN")
                time.sleep(10)

        # ── Sipariş döngüsü ──────────────────────────────────────────────────
        for i, numara in enumerate(self.numbers, 1):
            if self._dur: self.log("DURDURULDU","WARN"); break
            self.log(f"─── [{i}/{len(self.numbers)}] Sipariş: {numara} ───")
            self._aktif_numara = str(numara)  # dosya_bekle için sipariş numarasını güncelle
            basarili = False

            for deneme in range(1, cfg.get("retry_count",3)+1):
                if self._dur: break
                if deneme > 1:
                    self.log(f"Tekrar {deneme}...","WARN")
                    time.sleep(random.uniform(4.0,6.0))
                try:
                    # Her deneme başında dosya takibini sıfırla
                    self._son_dosya = None

                    for sayfa in siparis_sayfalar:
                        if self._dur: break
                        isim_s=sayfa.get("isim",""); goruntu=sayfa.get("goruntu_b64","")
                        if goruntu:
                            self.log(f"  ── Sayfa bekleniyor: {isim_s} ──")
                            if not sayfa_goruntu_bekle(goruntu,self.cfg.get("ekran_bekleme_timeout",12),self.cfg.get("ekran_bekleme_esik",0.80)):
                                self.log(f"  ⚠ Goruntu eslesemedi: {isim_s}","WARN")
                        else:
                            self.log(f"  ── Sayfa: {isim_s} ──")
                        for adim in sayfa.get("adimlar",[]):
                            if self._dur: break
                            self._adim_calistir(adim, excel_numara=numara)
                            time.sleep(random.uniform(0.1, 0.3))

                    # Eğer sipariş sayfası tanımlı değilse eski fallback
                    if not siparis_sayfalar:
                        self.log("  (Sipariş sayfası tanımlanmamış, atlanıyor)","WARN")
                        basarili = True; break

                    # Dosya takibi: dosya_bekle eylemi akışta tanımlıysa o halleder
                    # Akışta dosya_bekle yoksa _son_dosya kontrolü yap
                    dosya = getattr(self, "_son_dosya", None)
                    self._son_dosya = None
                    if dosya:
                        res["success"]+=1; res["files"].append(dosya); basarili=True; break
                    else:
                        # Fallback: siparis sayfalarinda dosya_bekle eylemi var mi?
                        has_dosya_bekle = any(
                            e.get("eylem") == "dosya_bekle"
                            for s in siparis_sayfalar
                            for a in s.get("adimlar",[])
                            for e in a.get("eylem_zinciri",[{"eylem":a.get("eylem","")}])
                        )
                        if not has_dosya_bekle:
                            self.log("  Dosya bekleme adımı akışa eklenmedi","WARN")
                        basarili = True; break

                except Exception as e:
                    import traceback
                    self.log(f"  Hata: {e}","ERROR")
                    self.log(traceback.format_exc(),"ERROR")

            if not basarili:
                res["failed"]+=1; res["errors"].append(numara)
                self._ardisik_hata += 1              # Gözcü için sayaç
            else:
                self._ardisik_hata = 0                # başarılı → sıfırla

            # İlerleme — sipariş bittikten sonra sinyal gönder (başında değil)
            self.ilerleme.emit(i, len(self.numbers))

            # ── Döngü Sonu sayfaları (her siparişten sonra) ──────────────────
            if dongu_sonu_sayfalar and not self._dur:
                self.log(f"  🔄 Döngü sonu adımları çalışıyor...")
                for sayfa in dongu_sonu_sayfalar:
                    if self._dur: break
                    self.log(f"  ── 🔄 {sayfa.get('isim','')} ──")
                    for adim in sayfa.get("adimlar",[]):
                        if self._dur: break
                        self._adim_calistir(adim)
                        time.sleep(random.uniform(0.2, 0.6))

            # Siparişler arası insan gibi bekleme
            bekleme = cfg.get("delay_between_numbers",4) + random.uniform(-1.0, 2.0)
            time.sleep(max(1.0, bekleme))

        # ── Çıkış sayfaları (1 kez) ──────────────────────────────────────────
        if cikis_sayfalar and not self._dur:
            self.log("── 🚪 Çıkış adımları ──")
            for sayfa in cikis_sayfalar:
                for adim in sayfa.get("adimlar",[]):
                    self._adim_calistir(adim)
                    time.sleep(random.uniform(0.2, 0.6))

        self.log("─── TAMAMLANDI ───")
        self.log(f"Toplam:{res['total']}  Başarı:{res['success']}  Hata:{res['failed']}")
        # bitti_signal finally bloğunda emit ediliyor

# ── Sekme 2: Otomatik İndirici ────────────────────────────────────────────────
class IsYeriFiltreDlg(QDialog):
    def __init__(self,is_yerleri,parent=None):
        super().__init__(parent); self.setWindowTitle("Is Yeri Filtresi")
        self.setMinimumSize(380,420); self.setStyleSheet(SS)
        self._is_yerleri=sorted(is_yerleri); self._kur()
    def _kur(self):
        lay=QVBoxLayout(self); lay.setSpacing(8); lay.setContentsMargins(16,16,16,16)
        b=QLabel("Hangi is yerleri dahil edilsin?"); b.setStyleSheet(f"color:{C['accent']};font-weight:bold;font-size:13px;"); lay.addWidget(b)
        bl=QHBoxLayout()
        bh=QPushButton("Tumunu Sec"); bh.clicked.connect(self._hepsini_sec); bl.addWidget(bh)
        bk=QPushButton("Tumunu Kaldir"); bk.clicked.connect(self._hepsini_kaldir); bl.addWidget(bk)
        lay.addLayout(bl)
        self.liste=QListWidget(); self.liste.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        for iy in self._is_yerleri:
            it=QListWidgetItem(iy); it.setSelected(True); self.liste.addItem(it)
        lay.addWidget(self.liste)
        bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Uygula"); bb.button(QDialogButtonBox.StandardButton.Ok).setObjectName("accent")
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject); lay.addWidget(bb)
    def _hepsini_sec(self):
        for i in range(self.liste.count()): self.liste.item(i).setSelected(True)
    def _hepsini_kaldir(self):
        for i in range(self.liste.count()): self.liste.item(i).setSelected(False)
    def secili_is_yerleri(self):
        return [self.liste.item(i).text() for i in range(self.liste.count()) if self.liste.item(i).isSelected()]

class HaftalikOnizlemeDlg(QDialog):
    def __init__(self,benzersiz_liste,dosya_yolu,dl_dir,parent=None):
        super().__init__(parent); self.setWindowTitle("Haftalik Program - Onizleme")
        self.setMinimumSize(540,560); self.setStyleSheet(SS)
        self._liste=list(benzersiz_liste); self._dosya=dosya_yolu
        self._dl_dir=dl_dir; self._inmis=set(); self._inmisleri_gizle=True; self._kur()
    def _zaten_indi_mi(self,numara):
        if not self._dl_dir: return False
        dl=Path(self._dl_dir)
        if not dl.exists(): return False
        for f in dl.iterdir():
            if f.stem.split("_")[0]==str(numara): return True
        return False
    def _kur(self):
        lay=QVBoxLayout(self); lay.setSpacing(8); lay.setContentsMargins(16,16,16,16)
        b=QLabel(f"Dosya: {Path(self._dosya).name}"); b.setStyleSheet(f"color:{C['accent']};font-weight:bold;font-size:13px;"); b.setWordWrap(True); lay.addWidget(b)
        self.ozet_lbl=QLabel(); self.ozet_lbl.setStyleSheet(f"color:{C['text']};font-size:12px;"); lay.addWidget(self.ozet_lbl)
        self.liste_widget=QListWidget(); self.liste_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        lay.addWidget(self.liste_widget); self._listeyi_doldur()
        bl=QHBoxLayout()
        bs=QPushButton("Secilileri Cikar"); bs.setObjectName("danger"); bs.clicked.connect(self._secilileri_sil); bl.addWidget(bs)
        bt=QPushButton("Inmisleri Goster/Gizle"); bt.clicked.connect(self._inmisleri_toggle); bl.addWidget(bt)
        bl.addStretch(); lay.addLayout(bl)
        bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        bb.button(QDialogButtonBox.StandardButton.Ok).setText("Indirmeye Basla"); bb.button(QDialogButtonBox.StandardButton.Ok).setObjectName("accent")
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject); lay.addWidget(bb); self._ozet_guncelle()
    def _listeyi_doldur(self):
        self.liste_widget.clear(); self._inmis.clear()
        for numara in self._liste:
            indi=self._zaten_indi_mi(numara)
            if indi: self._inmis.add(numara)
            if indi and self._inmisleri_gizle: continue
            item=QListWidgetItem()
            if indi:
                item.setText(f"  {numara}"); item.setForeground(QColor(C["dim"]))
                f=item.font(); f.setStrikeOut(True); item.setFont(f)
            else:
                item.setText(f"  {numara}"); item.setForeground(QColor(C["text"]))
            self.liste_widget.addItem(item)
    def _ozet_guncelle(self):
        toplam=len(self._liste); inmis=len(self._inmis); bekleyen=toplam-inmis
        ra=C["accent"]; rd=C["dim"]
        self.ozet_lbl.setText(f"Toplam: <b>{toplam}</b>  |  <span style='color:{ra}'>Indirilecek: <b>{bekleyen}</b></span>  |  <span style='color:{rd}'>Zaten inmis: <b>{inmis}</b></span>")
        self.ozet_lbl.setTextFormat(Qt.TextFormat.RichText)
    def _secilileri_sil(self):
        for item in self.liste_widget.selectedItems():
            numara=item.text().strip()
            if numara in self._liste: self._liste.remove(numara)
            self._inmis.discard(numara); self.liste_widget.takeItem(self.liste_widget.row(item))
        self._ozet_guncelle()
    def _inmisleri_toggle(self):
        self._inmisleri_gizle=not self._inmisleri_gizle; self._listeyi_doldur(); self._ozet_guncelle()
    def secili_liste(self):
        return [n for n in self._liste if n not in self._inmis]

class IndiriciSekmesi(QWidget):
    def __init__(self, cfg):
        super().__init__()
        self.cfg = cfg
        self._worker = None
        self._worker_prev = None
        self._overlay = None
        self._gozcu = None
        self._kur()

    def _kur(self):
        ana = QHBoxLayout(self); ana.setContentsMargins(8,8,8,8); ana.setSpacing(8)

        # ── Sol ayarlar (ScrollArea içinde) ──────────────────────────────────
        sol_icerik = QWidget()
        sv = QVBoxLayout(sol_icerik); sv.setContentsMargins(4,4,4,4); sv.setSpacing(8)

        # SAP Giriş
        gb1 = QGroupBox("SAP Giriş")
        v1 = QVBoxLayout(gb1); v1.setSpacing(6)
        v1.addWidget(QLabel("URL:"))
        self.url_edit = QLineEdit(self.cfg.get("sap_url",""))
        self.url_edit.setPlaceholderText("https://..."); v1.addWidget(self.url_edit)
        v1.addWidget(QLabel("Kullanıcı:"))
        self.user_edit = QLineEdit(self.cfg.get("username","")); v1.addWidget(self.user_edit)
        v1.addWidget(QLabel("Şifre:"))
        self.pass_edit = QLineEdit(self.cfg.get("password",""))
        self.pass_edit.setEchoMode(QLineEdit.EchoMode.Password); v1.addWidget(self.pass_edit)
        self.fa2_chk = QCheckBox("Mail 2FA kodu gerekli")
        self.fa2_chk.setChecked(self.cfg.get("2fa_gerekli",False)); v1.addWidget(self.fa2_chk)
        sv.addWidget(gb1)

        # GitHub
        gb_gh = QGroupBox("GitHub Entegrasyonu")
        vg = QVBoxLayout(gb_gh); vg.setSpacing(6)
        vg.addWidget(QLabel("Personal Access Token:"))
        self.github_token_edit = QLineEdit(self.cfg.get("github_token",""))
        self.github_token_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.github_token_edit.setPlaceholderText("github_pat_...")
        vg.addWidget(self.github_token_edit)
        gh_repo_lay = QHBoxLayout()
        gh_repo_lay.addWidget(QLabel("Repo:"))
        self.github_repo_edit = QLineEdit(self.cfg.get("github_repo","Elcihad/Cloude"))
        gh_repo_lay.addWidget(self.github_repo_edit)
        vg.addLayout(gh_repo_lay)
        sv.addWidget(gb_gh)

        # Mail 2FA — Outlook
        gb_mail = QGroupBox("Mail 2FA — Outlook")
        vm = QVBoxLayout(gb_mail); vm.setSpacing(6)
        lbl_info = QLabel("Kurulu Outlook üzerinden okur.\npip install pywin32")
        lbl_info.setStyleSheet(f"color:{C['dim']};font-size:10px;")
        vm.addWidget(lbl_info)
        vm.addWidget(QLabel("Konu Filtresi (mail başlığında arar):"))
        self.mail_filtre_edit = QLineEdit(self.cfg.get("mail_konu_filtre",""))
        self.mail_filtre_edit.setPlaceholderText("Örn: Doğrulama Kodu (boş=son mail)")
        vm.addWidget(self.mail_filtre_edit)
        vm.addWidget(QLabel("Gönderen Filtresi (boş=herkesten):"))
        self.mail_gonderen_edit = QLineEdit(self.cfg.get("mail_gonderen_filtre",""))
        self.mail_gonderen_edit.setPlaceholderText("Örn: burak.cosgun@tupras.com.tr")
        vm.addWidget(self.mail_gonderen_edit)
        vm.addWidget(QLabel("Klasör Yolu (boş = Inbox):"))
        self.mail_klasor_edit = QLineEdit(self.cfg.get("mail_klasor_yolu",""))
        self.mail_klasor_edit.setPlaceholderText("Örn: Tüpraş/Doğrulama")
        vm.addWidget(self.mail_klasor_edit)
        # Test butonu
        btn_mail_test = QPushButton("📧  Bağlantıyı Test Et")
        btn_mail_test.clicked.connect(self._mail_test); vm.addWidget(btn_mail_test)
        sv.addWidget(gb_mail)

        # Excel Ayarları
        gb2 = QGroupBox("Excel Ayarları")
        v2 = QVBoxLayout(gb2); v2.setSpacing(6)
        v2.addWidget(QLabel("Haftalik Program:"))
        ef=QHBoxLayout(); self.excel_edit=QLineEdit(self.cfg.get("excel_file",""))
        ef.addWidget(self.excel_edit)
        bxs=QPushButton("..."); bxs.setFixedWidth(36); bxs.clicked.connect(self._sec_excel); ef.addWidget(bxs)
        bxr=QPushButton("x"); bxr.setFixedWidth(28); bxr.setObjectName("danger"); bxr.clicked.connect(self._excel_sifirla); ef.addWidget(bxr)
        v2.addLayout(ef)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Sayfa:"))
        self.sheet_spin = QSpinBox(); self.sheet_spin.setFixedWidth(55)
        self.sheet_spin.setValue(self.cfg.get("excel_sheet",0)); row1.addWidget(self.sheet_spin)
        row1.addSpacing(10)
        row1.addWidget(QLabel("Sütun:"))
        self.sutun_edit = QLineEdit(self.cfg.get("numara_sutun","C")); self.sutun_edit.setFixedWidth(45)
        row1.addWidget(self.sutun_edit); row1.addStretch(); v2.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Başlangıç satırı:"))
        self.satir_spin = QSpinBox(); self.satir_spin.setMinimum(1); self.satir_spin.setFixedWidth(55)
        self.satir_spin.setValue(self.cfg.get("numara_baslangic_satir",4))
        row2.addWidget(self.satir_spin); row2.addStretch(); v2.addLayout(row2)
        sv.addWidget(gb2)

        # İndirme Klasörü
        gb3 = QGroupBox("İndirme Klasörü")
        v3 = QVBoxLayout(gb3); v3.setSpacing(6)
        df = QHBoxLayout(); self.dl_edit = QLineEdit(self.cfg.get("download_folder",""))
        df.addWidget(self.dl_edit)
        btn_dl = QPushButton("..."); btn_dl.setFixedWidth(36)
        btn_dl.clicked.connect(self._sec_dl); df.addWidget(btn_dl); v3.addLayout(df)
        sv.addWidget(gb3)

        # Akış JSON — Excel ve PDF ayrı
        gb4 = QGroupBox("Akış JSON")
        v4 = QVBoxLayout(gb4); v4.setSpacing(6)

        v4.addWidget(QLabel("📊  Excel Akışı:"))
        jf1 = QHBoxLayout()
        self.json_excel_edit = QLineEdit(self.cfg.get("akis_json_excel",""))
        self.json_excel_edit.setPlaceholderText("sap_akis_excel.json")
        jf1.addWidget(self.json_excel_edit)
        btn_json1 = QPushButton("..."); btn_json1.setFixedWidth(36)
        btn_json1.clicked.connect(lambda: self._sec_json_alan(self.json_excel_edit))
        jf1.addWidget(btn_json1); v4.addLayout(jf1)

        v4.addWidget(QLabel("📄  PDF Akışı:"))
        jf2 = QHBoxLayout()
        self.json_pdf_edit = QLineEdit(self.cfg.get("akis_json_pdf",""))
        self.json_pdf_edit.setPlaceholderText("sap_akis_pdf.json")
        jf2.addWidget(self.json_pdf_edit)
        btn_json2 = QPushButton("..."); btn_json2.setFixedWidth(36)
        btn_json2.clicked.connect(lambda: self._sec_json_alan(self.json_pdf_edit))
        jf2.addWidget(btn_json2); v4.addLayout(jf2)

        sv.addWidget(gb4)

        # İnsan Gibi Davranma (daraltılmış)
        gb_insan = QGroupBox("İnsan Gibi Davranma")
        vi = QVBoxLayout(gb_insan); vi.setSpacing(4)
        for lbl_txt, attr, mn, mx, step, default_key in [
            ("Mouse hız min:", "mouse_min", 0.1, 3.0, 0.1, "mouse_hiz_min"),
            ("Mouse hız max:", "mouse_max", 0.1, 5.0, 0.1, "mouse_hiz_max"),
            ("Yazı hız min:",  "yazi_min",  0.01,1.0, 0.01,"yazi_hiz_min"),
            ("Yazı hız max:",  "yazi_max",  0.01,2.0, 0.01,"yazi_hiz_max"),
        ]:
            row = QHBoxLayout()
            row.addWidget(QLabel(lbl_txt))
            spin = QDoubleSpinBox(); spin.setRange(mn, mx); spin.setSingleStep(step)
            spin.setValue(self.cfg.get(default_key, 0.3)); spin.setFixedWidth(70)
            setattr(self, attr, spin); row.addWidget(spin); row.addStretch()
            vi.addLayout(row)
        sv.addWidget(gb_insan)

        btn_kaydet = QPushButton("💾  Ayarları Kaydet")
        btn_kaydet.setObjectName("accent")
        btn_kaydet.clicked.connect(self._ayar_kaydet)
        sv.addWidget(btn_kaydet); sv.addStretch()

        # ScrollArea — sol panel kaydırılabilir
        sol_scroll = QScrollArea()
        sol_scroll.setWidget(sol_icerik)
        sol_scroll.setWidgetResizable(True)
        sol_scroll.setFixedWidth(310)
        sol_scroll.setFrameShape(QFrame.Shape.NoFrame)
        sol_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        ana.addWidget(sol_scroll)

        # ── Sağ: Log + Akış Önizleme + Kontrol ──────────────────────────────
        sag = QWidget(); sagv = QVBoxLayout(sag); sagv.setContentsMargins(0,0,0,0); sagv.setSpacing(6)

        # İlerleme
        ilerleme_lay = QHBoxLayout()
        self.progress = QProgressBar(); self.progress.setValue(0)
        ilerleme_lay.addWidget(self.progress)
        self.ilerleme_lbl = QLabel("0/0"); ilerleme_lay.addWidget(self.ilerleme_lbl)
        sagv.addLayout(ilerleme_lay)

        # Akış önizleme
        gb_onizleme = QGroupBox("Akış Önizleme")
        vo = QVBoxLayout(gb_onizleme)
        self.akis_onizleme = QTextEdit(); self.akis_onizleme.setReadOnly(True)
        self.akis_onizleme.setFont(QFont("Consolas",9)); self.akis_onizleme.setMaximumHeight(140)
        vo.addWidget(self.akis_onizleme)
        btn_akis_yenile = QPushButton("🔄  JSON'u Yenile ve Önizle"); btn_akis_yenile.setFixedHeight(26)
        btn_akis_yenile.clicked.connect(self._akis_onizle); vo.addWidget(btn_akis_yenile)
        sagv.addWidget(gb_onizleme)

        # Log
        gb_log = QGroupBox("Log")
        vlog = QVBoxLayout(gb_log)
        self.log_edit = QTextEdit(); self.log_edit.setReadOnly(True)
        self.log_edit.setFont(QFont("Consolas",10)); vlog.addWidget(self.log_edit)
        sagv.addWidget(gb_log)

        # Butonlar
        btn_lay = QHBoxLayout(); btn_lay.setSpacing(8)
        self.btn_excel = QPushButton("📊  EXCEL İNDİR"); self.btn_excel.setObjectName("accent")
        self.btn_excel.clicked.connect(lambda: self._baslat("excel")); btn_lay.addWidget(self.btn_excel)

        self.btn_pdf = QPushButton("📄  PDF İNDİR")
        self.btn_pdf.setStyleSheet("QPushButton{background:#6c63ff;color:#fff;border:none;border-radius:6px;padding:7px 16px;font-size:12px;min-height:32px;font-weight:bold;}QPushButton:hover{background:#857dff;}")
        self.btn_pdf.clicked.connect(lambda: self._baslat("pdf")); btn_lay.addWidget(self.btn_pdf)

        self.btn_dur = QPushButton("⏹  DURDUR"); self.btn_dur.setObjectName("danger")
        self.btn_dur.clicked.connect(self._durdur); self.btn_dur.setEnabled(False)
        btn_lay.addWidget(self.btn_dur)
        btn_log_temizle = QPushButton("Log Temizle"); btn_log_temizle.clicked.connect(self.log_edit.clear)
        btn_lay.addWidget(btn_log_temizle); sagv.addLayout(btn_lay); ana.addWidget(sag)

    def _sec_excel(self):
        x,_=QFileDialog.getOpenFileName(self,"Haftalik Program Sec",self.excel_edit.text(),"Excel (*.xlsx *.xls)")
        if not x: return
        self.excel_edit.setText(x); self._haftalik_onizle(x)
    def _excel_sifirla(self):
        if QMessageBox.question(self,"Sifirla","Excel listesi temizlensin mi?",
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            self.excel_edit.setText(""); self.cfg["excel_file"]=""; save_config(self.cfg)
    @staticmethod
    def _col_idx(h):
        i=0
        for c in h.upper(): i=i*26+(ord(c)-ord("A")+1)
        return i-1
    def _excel_oku_filtreli(self,dosya_yolu):
        import pandas as pd
        # NOT: numara_baslangic_satir config'i var ama şu anda kullanılmıyor.
        # Kullanıcının mevcut Excel kurulumunu bozmamak için header=0 sabit bırakıldı.
        # Gelecekte UI'da açıkça "Başlık satırı" seçimi gerekirse şuraya bağlanabilir.
        df=pd.read_excel(dosya_yolu,sheet_name=self.sheet_spin.value(),header=0,dtype=str)
        df.dropna(how="all",inplace=True)
        siparis_col=df.iloc[:,self._col_idx(self.sutun_edit.text().strip().upper() or "C")]
        try:
            is_yeri_col=df.iloc[:,11]
            is_yerleri=sorted([x for x in is_yeri_col.dropna().unique() if str(x).strip()])
        except Exception:
            is_yerleri=[]
        if is_yerleri:
            fdlg=IsYeriFiltreDlg(is_yerleri,self)
            if not fdlg.exec(): return None,is_yerleri
            secili=fdlg.secili_is_yerleri()
            if not secili:
                QMessageBox.warning(self,"Filtre","En az bir is yeri secin!"); return None,is_yerleri
            siparis_listesi=siparis_col[is_yeri_col.isin(secili)]
        else:
            siparis_listesi=siparis_col
        ham=[]
        for raw in siparis_listesi:
            val=str(raw).strip().split(".")[0]
            if val not in ("","nan","NaN","None"): ham.append(val)
        return list(dict.fromkeys(ham)),is_yerleri
    def _haftalik_onizle(self,dosya_yolu,mod="excel"):
        try:
            benzersiz,_=self._excel_oku_filtreli(dosya_yolu)
            if benzersiz is None: return
            if not benzersiz:
                QMessageBox.warning(self,"Bos","Secili is yerlerinde siparis bulunamadi!"); return
            # pdf_dl_edit widget'ı oluşturulmadı — config'ten oku
            if mod == "pdf":
                dl_dir = self.cfg.get("pdf_download_folder","").strip()
            else:
                dl_dir = self.dl_edit.text().strip()
            HaftalikOnizlemeDlg(benzersiz,dosya_yolu,dl_dir,self).exec()
        except Exception as e:
            QMessageBox.critical(self,"Hata",f"Excel okunamadi:\n{e}")
    def _sec_dl(self):
        x=QFileDialog.getExistingDirectory(self,"Excel Indirme",self.dl_edit.text())
        if x: self.dl_edit.setText(x)

    def _mail_test(self):
        """Outlook bağlantısını test et — alt klasörü de kontrol eder."""
        self._ayar_kaydet(sessiz=True)
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns      = outlook.GetNamespace("MAPI")
            inbox   = ns.GetDefaultFolder(6)
            klasor_yolu = self.cfg.get("mail_klasor_yolu", "").strip()
            hedef = _outlook_klasor_bul(inbox, klasor_yolu)
            if hedef is None:
                QMessageBox.warning(self, "Klasör Bulunamadı",
                    f"Outlook bağlantısı başarılı AMA\n"
                    f"'{klasor_yolu}' klasörü bulunamadı!\n\n"
                    f"Klasör adlarını kontrol edin.")
            else:
                sayi = hedef.Items.Count
                klasor_adi = hedef.Name if klasor_yolu else "Inbox"
                QMessageBox.information(self, "✓ Outlook Bağlandı",
                    f"Outlook bağlantısı başarılı!\n"
                    f"Klasör: {klasor_adi}\n"
                    f"Mail sayısı: {sayi}")
        except ImportError:
            QMessageBox.warning(self,"pywin32 Eksik",
                "Önce şunu çalıştırın:\npip install pywin32")
        except Exception as e:
            QMessageBox.critical(self,"Hata",
                f"Outlook bağlantı hatası:\n{e}\n\nOutlook açık mı?")

    def _sec_json_alan(self, edit_widget):
        x,_ = QFileDialog.getOpenFileName(self,"JSON","","JSON (*.json)")
        if x:
            edit_widget.setText(x)
            self._akis_onizle()

    def _sec_json(self):
        self._sec_json_alan(self.json_excel_edit)

    def _akis_onizle(self):
        # Önce Excel akışını, yoksa PDF akışını önizle
        yol = self.json_excel_edit.text().strip()
        if not yol:
            yol = self.json_pdf_edit.text().strip()
        if not yol or not os.path.exists(yol):
            self.akis_onizleme.setText("JSON bulunamadı."); return
        try:
            veri = json.loads(Path(yol).read_text(encoding="utf-8"))
            sayfalar = veri.get("sayfalar",[])
            satirlar = []
            for sayfa in sayfalar:
                satirlar.append(f"📄 {sayfa.get('isim','')}")
                for adim in sayfa.get("adimlar",[]):
                    eylem_meta = EYLEMLER.get(adim.get("eylem",""),{})
                    satirlar.append(f"   {eylem_meta.get('icon','?')} {adim.get('isim','')}  [{eylem_meta.get('etiket','')}]")
            self.akis_onizleme.setText("\n".join(satirlar))
        except Exception as e:
            self.akis_onizleme.setText(f"Hata: {e}")

    def _ayar_kaydet(self, sessiz=False):
        self.cfg.update({
            "sap_url":          self.url_edit.text().strip(),
            "github_token":    self.github_token_edit.text().strip() if hasattr(self,"github_token_edit") else self.cfg.get("github_token",""),
            "github_repo":     self.github_repo_edit.text().strip() if hasattr(self,"github_repo_edit") else self.cfg.get("github_repo","Elcihad/Cloude"),
            "username":         self.user_edit.text().strip(),
            "password":         self.pass_edit.text(),
            "2fa_gerekli":      self.fa2_chk.isChecked(),
            "excel_file":       self.excel_edit.text().strip(),
            "excel_sheet":      self.sheet_spin.value(),
            "numara_sutun":     self.sutun_edit.text().strip().upper(),
            "numara_baslangic_satir": self.satir_spin.value(),
            "download_folder":  self.dl_edit.text().strip(),
            "akis_json":        self.json_excel_edit.text().strip(),
            "akis_json_excel":  self.json_excel_edit.text().strip(),
            "akis_json_pdf":    self.json_pdf_edit.text().strip(),
            "mail_konu_filtre": self.mail_filtre_edit.text().strip(),
            "mail_gonderen_filtre": self.mail_gonderen_edit.text().strip(),
            "mail_klasor_yolu": self.mail_klasor_edit.text().strip(),
            "mouse_hiz_min":    self.mouse_min.value(),
            "mouse_hiz_max":    self.mouse_max.value(),
            "yazi_hiz_min":     self.yazi_min.value(),
            "yazi_hiz_max":     self.yazi_max.value(),
        })
        save_config(self.cfg)
        if not sessiz:
            # Popup yerine status bar'a yaz
            try: self.window().statusBar().showMessage("  ✓ Ayarlar kaydedildi", 3000)
            except: pass

    def _log(self, msg):
        self.log_edit.append(msg)
        self.log_edit.moveCursor(QTextCursor.MoveOperation.End)

    def _baslat(self,mod="excel"):
        self._ayar_kaydet(sessiz=True)
        anahtar="akis_json_excel" if mod=="excel" else "akis_json_pdf"
        etiket="Excel" if mod=="excel" else "PDF"
        akis_json=self.cfg.get(anahtar,"")
        if not akis_json or not os.path.exists(akis_json):
            QMessageBox.warning(self,"Eksik",f"{etiket} akis JSON secilmedi!"); return
        try:
            veri=json.loads(Path(akis_json).read_text(encoding="utf-8")); sayfalar=veri.get("sayfalar",[])
        except Exception as e:
            QMessageBox.critical(self,"Hata",f"JSON okunamadi:\n{e}"); return
        if not sayfalar:
            QMessageBox.warning(self,"Bos","Akis JSON'da sayfa bulunamadi!"); return
        excel=self.cfg.get("excel_file","")
        if not excel or not os.path.exists(excel):
            QMessageBox.warning(self,"Eksik","Excel dosyasi secilmedi!"); return
        try:
            benzersiz,_=self._excel_oku_filtreli(excel)
            if benzersiz is None: return
        except Exception as e:
            QMessageBox.critical(self,"Hata",f"Excel okunamadi:\n{e}"); return
        if not benzersiz:
            QMessageBox.warning(self,"Bos","Secili is yerlerinde siparis bulunamadi!"); return
        dl_dir=self.cfg.get("pdf_download_folder","") if mod=="pdf" else self.cfg.get("download_folder","")
        onizleme=HaftalikOnizlemeDlg(benzersiz,excel,dl_dir,self)
        if not onizleme.exec(): return
        numbers=onizleme.secili_liste()
        if not numbers:
            QMessageBox.information(self,"Bilgi","Tum siparisler zaten inmis."); return
        self._log(f"── {'📊' if mod=='excel' else '📄'} {etiket} — {len(numbers)} siparis ──")
        self.progress.setMaximum(len(numbers)); self.progress.setValue(0); self.ilerleme_lbl.setText(f"0/{len(numbers)}")
        self.btn_excel.setEnabled(False); self.btn_pdf.setEnabled(False); self.btn_dur.setEnabled(True)
        cfg_w=dict(self.cfg)
        if mod=="pdf" and dl_dir: cfg_w["download_folder"]=dl_dir
        # Hâlâ çalışan worker varsa YENİ worker oluşturmadan ÖNCE durdur
        # (normalde buton disabled olur ama programatik çağrıya karşı koruma)
        for w_attr in ("_worker", "_worker_prev"):
            w = getattr(self, w_attr, None)
            if w and w.isRunning():
                w.dur(); w.wait(3000)
        self._worker_prev = None
        self._worker=IndiriciWorker(cfg_w,sayfalar,numbers)
        self._worker.log_signal.connect(self._log)
        self._worker.ilerleme.connect(self._ilerleme_guncelle)
        self._worker.bitti_signal.connect(self._bitti)
        # ── FloatingOverlay — Chrome önde iken üstte görünen bilgi çubuğu ──
        self._overlay = FloatingOverlay(len(numbers))
        self._overlay.show()
        # ── Gözcü — uzun çalışmada sistemi izle ──────────────────────────
        self._gozcu = Gozcu(self._worker)
        self._gozcu.uyari_signal.connect(self._gozcu_uyari)
        self._gozcu.kritik_durdur.connect(self._gozcu_kritik)
        self._gozcu.start()
        self._worker.start(); QTimer.singleShot(400,self.window().showMinimized)


    def _gozcu_uyari(self, seviye: str, mesaj: str):
        """Gözcü'den gelen normal uyarı — log + overlay kırmızı."""
        self._log(f"[GÖZCÜ/{seviye}] {mesaj}")
        if self._overlay:
            self._overlay.uyari_goster(mesaj, seviye)

    def _gozcu_kritik(self, sebep: str):
        """Gözcü'den gelen KRİTİK uyarı — worker'ı durdur."""
        self._log(f"[GÖZCÜ/KRITIK] {sebep}")
        if self._overlay:
            self._overlay.uyari_goster(sebep, "KRITIK")
        if self._worker and self._worker.isRunning():
            self._worker.dur()

    def _ilerleme_guncelle(self, tamamlandi: int, toplam: int):
        """İlerleme çubuğu + FloatingOverlay'i aynı anda günceller."""
        self.progress.setValue(tamamlandi)
        self.ilerleme_lbl.setText(f"{tamamlandi}/{toplam}")
        if hasattr(self, "_overlay") and self._overlay:
            self._overlay.guncelle(tamamlandi, toplam)

    def _durdur(self):
        if self._worker: self._worker.dur()
        self.btn_dur.setEnabled(False)
        # Overlay ve Gözcü'yü kapat
        if hasattr(self, "_overlay") and self._overlay:
            self._overlay.hide(); self._overlay = None
        if hasattr(self, "_gozcu") and self._gozcu:
            self._gozcu.dur(); self._gozcu = None

    def _bitti(self, res):
        # Overlay ve Gözcü'yü kapat
        if hasattr(self, "_overlay") and self._overlay:
            self._overlay.hide(); self._overlay = None
        if hasattr(self, "_gozcu") and self._gozcu:
            self._gozcu.dur(); self._gozcu = None
        self._worker_prev = self._worker   # GC için referansı sakla
        self._worker = None
        self.btn_excel.setEnabled(True); self.btn_pdf.setEnabled(True)
        self.btn_dur.setEnabled(False); self.progress.setValue(res["total"])
        self.window().showNormal(); self.window().raise_(); self.window().activateWindow()
        QMessageBox.information(self,"Tamamlandi",
            f"Toplam: {res['total']}\nBaşarı: {res['success']}\nHata: {res['failed']}")

# ── Ana Pencere ───────────────────────────────────────────────────────────────
class SAPSuite(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAP Suite v2 — Alan Tanıtma + Otomatik İndirici")
        self.setMinimumSize(1200, 760); self.resize(1440, 860)
        self.setStyleSheet(SS)
        self.cfg = load_config()
        self._kur()
        self.statusBar().showMessage("  Hazır — v2.0  |  F9: Anlık Ekran Görüntüsü")

    def _kur(self):
        merkez = QWidget(); self.setCentralWidget(merkez)
        lay = QVBoxLayout(merkez); lay.setContentsMargins(8,8,8,8)
        hdr = QWidget(); hdr.setStyleSheet(f"background:{C['card']};border-radius:8px;")
        hl = QHBoxLayout(hdr); hl.setContentsMargins(16,8,16,8)
        t = QLabel("SAP SUITE v2"); t.setFont(QFont("Segoe UI",16,QFont.Weight.Bold))
        t.setStyleSheet(f"color:{C['accent']};background:transparent;"); hl.addWidget(t)
        s = QLabel("Dinamik Akış Tanımlama  +  İnsan Gibi Otomasyon")
        s.setStyleSheet(f"color:{C['dim']};background:transparent;"); hl.addWidget(s)
        hl.addStretch()
        # F9 ipucu etiketi
        f9_lbl = QLabel("[ F9 → Ekran Görüntüsü ]")
        f9_lbl.setStyleSheet(f"color:{C['warn']};background:transparent;font-size:11px;font-weight:bold;")
        hl.addWidget(f9_lbl)
        lay.addWidget(hdr)
        self.tabs = QTabWidget()
        self.alan_tab = AlanTanitmaSekmesi(self.cfg)
        self.indirici_tab = IndiriciSekmesi(self.cfg)
        self.tabs.addTab(self.alan_tab, "🎯  Alan Tanıtma")
        self.tabs.addTab(self.indirici_tab, "⬇  Otomatik İndirici")
        lay.addWidget(self.tabs)
        self.setStatusBar(QStatusBar())

        # Global F9 kısayolu — Qt odakta olmasa da (tarayıcı, SAP vs.) çalışır
        self._global_hook_baslat()

    def _global_hook_baslat(self):
        """pynput ile arka planda F9'u dinle — tarayıcı/SAP odaktayken de çalışır."""
        self._f9_bekliyor = False
        try:
            from pynput import keyboard as pynput_kb

            def _on_press(key):
                try:
                    if key == pynput_kb.Key.f9 and not self._f9_bekliyor:
                        self._f9_bekliyor = True
                        QTimer.singleShot(0, self._f9_tetiklendi)
                except Exception:
                    pass

            self._pynput_listener = pynput_kb.Listener(on_press=_on_press)
            self._pynput_listener.daemon = True
            self._pynput_listener.start()
            self._keyboard_aktif = True
        except Exception:
            # pynput yoksa Qt shortcut'a düş (sadece uygulama odaktayken çalışır)
            self._keyboard_aktif = False
            self._f9_shortcut = QShortcut(QKeySequence("F9"), self)
            self._f9_shortcut.setContext(Qt.ShortcutContext.ApplicationShortcut)
            self._f9_shortcut.activated.connect(self._f9_tetiklendi)

    def _f9_tetiklendi(self):
        """F9 basıldı — pencereyi gizle, ekran görüntüsü al."""
        self._f9_bekliyor = False
        # Minimize et, tarayıcı/SAP görünsün
        self.showMinimized()
        QTimer.singleShot(800, self._f9_yakala)

    def closeEvent(self, event):
        """Pencere kapanınca keyboard hook'u, worker thread'i ve gözcü'yü temizle."""
        # 1) Worker thread çalışıyorsa durdur — aksi halde Chrome'u karıştırmaya devam eder
        try:
            worker = getattr(self.indirici_tab, "_worker", None)
            if worker and worker.isRunning():
                worker.dur()
                worker.wait(5000)   # 5sn bekle, ardından zorla bırak
        except Exception:
            pass
        # 2) Gözcü'yü durdur
        try:
            gozcu = getattr(self.indirici_tab, "_gozcu", None)
            if gozcu and gozcu.isRunning():
                gozcu.dur()
                gozcu.wait(2000)
        except Exception:
            pass
        # 3) pynput dinleyiciyi durdur
        try:
            if hasattr(self, '_pynput_listener'):
                self._pynput_listener.stop()
        except Exception:
            pass
        super().closeEvent(event)

    def _f9_yakala(self):
        alan_tab = self.alan_tab
        try:
            import mss
            with mss.mss() as sct:
                mon = sct.monitors[1]
                shot = sct.grab(mon)
                cv_img = cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)

            alan_tab._cv_img = cv_img
            alan_tab.canvas.goruntu_yukle(cv_img)

            # Aktif sayfa varsa görüntüyü ona kaydet
            if alan_tab._aktif_sayfa_item is not None:
                b64 = alan_tab.akis_agaci._cv_to_b64(cv_img)
                alan_tab.akis_agaci._sayfa_goruntu_guncelle(alan_tab._aktif_sayfa_item, b64)
                isim = alan_tab._aktif_sayfa_item.data(0, Qt.ItemDataRole.UserRole).get("isim","")
                alan_tab.aktif_sayfa_lbl.setText(f"→ {isim} 🖼")

            self.showNormal()
            self.raise_()
            self.activateWindow()
            self.statusBar().showMessage("  ✓ F9 — Ekran görüntüsü alındı", 4000)
        except Exception as e:
            self.showNormal()
            QMessageBox.warning(self, "Hata", f"Ekran görüntüsü alınamadı:\n{e}")

def main():
    try:
        app = QApplication(sys.argv); app.setStyle("Fusion")
        pencere = SAPSuite(); pencere.show()
        sys.exit(app.exec())
    except Exception as e:
        import traceback
        hata = traceback.format_exc()
        # Hata dosyasına yaz
        try:
            (APP_DIR / "crash_log.txt").write_text(hata, encoding="utf-8")
        except Exception:
            pass
        # Konsola da yaz
        print("HATA:", hata, file=sys.stderr)
        # Eğer Qt başladıysa mesaj kutusu göster
        try:
            app2 = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Başlatma Hatası",
                f"Program başlatılamadı:\n\n{e}\n\nDetay: {APP_DIR}\\crash_log.txt")
        except Exception:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()
