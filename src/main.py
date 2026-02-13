# -*- coding: utf-8 -*-
"""
VBS klikacz NTQ – v4.2.3 (STABLE)

Zmiany v4.2.3:
• Powiadomienia: konfiguracja typów, ON/OFF, głośność, wybór pliku i test dźwięku.
• Powiadomienia: zapis/wczytywanie ustawień lokalnych (settings.json).
• Audio: poprawione odtwarzanie MP3/WAV oraz wywołania dźwięku dla zdarzeń START/STOP.
• Zachowane: stabilna logika rezerwacji slotów i potwierdzeń TAK/OK.

Zasady licencji:
- Status OK -> program działa
- Status EXPIRED lub BLOCKED -> program zamyka się po naciśnięciu START lub STOP
"""

import re
import time
import threading
import hashlib
import getpass
import json
import datetime as dt
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from playwright.sync_api import sync_playwright

VERSION = "v4.2.3"
BUILD_TIME = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
SETTINGS_FILE = Path(__file__).resolve().parent / "settings.json"

CHANGELOG = [
    "FIX: pętla potwierdzeń TAK/OK po kliknięciu slotu – klika do skutku dopóki przyciski są aktywne.",
    "FIX: poprawione klikanie slotów (kafelek zawiera też liczby zajętości, np. 44/130).",
    "Info: status licencji + ważna do + data ostatniego sprawdzenia.",
    "Utrzymano szybkie odświeżanie 1 dnia: klik STANDARDOWE."
]

# --- LICENCJA (Apps Script) ---
LICENSE_URL = "https://script.google.com/macros/s/AKfycbzRSqVDxYLQSst83z2aW_S3ftMV-jfyLTdp4AUWsHRdNxJ3epkbANOK-0KwZY5d5F1K/exec"
LICENSE_HTTP_TIMEOUT_S = 8

# ------------------ REGEX ------------------

SLOT_RE = re.compile(r"(\d{2}:\d{2})-(\d{2}:\d{2})\s+(\d+)/(\d+)")


# ------------------ LICENSE / MACHINE ID ------------------

def _get_machine_guid_windows():
    """Windows MachineGuid z rejestru. Zwraca string albo None."""
    try:
        import winreg  # type: ignore
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography"
        ) as key:
            return winreg.QueryValueEx(key, "MachineGuid")[0]
    except Exception:
        return None


def generate_machine_id():
    """
    Stabilne ID komputera:
    - bazuje na MachineGuid (Windows) + user
    - hash SHA-256, skrócony do 16 znaków
    """
    user = getpass.getuser() or "unknown"
    guid = _get_machine_guid_windows() or "no-guid"
    raw = f"{guid}|{user}"
    h = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()
    return h[:16]


def fetch_license_status(hwid: str) -> dict:
    """
    Woła Google Apps Script:
    GET .../exec?hwid=XXXX

    Oczekiwane odpowiedzi:
      { hwid, status: 'OK'|'EXPIRED'|'BLOCKED', valid_to: ... }
    Możliwe też: NOT_FOUND / error itp.
    """
    qs = urlencode({"hwid": hwid})
    url = f"{LICENSE_URL}?{qs}"

    req = Request(url, headers={
        "User-Agent": f"NTQ-VBS/{VERSION} (Python)"
    })

    with urlopen(req, timeout=LICENSE_HTTP_TIMEOUT_S) as resp:
        body = resp.read().decode("utf-8", errors="replace")
        data = json.loads(body)

    status = str(data.get("status", "")).strip().upper()
    valid_to = data.get("valid_to", "")

    return {
        "hwid": str(data.get("hwid", hwid)),
        "status": status if status else "UNKNOWN",
        "valid_to": valid_to if valid_to is not None else ""
    }


def is_license_valid(status: str) -> bool:
    status = (status or "").strip().upper()
    return status == "OK"


# ------------------ PLAYWRIGHT HELPERS ------------------

def wait_for_slots_loaded(page, timeout_ms):
    """Czekaj aż zniknie 'Ładowanie slotów' (best-effort)."""
    try:
        loc = page.locator("text=Ładowanie slotów").first
        if loc.count() and loc.is_visible():
            loc.wait_for(state="hidden", timeout=int(timeout_ms))
    except Exception:
        pass


def ensure_slot_screen(page):
    """Sprawdza, czy jesteśmy na ekranie wyboru slotów."""
    try:
        has_std = page.locator("text=STANDARDOWE").count() > 0
        has_slot = page.locator(r"text=/\b\d{2}:\d{2}-\d{2}:\d{2}\b/").count() > 0
        if has_std and has_slot:
            return
    except Exception:
        pass
    raise RuntimeError(
        "Nie jestem na ekranie wyboru okienek (slotów). "
        "Otwórz awizację w widoku slotów (kalendarz + siatka slotów) i dopiero kliknij START."
    )


def fast_read_slots(page):
    """Zwraca dict: { 'HH:MM-HH:MM': (used, total) }"""
    out = {}
    texts = page.locator(r"text=/\b\d{2}:\d{2}-\d{2}:\d{2}\b/").all_inner_texts()
    for t in texts:
        m = SLOT_RE.search(t)
        if not m:
            continue
        key = f"{m.group(1)}-{m.group(2)}"
        out[key] = (int(m.group(3)), int(m.group(4)))
    return out


def click_standardowe(page, load_timeout):
    """Jedno kliknięcie STANDARDOWE = refresh."""
    try:
        btn = page.locator("text=STANDARDOWE").first
        btn.scroll_into_view_if_needed()
        btn.click(timeout=1500)
        wait_for_slots_loaded(page, load_timeout)
        return True
    except Exception:
        return False


def toast_no_slots(page):
    """Toast: Brak dostępnych slotów."""
    try:
        loc = page.locator("text=Brak dostępnych slotów").first
        if loc.count() and loc.is_visible():
            # spróbuj zamknąć X
            try:
                page.locator("button:has-text('×')").first.click(timeout=300)
            except Exception:
                pass
            return True
    except Exception:
        pass
    return False


def success_visible(page):
    """Sukces jeśli pojawi się komunikat o wysłaniu do kierowcy (widoczny)."""
    try:
        loc = page.locator("text=Powiadomienie zostało wysłane do kierowcy").first
        return loc.count() > 0 and loc.is_visible()
    except Exception:
        return False


def success_confirmed(page, timeout_ms):
    """Sukces tylko gdy pojawi się komunikat o wysłaniu do kierowcy."""
    try:
        page.locator("text=Powiadomienie zostało wysłane do kierowcy").wait_for(timeout=int(timeout_ms))
        return True
    except Exception:
        return False


def confirm_loop_fast(page, max_clicks=25):
    """
    v4.2.3: Klikaj TAK/OK najszybciej jak się da, dopóki:
      - widzimy przyciski TAK/OK
      - lub do osiągnięcia limitu kliknięć
    Priorytet: TAK, potem OK/Ok.
    Zatrzymaj, jeśli widoczny jest sukces.
    """
    for _ in range(max_clicks):
        if success_visible(page):
            return

        clicked = False

        # Priorytet: TAK (bez wcześniejszego count/is_visible/is_enabled,
        # żeby minimalizować koszt odpytywania DOM i kliknąć jak najszybciej)
        try:
            page.get_by_role("button", name="Tak").first.click(timeout=160)
            clicked = True
        except Exception:
            pass

        if not clicked:
            # OK / Ok
            for txt in ("OK", "Ok"):
                try:
                    page.get_by_role("button", name=txt).first.click(timeout=160)
                    clicked = True
                    break
                except Exception:
                    pass

        if not clicked:
            return

        # minimalna pauza na repaint DOM
        time.sleep(0.005)


def get_selected_day_number(page):
    """Best-effort: próba ustalenia zaznaczonego dnia w kalendarzu."""
    selectors = [
        "[aria-current='date']",
        "button.active",
        "a.active",
        "td.active button",
        "td.active a",
        "td.active",
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible():
                txt = (loc.inner_text() or "").strip()
                m = re.search(r"\b(\d{1,2})\b", txt)
                if m:
                    return int(m.group(1))
        except Exception:
            pass
    return None


def click_day_by_coordinates(page, day: dt.date, load_to: int):
    """Kliknięcie dnia w kalendarzu po współrzędnych."""
    label = str(day.day)

    month_box = page.locator(
        "xpath=//div[contains(., 'styczeń') or contains(., 'luty') or contains(., 'marzec') or contains(., 'kwiecień') or "
        "contains(., 'maj') or contains(., 'czerwiec') or contains(., 'lipiec') or contains(., 'sierpień') or "
        "contains(., 'wrzesień') or contains(., 'październik') or contains(., 'listopad') or contains(., 'grudzień')]"
    ).first

    cand = month_box.locator(f"xpath=.//*[normalize-space(text())='{label}']").first
    if cand.count() == 0:
        return False

    try:
        cand.scroll_into_view_if_needed()
        box = cand.bounding_box()
        if not box:
            return False
        cx = box["x"] + box["width"] / 2
        cy = box["y"] + box["height"] / 2
        page.mouse.click(cx, cy)
        wait_for_slots_loaded(page, load_to)
        return True
    except Exception:
        return False


def ensure_day_selected(page, day: dt.date, load_to: int, tries: int = 5):
    """Wymusza przejście na konkretny dzień – z retry."""
    target = day.day
    for _ in range(tries):
        cur = get_selected_day_number(page)
        if cur == target:
            return True
        ok = click_day_by_coordinates(page, day, load_to)
        time.sleep(0.10)
        cur2 = get_selected_day_number(page)
        if ok and cur2 == target:
            return True
    return False


# ------------------ WORKER ------------------

class Worker(threading.Thread):
    def __init__(self, ui):
        super().__init__(daemon=True)
        self.ui = ui
        self.stop_evt = threading.Event()

    def run(self):
        try:
            self.logic()
        except Exception as e:
            self.ui.log(f"[FATAL] {e}")
            self.ui.popup("Błąd", str(e))

    def stop(self):
        self.stop_evt.set()

    def logic(self):
        ui = self.ui
        start_d, start_h, end_d, end_h = ui.get_range()
        poll_s, load_to, success_to = ui.get_params()

        # lista dni w zakresie
        days = []
        d = start_d
        while d <= end_d:
            days.append(d)
            d += dt.timedelta(days=1)

        with sync_playwright() as p:
            ui.log("[PW] Łączenie z Chrome CDP...")
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            ctx = browser.contexts[0]
            page = ctx.pages[0]

            ui.log(f"[OK] Strona: {page.url}")
            ensure_slot_screen(page)

            while not self.stop_evt.is_set():
                for day in days:
                    if self.stop_evt.is_set():
                        return

                    # 1) ZAWSZE ustaw właściwy dzień
                    if not ensure_day_selected(page, day, load_to):
                        ui.log(f"[WARN] Nie udało się ustawić dnia {day.isoformat()} – pomijam i wracam do pętli.")
                        continue

                    # 2) Odświeżanie
                    if len(days) == 1:
                        click_standardowe(page, load_to)
                    else:
                        click_day_by_coordinates(page, day, load_to)

                    # 3) Safety: jeśli UI przeskoczyło dzień, nie klikamy slotów
                    cur = get_selected_day_number(page)
                    if cur is not None and cur != day.day:
                        ui.log(f"[SAFE] Aktualnie zaznaczony dzień={cur}, oczekiwany={day.day}. Nie klikam slotów.")
                        continue

                    slots = fast_read_slots(page)

                    # 4) Sprawdź sloty tylko dla tego dnia i dla zakresu godzin
                    for h in ui.iter_hours_for_day(day):
                        if self.stop_evt.is_set():
                            return

                        slot_key = f"{h:02d}:00-{h:02d}:59"
                        if slot_key not in slots:
                            continue

                        used, total = slots[slot_key]
                        if used < total:
                            ui.log(f"[TRY] {day.isoformat()} {slot_key} {used}/{total}")

                            # klik slot + potwierdzenia
                            self.try_slot(page, slot_key, load_to, success_to)

                            # jeśli toast "brak slotów" -> wracamy do odświeżania
                            if toast_no_slots(page):
                                ui.log("[INFO] Toast 'Brak dostępnych slotów' – kontynuuję odświeżanie.")
                                continue

                            # sukces tylko po komunikacie o wysłaniu do kierowcy
                            if success_confirmed(page, success_to):
                                ui.emit_notification("slot_success")
                                ui.log("[SUCCESS] Awizacja utworzona (wysłane do kierowcy).")
                                return

                time.sleep(max(0.05, float(poll_s)))

    def try_slot(self, page, slot_key, load_to, success_to):
        """
        Kliknięcie slotu + agresywna pętla potwierdzeń TAK/OK (v4.2.3).
        """
        try:
            # Preferujemy button/a/role=button zawierające slot_key
            candidates = page.locator(
                f"button:has-text('{slot_key}'), a:has-text('{slot_key}'), [role='button']:has-text('{slot_key}')"
            )
            n = candidates.count()

            # Fallback: czasem kafel jest div/span - łapiemy tekst regexem z "xx/yy"
            if n == 0:
                candidates = page.locator(f"text=/{re.escape(slot_key)}\\s+\\d+\\/\\d+/")
                n = candidates.count()

            if n == 0:
                self.ui.log(f"[WARN] Nie znalazłem kafelka dla slotu: {slot_key}")
                return

            clicked = False
            for i in range(min(n, 12)):
                el = candidates.nth(i)
                try:
                    if el.is_visible():
                        el.scroll_into_view_if_needed()
                        el.click(timeout=2000)
                        clicked = True
                        break
                except Exception:
                    continue

            if not clicked:
                self.ui.log(f"[WARN] Kafelek slotu jest, ale nie udało się kliknąć: {slot_key}")
                return

            # Faza 1 (ultra-fast): od razu próbujemy klikać dialogi,
            # bez czekania na pełne dociągnięcie UI.
            confirm_loop_fast(page, max_clicks=20)

            # Faza 2: jeśli UI jeszcze ładuje, dokończ po załadowaniu.
            wait_for_slots_loaded(page, load_to)
            confirm_loop_fast(page, max_clicks=20)

            # jeśli sukces pojawi się szybko, kończymy od razu
            if success_visible(page):
                return

            # dodatkowo: jeszcze krótko poczekaj na sukces (minimalnie)
            success_confirmed(page, min(800, int(success_to)))

        except Exception as e:
            self.ui.log(f"[WARN] Kliknięcie slotu nie powiodło się: {e}")


# ------------------ UI ------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"VBS klikacz NTQ – {VERSION}")
        self.geometry("820x520")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.tab_main = ttk.Frame(nb)
        self.tab_params = ttk.Frame(nb)
        self.tab_notifications = ttk.Frame(nb)
        self.tab_info = ttk.Frame(nb)

        nb.add(self.tab_main, text="Rezerwacja")
        nb.add(self.tab_params, text="Parametry")
        nb.add(self.tab_notifications, text="Powiadomienia")
        nb.add(self.tab_info, text="Info")

        self.machine_id = generate_machine_id()

        # stan licencji w UI
        self.lic_status = tk.StringVar(value="(nie sprawdzono)")
        self.lic_valid_to = tk.StringVar(value="")
        self.lic_checked = tk.StringVar(value="")

        self.build_main()
        self.build_params()
        self.build_notifications()
        self.load_settings()
        self.build_info()

        self.worker = None

        # pierwsze sprawdzenie przy uruchomieniu (bez zamykania)
        self.refresh_license_status(close_on_invalid=False)

    # ---------- UI builders ----------

    def build_main(self):
        f = self.tab_main

        # Prosty stan ON/OFF (bez efektu dźwiękowego na tym etapie).
        self.sound_enabled = True

        # domyślny zakres: następna pełna godzina -> +1h
        now = dt.datetime.now()
        start_dt = now.replace(minute=0, second=0, microsecond=0) + dt.timedelta(hours=1)
        end_dt = start_dt + dt.timedelta(hours=1)

        ttk.Label(f, text="Zakres").pack(anchor="w", padx=10, pady=5)

        r = ttk.Frame(f)
        r.pack(anchor="w", padx=10)

        self.od_date = tk.StringVar(value=start_dt.strftime("%Y-%m-%d"))
        self.od_hour = tk.IntVar(value=start_dt.hour)
        self.do_date = tk.StringVar(value=end_dt.strftime("%Y-%m-%d"))
        self.do_hour = tk.IntVar(value=end_dt.hour)

        ttk.Label(r, text="Od").pack(side="left")
        ttk.Entry(r, width=12, textvariable=self.od_date).pack(side="left", padx=2)
        ttk.Spinbox(r, from_=0, to=23, width=3, textvariable=self.od_hour).pack(side="left")

        ttk.Label(r, text="Do").pack(side="left", padx=10)
        ttk.Entry(r, width=12, textvariable=self.do_date).pack(side="left", padx=2)
        ttk.Spinbox(r, from_=0, to=23, width=3, textvariable=self.do_hour).pack(side="left")

        b = ttk.Frame(f)
        b.pack(anchor="w", padx=10, pady=10)

        ttk.Button(b, text="START", command=self.start).pack(side="left", padx=5)
        ttk.Button(b, text="STOP", command=self.stop).pack(side="left", padx=5)
        self.sound_btn = tk.Button(
            b,
            text="DZWIEK",
            width=10,
            command=self.toggle_sound,
            relief="raised",
            bd=1,
        )
        self.sound_btn.pack(side="left", padx=5)
        self.update_sound_button_style()

        self.log_box = tk.Text(f, height=16)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=5)

    def build_params(self):
        f = self.tab_params

        self.poll_s = tk.IntVar(value=1)
        self.load_to = tk.IntVar(value=5000)
        self.success_to = tk.IntVar(value=4000)

        ttk.Label(f, text="Parametry pracy (edytowalne):").pack(anchor="w", padx=10, pady=8)

        for txt, var in (
            ("Interwał pętli (s)", self.poll_s),
            ("Timeout ładowania slotów (ms)", self.load_to),
            ("Timeout sukcesu (ms)", self.success_to),
        ):
            r = ttk.Frame(f)
            r.pack(anchor="w", padx=10, pady=6)
            ttk.Label(r, text=txt, width=28).pack(side="left")
            ttk.Entry(r, width=10, textvariable=var).pack(side="left", padx=5)

        ttk.Label(
            f,
            text="Uwaga: wartości są pobierane w momencie kliknięcia START.",
            foreground="#444"
        ).pack(anchor="w", padx=10, pady=8)

    def build_notifications(self):
        f = self.tab_notifications

        ttk.Label(
            f,
            text="Ustawienia powiadomień",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", padx=10, pady=(10, 8))

        self.notification_settings = {
            "start_stop": {
                "label": "Start/Stop programu",
                "enabled": True,
                "volume": tk.IntVar(value=80),
                "file": tk.StringVar(value="brak"),
                "button": None,
            },
            "slot_success": {
                "label": "Udane kliknięcie okienka",
                "enabled": True,
                "volume": tk.IntVar(value=80),
                "file": tk.StringVar(value="brak"),
                "button": None,
            },
        }

        self._build_notification_row(f, "start_stop")
        self._build_notification_row(f, "slot_success")

    def _build_notification_row(self, parent, key):
        cfg = self.notification_settings[key]

        box = ttk.LabelFrame(parent, text=cfg["label"])
        box.pack(fill="x", padx=10, pady=8)

        row = ttk.Frame(box)
        row.pack(fill="x", padx=10, pady=10)

        ttk.Label(row, text="Status:").pack(side="left")

        btn = tk.Button(
            row,
            text="",
            width=7,
            command=lambda k=key: self.toggle_notification(k),
            relief="raised",
            bd=1,
        )
        btn.pack(side="left", padx=(6, 12))
        cfg["button"] = btn

        ttk.Label(row, text="Głośność:").pack(side="left")
        tk.Scale(
            row,
            from_=0,
            to=100,
            orient="horizontal",
            showvalue=True,
            length=170,
            variable=cfg["volume"],
            command=lambda _v, k=key: self.on_notification_volume_change(k),
        ).pack(side="left", padx=(6, 12))

        ttk.Label(row, text="Plik:").pack(side="left")
        ttk.Label(row, textvariable=cfg["file"], width=18).pack(side="left", padx=(6, 6))
        ttk.Button(
            row,
            text="Wybierz",
            command=lambda k=key: self.select_notification_file(k),
        ).pack(side="left")
        ttk.Button(
            row,
            text="Test",
            command=lambda k=key: self.test_notification_sound(k),
        ).pack(side="left", padx=(8, 0))

        self.update_notification_button_style(key)

    def update_notification_button_style(self, key):
        cfg = self.notification_settings[key]
        btn = cfg["button"]
        if not btn:
            return

        if cfg["enabled"]:
            btn.config(
                text="ON",
                bg="#22C55E",
                activebackground="#16A34A",
                fg="white",
                activeforeground="white",
                disabledforeground="white",
            )
        else:
            btn.config(
                text="OFF",
                bg="#DC2626",
                activebackground="#B91C1C",
                fg="white",
                activeforeground="white",
                disabledforeground="white",
            )

    def toggle_notification(self, key):
        cfg = self.notification_settings[key]
        cfg["enabled"] = not cfg["enabled"]
        self.update_notification_button_style(key)
        self.save_settings()

    def select_notification_file(self, key):
        path = filedialog.askopenfilename(title="Wybierz plik dźwięku")
        if path:
            self.notification_settings[key]["file"].set(path)
            self.save_settings()

    def on_notification_volume_change(self, key):
        _ = key
        self.save_settings()

    def test_notification_sound(self, key):
        cfg = self.notification_settings.get(key)
        if not cfg:
            return

        if not self.sound_enabled:
            self.log(f"[TEST] {cfg['label']} - globalny DZWIEK jest OFF")
            return

        if not cfg["enabled"]:
            self.log(f"[TEST] {cfg['label']} - to powiadomienie jest OFF")
            return

        sound_file = cfg["file"].get().strip()
        if not sound_file or sound_file == "brak":
            self.log(f"[TEST] {cfg['label']} - brak wybranego pliku")
            return

        self.play_sound_file_once(sound_file, cfg["label"], int(cfg["volume"].get()))

    def play_sound_file_once(self, sound_file: str, label: str, volume_percent: int = 100):
        try:
            p = Path(sound_file)
            if not p.exists():
                self.log(f"[TEST] {label} - plik nie istnieje: {sound_file}")
                return

            if platform.system() != "Windows":
                self.log(f"[TEST] {label} - odtwarzanie działa tylko na Windows")
                return

            # Odtwarzanie plików (mp3/wav/inne wspierane przez MediaPlayer)
            # z uwzględnieniem głośności z UI.
            escaped_path = str(p).replace("'", "''")
            volume = max(0.0, min(1.0, float(volume_percent) / 100.0))
            ps_cmd = (
                "Add-Type -AssemblyName PresentationCore; "
                "$player = New-Object System.Windows.Media.MediaPlayer; "
                f"$player.Open([Uri]'{escaped_path}'); "
                f"$player.Volume = {volume}; "
                "$player.Play(); "
                "while (-not $player.NaturalDuration.HasTimeSpan) { Start-Sleep -Milliseconds 50 }; "
                "Start-Sleep -Milliseconds ([Math]::Ceiling($player.NaturalDuration.TimeSpan.TotalMilliseconds)); "
                "$player.Stop(); $player.Close();"
            )
            res = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_cmd],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=30,
            )
            if res.returncode != 0:
                self.log(f"[TEST] {label} - nie udało się odtworzyć pliku")
                return

            self.log(f"[TEST] {label} - odtworzono 1 raz")

        except Exception as e:
            self.log(f"[TEST] {label} - błąd odtwarzania: {e}")

    def is_notification_sound_enabled(self, key):
        cfg = self.notification_settings.get(key)
        if not cfg:
            return False
        return self.sound_enabled and bool(cfg["enabled"])

    def emit_notification(self, key):
        cfg = self.notification_settings.get(key)
        if not cfg:
            return

        if self.is_notification_sound_enabled(key):
            sound_file = str(cfg["file"].get()).strip()
            volume = int(cfg["volume"].get())

            self.log(
                f"[NOTIFY] {cfg['label']} ON | glosnosc={int(cfg['volume'].get())} | plik={cfg['file'].get()}"
            )

            if sound_file and sound_file != "brak":
                threading.Thread(
                    target=self.play_sound_file_once,
                    args=(sound_file, cfg["label"], volume),
                    daemon=True,
                ).start()
            else:
                self.log(f"[NOTIFY] {cfg['label']} - brak wybranego pliku")
        else:
            self.log(f"[NOTIFY] {cfg['label']} wyciszone")

    def save_settings(self):
        try:
            payload = {
                "sound_enabled": self.sound_enabled,
                "notifications": {},
            }
            for key, cfg in self.notification_settings.items():
                payload["notifications"][key] = {
                    "enabled": bool(cfg["enabled"]),
                    "volume": int(cfg["volume"].get()),
                    "file": str(cfg["file"].get()),
                }

            SETTINGS_FILE.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:
            self.log(f"[SETTINGS] Nie udało się zapisać ustawień: {e}")

    def load_settings(self):
        if not SETTINGS_FILE.exists():
            self.save_settings()
            return

        try:
            raw = SETTINGS_FILE.read_text(encoding="utf-8")
            data = json.loads(raw)

            self.sound_enabled = bool(data.get("sound_enabled", self.sound_enabled))
            self.update_sound_button_style()

            incoming = data.get("notifications", {})
            for key, cfg in self.notification_settings.items():
                src = incoming.get(key, {}) if isinstance(incoming, dict) else {}
                cfg["enabled"] = bool(src.get("enabled", cfg["enabled"]))
                cfg["volume"].set(int(src.get("volume", cfg["volume"].get())))
                cfg["file"].set(str(src.get("file", cfg["file"].get())))
                self.update_notification_button_style(key)

        except Exception as e:
            self.log(f"[SETTINGS] Nie udało się wczytać ustawień, używam domyślnych: {e}")
            self.save_settings()

    def build_info(self):
        f = self.tab_info

        ttk.Label(
            f,
            text="NTQ Intermodal – VBS Klikacz",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="nw", padx=10, pady=(10, 2))

        ttk.Label(f, text=f"Wersja: {VERSION}").pack(anchor="nw", padx=10)
        ttk.Label(f, text=f"Build: {BUILD_TIME}").pack(anchor="nw", padx=10, pady=(0, 10))

        box_changes = ttk.LabelFrame(f, text="Zmiany w tej wersji")
        box_changes.pack(fill="x", padx=10, pady=10)
        txt = "\n".join([f"• {x}" for x in CHANGELOG])
        ttk.Label(box_changes, text=txt, justify="left").pack(anchor="w", padx=10, pady=8)

        box_lic = ttk.LabelFrame(f, text="Licencja")
        box_lic.pack(fill="x", padx=10, pady=10)

        row1 = ttk.Frame(box_lic)
        row1.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(row1, text="ID komputera:", width=14).pack(side="left")
        self.machine_id_var = tk.StringVar(value=self.machine_id)
        ent = ttk.Entry(row1, textvariable=self.machine_id_var, width=24, state="readonly")
        ent.pack(side="left", padx=5)
        ttk.Button(row1, text="Kopiuj", command=self.copy_machine_id).pack(side="left", padx=8)

        row2 = ttk.Frame(box_lic)
        row2.pack(fill="x", padx=10, pady=4)
        ttk.Label(row2, text="Status:", width=14).pack(side="left")
        ttk.Label(row2, textvariable=self.lic_status).pack(side="left")

        row3 = ttk.Frame(box_lic)
        row3.pack(fill="x", padx=10, pady=4)
        ttk.Label(row3, text="Ważna do:", width=14).pack(side="left")
        ttk.Label(row3, textvariable=self.lic_valid_to).pack(side="left")

        row4 = ttk.Frame(box_lic)
        row4.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Label(row4, text="Sprawdzono:", width=14).pack(side="left")
        ttk.Label(row4, textvariable=self.lic_checked).pack(side="left")

        ttk.Button(
            f,
            text="Sprawdź licencję teraz",
            command=lambda: self.refresh_license_status(close_on_invalid=False)
        ).pack(anchor="nw", padx=10, pady=(0, 10))

    # ---------- helpers ----------

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")

    def popup(self, title, msg):
        self.after(0, lambda: messagebox.showerror(title, msg))

    def copy_machine_id(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self.machine_id_var.get())
            self.update()
        except Exception:
            pass

    def get_range(self):
        sd = dt.datetime.strptime(self.od_date.get(), "%Y-%m-%d").date()
        ed = dt.datetime.strptime(self.do_date.get(), "%Y-%m-%d").date()
        return sd, int(self.od_hour.get()), ed, int(self.do_hour.get())

    def iter_hours_for_day(self, d):
        sd, sh, ed, eh = self.get_range()
        if sd == ed:
            if eh < sh:
                return range(sh, 24)
            return range(sh, eh + 1)
        if d == sd:
            return range(sh, 24)
        if d == ed:
            return range(0, eh + 1)
        return range(0, 24)

    def get_params(self):
        try:
            poll = max(0, float(self.poll_s.get()))
        except Exception:
            poll = 1.0
        try:
            load_to = max(500, int(self.load_to.get()))
        except Exception:
            load_to = 5000
        try:
            success_to = max(500, int(self.success_to.get()))
        except Exception:
            success_to = 4000
        return poll, load_to, success_to

    def update_sound_button_style(self):
        if self.sound_enabled:
            self.sound_btn.config(
                bg="#22C55E",
                activebackground="#16A34A",
                fg="white",
                activeforeground="white",
                disabledforeground="white",
            )
        else:
            self.sound_btn.config(
                bg="#DC2626",
                activebackground="#B91C1C",
                fg="white",
                activeforeground="white",
                disabledforeground="white",
            )

    def toggle_sound(self):
        self.sound_enabled = not self.sound_enabled
        self.update_sound_button_style()
        state = "ON" if self.sound_enabled else "OFF"
        self.log(f"[UI] DZWIEK: {state}")
        self.save_settings()

    # ---------- license ----------

    def refresh_license_status(self, close_on_invalid: bool):
        """
        Sprawdza status licencji przez Apps Script, aktualizuje Info.
        Jeśli close_on_invalid=True i status != OK -> zamyka program.
        """
        hwid = self.machine_id
        checked_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            data = fetch_license_status(hwid)
            status = str(data.get("status", "UNKNOWN")).strip().upper()
            valid_to = data.get("valid_to", "")

            self.lic_status.set(status)
            self.lic_valid_to.set(str(valid_to))
            self.lic_checked.set(checked_at)

            if close_on_invalid and not is_license_valid(status):
                messagebox.showerror(
                    "Licencja",
                    f"Licencja nieważna: {status}\nWażna do: {valid_to}"
                )
                self.destroy()

        except Exception as e:
            self.lic_status.set("ERROR")
            self.lic_valid_to.set("")
            self.lic_checked.set(checked_at)

            if close_on_invalid:
                messagebox.showerror(
                    "Licencja",
                    f"Nie udało się sprawdzić licencji (ERROR).\n{e}"
                )
                self.destroy()

    # ---------- controls ----------

    def start(self):
        self.emit_notification("start_stop")

        # START ma sprawdzić licencję i zamknąć program jeśli nieważna
        self.refresh_license_status(close_on_invalid=True)
        if not self.winfo_exists():
            return

        if self.worker and self.worker.is_alive():
            return

        self.log("[UI] START")
        self.worker = Worker(self)
        self.worker.start()
        self.emit_notification("start_stop")

    def stop(self):
        self.emit_notification("start_stop")

        # STOP ma też sprawdzić licencję i zamknąć program jeśli nieważna
        self.refresh_license_status(close_on_invalid=True)
        if not self.winfo_exists():
            return

        if self.worker:
            self.worker.stop()
            self.log("[UI] STOP")
            self.emit_notification("start_stop")


if __name__ == "__main__":
    App().mainloop()
