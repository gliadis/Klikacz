# -*- coding: utf-8 -*-
"""
VBS klikacz NTQ – v4.2.2 (STABLE)

Zmiany v4.2.2:
• FIX: agresywne „dobijanie” dialogów po kliknięciu slotu (TAK/OK) – pętla klika do skutku, dopóki przyciski są aktywne.
• Zachowane: poprawione klikanie slotów z tekstem typu "02:00-02:59 44/130".
• Info: wersja/build/zmiany/status licencji, ID komputera + Kopiuj.
• Refresh 1-dniowy: klik STANDARDOWE.

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
import tkinter as tk
from tkinter import ttk, messagebox
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from playwright.sync_api import sync_playwright

VERSION = "v4.2.2"
BUILD_TIME = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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
    v4.2.2: Klikaj TAK/OK najszybciej jak się da, dopóki:
      - widzimy przyciski TAK/OK
      - lub do osiągnięcia limitu kliknięć
    Priorytet: TAK, potem OK/Ok.
    Zatrzymaj, jeśli widoczny jest sukces.
    """
    for _ in range(max_clicks):
        if success_visible(page):
            return

        clicked = False

        # Priorytet: TAK
        try:
            b = page.locator("button:has-text('Tak')").first
            if b.count() and b.is_visible() and b.is_enabled():
                b.click(timeout=600)
                clicked = True
        except Exception:
            pass

        if not clicked:
            # OK / Ok
            for txt in ("OK", "Ok"):
                try:
                    b2 = page.locator(f"button:has-text('{txt}')").first
                    if b2.count() and b2.is_visible() and b2.is_enabled():
                        b2.click(timeout=600)
                        clicked = True
                        break
                except Exception:
                    pass

        if not clicked:
            return

        # minimalna pauza na repaint DOM (bardzo krótka)
        time.sleep(0.02)


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
                                ui.log("[SUCCESS] Awizacja utworzona (wysłane do kierowcy).")
                                return

                time.sleep(max(0.05, float(poll_s)))

    def try_slot(self, page, slot_key, load_to, success_to):
        """
        Kliknięcie slotu + agresywna pętla potwierdzeń TAK/OK (v4.2.2).
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

            wait_for_slots_loaded(page, load_to)

            # v4.2.2: agresywne dobijanie dialogów TAK/OK
            # klikamy do skutku dopóki są aktywne przyciski lub do limitu
            confirm_loop_fast(page, max_clicks=30)

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
        self.tab_info = ttk.Frame(nb)

        nb.add(self.tab_main, text="Rezerwacja")
        nb.add(self.tab_params, text="Parametry")
        nb.add(self.tab_info, text="Info")

        self.machine_id = generate_machine_id()

        # stan licencji w UI
        self.lic_status = tk.StringVar(value="(nie sprawdzono)")
        self.lic_valid_to = tk.StringVar(value="")
        self.lic_checked = tk.StringVar(value="")

        self.build_main()
        self.build_params()
        self.build_info()

        self.worker = None

        # pierwsze sprawdzenie przy uruchomieniu (bez zamykania)
        self.refresh_license_status(close_on_invalid=False)

    # ---------- UI builders ----------

    def build_main(self):
        f = self.tab_main

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
        # START ma sprawdzić licencję i zamknąć program jeśli nieważna
        self.refresh_license_status(close_on_invalid=True)
        if not self.winfo_exists():
            return

        if self.worker and self.worker.is_alive():
            return

        self.log("[UI] START")
        self.worker = Worker(self)
        self.worker.start()

    def stop(self):
        # STOP ma też sprawdzić licencję i zamknąć program jeśli nieważna
        self.refresh_license_status(close_on_invalid=True)
        if not self.winfo_exists():
            return

        if self.worker:
            self.worker.stop()
            self.log("[UI] STOP")


if __name__ == "__main__":
    App().mainloop()
