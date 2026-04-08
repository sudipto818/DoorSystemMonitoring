"""
control_app.py
─────────────────────────────────────────────────────
Owner's control panel for the Door-Sign-Status system.
Can be opened / closed freely — the display_app.py
keeps running and reads from the shared SQLite database.

PROFESSIONAL LIGHT-TONE REDESIGN — clean, muted colours,
minimal card layout, refined typography.
"""

import customtkinter as ctk
from tkinter import ttk, messagebox
import threading
import time
import os
import sys
from datetime import datetime, timedelta, timezone

from db_manager import (
    init_db, read_status, write_status, update_return_time,
    get_visitors, delete_visitor,
    get_priority, STATUS_PRIORITY,
)

try:
    import requests as http_requests
    import recurring_ical_events
    import icalendar
    ICS_AVAILABLE = True
except ImportError:
    ICS_AVAILABLE = False


def _get_base_dir() -> str:
    """Return the folder where the app lives (works for .py and .exe)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _ics_url_path() -> str:
    """Path where the ICS URL is saved locally."""
    return os.path.join(_get_base_dir(), "outlook_ics_url.txt")


def _load_ics_url() -> str:
    path = _ics_url_path()
    if os.path.exists(path):
        with open(path, "r") as f:
            return f.read().strip()
    return ""


def _save_ics_url(url: str):
    with open(_ics_url_path(), "w") as f:
        f.write(url.strip())


# ═══════════════════════════════════════════════════════════════════════
#  APPEARANCE — Professional Light Theme
# ═══════════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Typography
FONT_TITLE   = ("Segoe UI", 22, "bold")
FONT_SECTION = ("Segoe UI", 13, "bold")
FONT_BODY    = ("Segoe UI", 13)
FONT_SMALL   = ("Segoe UI", 11)
FONT_BTN     = ("Segoe UI", 12, "bold")
FONT_TINY    = ("Segoe UI", 10)

# Colours — Light, professional palette
BG_BASE    = "#f8f9fa"
BG_CARD    = "#ffffff"
BG_INPUT   = "#f1f3f5"

FG_PRIMARY = "#212529"
FG_MUTED   = "#6c757d"
FG_DIM     = "#adb5bd"

ACCENT     = "#4263eb"
BORDER     = "#dee2e6"

BTN = {
    "green":  ("#2b8a3e", "#37a547"),
    "amber":  ("#e67700", "#f08c00"),
    "red":    ("#c92a2a", "#e03131"),
    "purple": ("#7048e8", "#7c5ce8"),
    "blue":   ("#4263eb", "#4c6ef5"),
    "teal":   ("#0c8599", "#1098ad"),
    "orange": ("#d9480f", "#e8590c"),
    "gray":   ("#868e96", "#adb5bd"),
    "slate":  ("#495057", "#6c757d"),
    "light":  ("#e9ecef", "#dee2e6"),
}


# ═══════════════════════════════════════════════════════════════════════
#  HELPER — Section Card
# ═══════════════════════════════════════════════════════════════════════

def section_card(parent, **kwargs):
    return ctk.CTkFrame(
        parent, fg_color=BG_CARD, corner_radius=10,
        border_width=1, border_color=BORDER, **kwargs
    )


def section_label(parent, text, icon=""):
    ctk.CTkLabel(
        parent, text=f"{icon}  {text}" if icon else text,
        font=FONT_SECTION, text_color=FG_MUTED
    ).pack(anchor="w", padx=16, pady=(14, 8))


# ═══════════════════════════════════════════════════════════════════════
#  CONTROL PANEL APPLICATION
# ═══════════════════════════════════════════════════════════════════════

class ControlApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        init_db()

        self.title("Door Sign — Control Panel")
        self.geometry("820x980")
        self.minsize(750, 880)
        self.configure(fg_color=BG_BASE)

        # ICS thread state
        self._ics_running = False
        self._ics_thread = None
        self._ics_meeting_active = False
        self._ics_status_text = "Not connected"

        # Scrollable container
        self.scroll = ctk.CTkScrollableFrame(
            self, fg_color=BG_BASE, corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=FG_DIM
        )
        self.scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # Build all sections
        self._build_header()
        self._build_status_bar()
        self._build_location_card()
        self._build_status_card()
        self._build_custom_card()
        self._build_time_card()
        self._build_visitor_card()
        self._build_schedule_card()
        self._build_outlook_card()

        # Initial data
        self._refresh_current_status()
        self._tick()

        # Auto-connect if URL was previously saved
        saved_url = _load_ics_url()
        if ICS_AVAILABLE and saved_url:
            self.ics_url_entry.insert(0, saved_url)
            self._start_ics_thread(saved_url)

    # ═══════════════════════ HEADER ═══════════════════════════

    def _build_header(self):
        frame = ctk.CTkFrame(self.scroll, fg_color="transparent")
        frame.pack(fill="x", padx=24, pady=(18, 4))

        ctk.CTkLabel(
            frame, text="Door Sign Control",
            font=FONT_TITLE, text_color=FG_PRIMARY
        ).pack(side="left")

        ctk.CTkLabel(
            frame, text="v1.0",
            font=FONT_TINY, text_color=FG_DIM
        ).pack(side="left", padx=(10, 0), pady=(6, 0))

    # ═══════════════════ STATUS BAR ═══════════════════════════

    def _build_status_bar(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=(10, 6))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(
            inner, text="CURRENTLY DISPLAYING",
            font=FONT_TINY, text_color=FG_DIM
        ).pack(anchor="w")

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x", pady=(4, 0))

        self.status_dot = ctk.CTkLabel(
            row, text="●", font=("Segoe UI", 14),
            text_color=ACCENT, width=20
        )
        self.status_dot.pack(side="left")

        self.current_lbl = ctk.CTkLabel(
            row, text="…", font=("Segoe UI", 17, "bold"),
            text_color=FG_PRIMARY
        )
        self.current_lbl.pack(side="left", padx=(6, 0))

        self.source_lbl = ctk.CTkLabel(
            row, text="", font=FONT_TINY, text_color=FG_DIM
        )
        self.source_lbl.pack(side="right")

    def _refresh_current_status(self):
        row = read_status()
        self.current_lbl.configure(text=row["current_status"])
        src = row.get("source", "manual")
        self.source_lbl.configure(text=f"source: {src}")
        base = row["current_status"].split(" — ")[0].strip()
        dot_colors = {
            "Available": "#2b8a3e", "Please Knock": "#e67700",
            "In a Meeting": "#c92a2a", "Do Not Disturb": "#7048e8",
            "Working remotely": "#4263eb", "Back soon": "#0c8599",
            "Busy": "#d9480f", "Out of office": "#868e96",
        }
        self.status_dot.configure(text_color=dot_colors.get(base, ACCENT))

    # ════════════════ LOCATION CARD ═══════════════════════════

    def _build_location_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        section_label(card, "Location", "📍")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 14))

        self.location_var = ctk.StringVar(value="inside")

        self.inside_btn = ctk.CTkButton(
            row, text="Inside Office", font=FONT_BTN,
            fg_color=ACCENT, hover_color=BTN["blue"][1],
            text_color="#ffffff", height=40, corner_radius=8,
            command=lambda: self._set_location("inside")
        )
        self.inside_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.outside_btn = ctk.CTkButton(
            row, text="Outside Office", font=FONT_BTN,
            fg_color=BTN["light"][0], hover_color=BTN["light"][1],
            text_color=FG_MUTED, height=40, corner_radius=8,
            command=lambda: self._set_location("outside")
        )
        self.outside_btn.pack(side="left", fill="x", expand=True, padx=(6, 0))

    def _set_location(self, loc: str):
        self.location_var.set(loc)
        if loc == "inside":
            self.inside_btn.configure(fg_color=ACCENT, text_color="#ffffff")
            self.outside_btn.configure(fg_color=BTN["light"][0], text_color=FG_MUTED)
        else:
            self.inside_btn.configure(fg_color=BTN["light"][0], text_color=FG_MUTED)
            self.outside_btn.configure(fg_color=ACCENT, text_color="#ffffff")
        self._build_status_buttons(loc)
        self._toggle_time_section(loc)

    # ═══════════════ STATUS BUTTONS CARD ══════════════════════

    def _build_status_card(self):
        self.status_card = section_card(self.scroll)
        self.status_card.pack(fill="x", padx=24, pady=6)

        section_label(self.status_card, "Quick Status", "⚡")

        self.btn_container = ctk.CTkFrame(self.status_card, fg_color="transparent")
        self.btn_container.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_frame = None
        self._build_status_buttons("inside")

    def _build_status_buttons(self, location: str):
        if self.btn_frame is not None:
            self.btn_frame.destroy()

        self.btn_frame = ctk.CTkFrame(self.btn_container, fg_color="transparent")
        self.btn_frame.pack(fill="x")

        if location == "inside":
            buttons = [
                ("Available",      "Available",      "slate"),
                ("Please Knock",   "Please Knock",   "slate"),
                ("In a Meeting",   "In a Meeting",   "slate"),
                ("Do Not Disturb", "Do Not Disturb", "slate"),
            ]
        else:
            buttons = [
                ("Working Remotely", "Working remotely", "slate"),
                ("Back Soon",        "Back soon",        "slate"),
                ("Out of Office",    "Out of office",    "slate"),
                ("Busy",             "Busy",             "slate"),
            ]

        for i, (label, status, color) in enumerate(buttons):
            ctk.CTkButton(
                self.btn_frame, text=label, font=FONT_BTN,
                fg_color=BTN[color][0], hover_color=BTN[color][1],
                text_color="#ffffff", height=44, corner_radius=8,
                command=lambda s=status: self._apply_status(s),
            ).grid(row=0, column=i, padx=4, pady=4, sticky="ew")
            self.btn_frame.columnconfigure(i, weight=1)

    # ═══════════════ CUSTOM MESSAGE CARD ══════════════════════

    def _build_custom_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        section_label(card, "Custom Status", "✏️")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 14))

        self.custom_entry = ctk.CTkEntry(
            row, placeholder_text="Type a custom status message…",
            font=FONT_BODY, height=40, corner_radius=8,
            fg_color=BG_INPUT, border_color=BORDER,
            text_color=FG_PRIMARY
        )
        self.custom_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            row, text="Set Status", font=FONT_BTN, width=110,
            fg_color=ACCENT, hover_color=BTN["blue"][1],
            text_color="#ffffff", height=40, corner_radius=8,
            command=self._apply_custom
        ).pack(side="right")

    def _apply_custom(self):
        txt = self.custom_entry.get().strip()
        if not txt:
            return
        ret = ""
        try:
            ret = self.return_entry.get().strip()
        except Exception:
            pass
        self._apply_status(txt, return_time=ret)
        self.custom_entry.delete(0, "end")

    # ═══════════════ TIME CARD ════════════════════════════════

    def _build_time_card(self):
        self.time_card = section_card(self.scroll)
        self.time_card.pack(fill="x", padx=24, pady=6)

        self.time_section_label = ctk.CTkLabel(
            self.time_card, text="⏱  Expected Time Availability",
            font=FONT_SECTION, text_color=FG_MUTED
        )
        self.time_section_label.pack(anchor="w", padx=16, pady=(14, 8))

        row = ctk.CTkFrame(self.time_card, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 14))

        self.return_entry = ctk.CTkEntry(
            row, placeholder_text="e.g. 2:30 PM",
            font=FONT_BODY, width=160, height=40,
            corner_radius=8, fg_color=BG_INPUT,
            border_color=BORDER, text_color=FG_PRIMARY
        )
        self.return_entry.pack(side="left", padx=(0, 8))

        self._time_buttons = []
        for delta, label in [(15, "+15m"), (30, "+30m"), (60, "+1h")]:
            btn = ctk.CTkButton(
                row, text=label, font=FONT_SMALL, width=56,
                fg_color=BTN["light"][0], hover_color=BTN["light"][1],
                text_color=FG_MUTED, height=40, corner_radius=8,
                command=lambda d=delta: self._quick_time(d)
            )
            btn.pack(side="left", padx=3)
            self._time_buttons.append(btn)

        self.update_time_btn = ctk.CTkButton(
            row, text="Update Time", font=FONT_BTN, width=130,
            fg_color=BTN["green"][0], hover_color=BTN["green"][1],
            text_color="#ffffff", height=40, corner_radius=8,
            command=self._update_time_only
        )
        self.update_time_btn.pack(side="right")

        self._toggle_time_section("inside")

    def _toggle_time_section(self, loc: str):
        # Always allow time entry now, just update label colors if needed
        self.return_entry.configure(state="normal")
        self.update_time_btn.configure(state="normal")
        for btn in self._time_buttons:
            btn.configure(state="normal")
        self.time_section_label.configure(text_color=FG_MUTED)

    def _quick_time(self, minutes: int):
        t = datetime.now() + timedelta(minutes=minutes)
        self.return_entry.delete(0, "end")
        self.return_entry.insert(0, t.strftime("%I:%M %p"))

    def _update_time_only(self):
        new_time = self.return_entry.get().strip()
        if not new_time:
            messagebox.showinfo("No Time", "Enter a return time first.")
            return
        update_return_time(new_time)
        self._refresh_current_status()

    # ═══════════════ VISITOR LOG CARD ═════════════════════════

    def _build_visitor_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            header, text="📬  Visitor Messages",
            font=FONT_SECTION, text_color=FG_MUTED
        ).pack(side="left")

        ctk.CTkButton(
            header, text="Delete", font=FONT_SMALL,
            fg_color=BTN["red"][0], hover_color=BTN["red"][1],
            text_color="#ffffff", width=80, height=30, corner_radius=6,
            command=self._delete_selected_visitor
        ).pack(side="right")

        ctk.CTkButton(
            header, text="Refresh", font=FONT_SMALL,
            fg_color=BTN["light"][0], hover_color=BTN["light"][1],
            text_color=FG_MUTED, width=80, height=30, corner_radius=6,
            command=self._refresh_visitors
        ).pack(side="right", padx=(0, 6))

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Visitor.Treeview",
                        background="#ffffff", foreground=FG_PRIMARY,
                        fieldbackground="#ffffff", rowheight=32,
                        font=("Segoe UI", 11))
        style.configure("Visitor.Treeview.Heading",
                        background=BG_INPUT, foreground=FG_MUTED,
                        font=("Segoe UI", 11, "bold"), borderwidth=0)
        style.map("Visitor.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "#ffffff")])
        style.layout("Visitor.Treeview", [
            ("Visitor.Treeview.treearea", {"sticky": "nswe"})
        ])

        tree_frame = ctk.CTkFrame(card, fg_color=BG_INPUT, corner_radius=8)
        tree_frame.pack(fill="x", padx=16, pady=(0, 14))

        columns = ("name", "purpose", "time")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Visitor.Treeview", selectmode="browse", height=6
        )
        self.tree.heading("name", text="Visitor")
        self.tree.heading("purpose", text="Message")
        self.tree.heading("time", text="Time")
        self.tree.column("name", width=130, minwidth=80)
        self.tree.column("purpose", width=300, minwidth=150)
        self.tree.column("time", width=130, minwidth=80)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical",
                                  command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.pack(side="right", fill="y", pady=8, padx=(0, 8))

        self._refresh_visitors()

    def _refresh_visitors(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for v in get_visitors():
            self.tree.insert("", "end", iid=str(v["id"]),
                             values=(v["name"], v["purpose"], v["timestamp"]))

    def _delete_selected_visitor(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No Selection", "Select a visitor message first.")
            return
        vid = int(sel[0])
        delete_visitor(vid)
        self._refresh_visitors()

    # ═══════════════ OUTLOOK SCHEDULE CARD ═════════════════════

    def _build_schedule_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            header, text="🗓  Today's Outlook Schedule",
            font=FONT_SECTION, text_color=FG_MUTED
        ).pack(side="left")

        self.schedule_status = ctk.CTkLabel(
            header, text="No events fetched", font=FONT_TINY, text_color=FG_DIM
        )
        self.schedule_status.pack(side="right")

        # Reuse styling from visitor treeview
        tree_frame = ctk.CTkFrame(card, fg_color=BG_INPUT, corner_radius=8)
        tree_frame.pack(fill="x", padx=16, pady=(0, 14))

        columns = ("sn", "title", "duration")
        self.sched_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Visitor.Treeview", selectmode="none", height=6
        )
        self.sched_tree.heading("sn", text="#")
        self.sched_tree.heading("title", text="Meeting Title")
        self.sched_tree.heading("duration", text="Duration")
        self.sched_tree.column("sn", width=50, minwidth=40, anchor="center")
        self.sched_tree.column("title", width=400, minwidth=150)
        self.sched_tree.column("duration", width=250, minwidth=150)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical",
                                  command=self.sched_tree.yview)
        self.sched_tree.configure(yscrollcommand=scrollbar.set)
        self.sched_tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.pack(side="right", fill="y", pady=8, padx=(0, 8))

        self.today_schedule_data = []

    def _refresh_schedule_ui(self):
        # Refresh the table with today_schedule_data
        for item in self.sched_tree.get_children():
            self.sched_tree.delete(item)
        if hasattr(self, 'today_schedule_data'):
            for i, event in enumerate(self.today_schedule_data, 1):
                self.sched_tree.insert("", "end", values=(i, event['title'], event['duration']))
            if self.today_schedule_data:
                self.schedule_status.configure(text=f"Updated: {datetime.now().strftime('%I:%M %p')}")

    # ═══════════════ OUTLOOK / ICS CARD ═══════════════════════

    def _build_outlook_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=(6, 18))

        # ── Header row ──
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 4))

        ctk.CTkLabel(
            header, text="📆  Outlook Calendar Sync",
            font=FONT_SECTION, text_color=FG_MUTED
        ).pack(side="left")

        self.ics_status_lbl = ctk.CTkLabel(
            header, text="Not connected", font=FONT_SMALL, text_color=FG_DIM
        )
        self.ics_status_lbl.pack(side="right")

        if not ICS_AVAILABLE:
            self.ics_status_lbl.configure(
                text="⚠ Missing: pip install icalendar recurring-ical-events requests",
                text_color="#e67700"
            )
            return

        # ── Info label ──
        ctk.CTkLabel(
            card,
            text="Paste your Outlook calendar ICS link below. "
                 "It auto-syncs every 60 seconds.",
            font=FONT_SMALL, text_color=FG_DIM, wraplength=700, justify="left"
        ).pack(anchor="w", padx=16, pady=(0, 8))

        # ── URL input row ──
        url_row = ctk.CTkFrame(card, fg_color="transparent")
        url_row.pack(fill="x", padx=16, pady=(0, 14))

        self.ics_url_entry = ctk.CTkEntry(
            url_row,
            placeholder_text="https://outlook.live.com/owa/calendar/…/calendar.ics",
            font=FONT_SMALL, height=38, corner_radius=8,
            fg_color=BG_INPUT, border_color=BORDER,
            text_color=FG_PRIMARY
        )
        self.ics_url_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            url_row, text="Connect", font=FONT_BTN, width=100,
            fg_color=ACCENT, hover_color=BTN["blue"][1],
            text_color="#ffffff", height=38, corner_radius=8,
            command=self._connect_ics
        ).pack(side="right")

        # ── Help text ──
        help_frame = ctk.CTkFrame(card, fg_color="#edf2ff", corner_radius=8)
        help_frame.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(
            help_frame,
            text="💡  How to get your ICS link:\n"
                 "1. Open outlook.com → Sign in with raptorhero@outlook.com\n"
                 "2. Click the ⚙ Settings gear → View all Outlook settings\n"
                 "3. Go to Calendar → Shared calendars\n"
                 "4. Under 'Publish a calendar', select 'Calendar' and 'Can view all details'\n"
                 "5. Click Publish → Copy the ICS link\n"
                 "6. Paste it above and click Connect",
            font=FONT_SMALL, text_color="#364fc7",
            justify="left", wraplength=720
        ).pack(padx=12, pady=10, anchor="w")

    # ═══════════════ STATUS APPLICATION ═══════════════════════

    def _apply_status(self, status_text: str, return_time: str = ""):
        if not return_time:
            try:
                return_time = self.return_entry.get().strip()
            except Exception:
                return_time = ""

        new_priority = get_priority(status_text)

        if self._ics_meeting_active:
            current = read_status()
            if current["source"] == "outlook" and new_priority < current["priority"]:
                proceed = messagebox.askyesno(
                    "Outlook Meeting Active",
                    f"An Outlook meeting is currently active.\n\n"
                    f"Override with \"{status_text}\" anyway?",
                    icon="warning"
                )
                if not proceed:
                    return

        write_status(status_text, return_time=return_time, source="manual")
        self._refresh_current_status()

        if self.location_var.get() == "outside":
            self.return_entry.delete(0, "end")

    # ═══════════════ ICS CALENDAR THREAD ══════════════════════

    def _connect_ics(self):
        url = self.ics_url_entry.get().strip()
        if not url:
            messagebox.showinfo("No URL", "Please paste your Outlook ICS link first.")
            return
        if not url.startswith("http"):
            messagebox.showerror("Invalid URL", "The link must start with http:// or https://")
            return

        # Stop existing thread if running
        self._ics_running = False
        if self._ics_thread and self._ics_thread.is_alive():
            self._ics_thread.join(timeout=2)

        _save_ics_url(url)
        self._ics_status_text = "Connecting…"
        self._start_ics_thread(url)

    def _start_ics_thread(self, url: str):
        self._ics_running = True
        self._ics_thread = threading.Thread(
            target=self._ics_worker, args=(url,), daemon=True
        )
        self._ics_thread.start()

    def _ics_worker(self, url: str):
        """Background thread: fetches and parses the ICS URL every 60s."""
        while self._ics_running:
            try:
                resp = http_requests.get(url, timeout=15)
                if resp.status_code != 200:
                    self._ics_status_text = f"⚠ HTTP {resp.status_code} — check your link"
                    time.sleep(60)
                    continue

                cal = icalendar.Calendar.from_ical(resp.content)

                # 1. Fetch current status
                now_utc = datetime.now(timezone.utc)
                start = now_utc - timedelta(minutes=1)
                end   = now_utc + timedelta(minutes=1)

                events_now = recurring_ical_events.of(cal).between(start, end)
                meeting_found = len(events_now) > 0
                meeting_subject = ""
                busy_status = 0

                for event in events_now:
                    meeting_subject = str(event.get("SUMMARY", "Untitled"))
                    transp = str(event.get("TRANSP", "OPAQUE")).upper()
                    if transp != "TRANSPARENT":
                        busy_status = max(busy_status, 2)

                # 2. Fetch full schedule for "TODAY" (local timezone)
                today_local = datetime.now().date()
                fetch_start = datetime.combine(today_local, datetime.min.time())
                fetch_end = datetime.combine(today_local, datetime.max.time())
                
                todays_events = recurring_ical_events.of(cal).between(fetch_start, fetch_end)
                
                now_local = datetime.now()
                sched_list = []
                for event in todays_events:
                    title = str(event.get("SUMMARY", "Untitled"))
                    start_dt = event.get("DTSTART").dt
                    end_dt = event.get("DTEND").dt
                    
                    # --- FILTER: Only show meetings that haven't ended yet ---
                    # Convert end_dt to a comparable naive datetime if it's local
                    cmp_end = end_dt
                    if hasattr(end_dt, "timestamp"):
                        # If it has a timezone, convert to local naive for comparison with now_local
                        # or keep both as UTC. Let's convert to local naive.
                        if end_dt.tzinfo:
                            cmp_end = end_dt.astimezone().replace(tzinfo=None)
                    
                    if isinstance(cmp_end, datetime) and cmp_end < now_local:
                        continue # Skip completed meetings
                    # ---------------------------------------------------------
                    
                    def fmt_time(dt):
                        if hasattr(dt, "strftime"):
                            return dt.strftime("%I:%M %p")
                        return str(dt)

                    duration_str = f"{fmt_time(start_dt)} - {fmt_time(end_dt)}"
                    sched_list.append({"title": title, "duration": duration_str})
                
                self.today_schedule_data = sched_list

                # 3. Apply status logic
                self._ics_meeting_active = meeting_found
                if meeting_found and busy_status > 0:
                    current = read_status()
                    target_status = "In a Meeting (Outlook)"
                    target_priority = get_priority(target_status)
                    if target_priority > current["priority"] or current["source"] == "outlook":
                        write_status(target_status, source="outlook")
                    self._ics_status_text = f"✅ Meeting: \"{meeting_subject}\""
                else:
                    self._ics_meeting_active = False
                    current = read_status()
                    if current["source"] == "outlook":
                        last_status = current.get("last_manual_status", "Available")
                        last_time = current.get("last_manual_return_time", "")
                        base_status = last_status.split(" — ")[0].strip()
                        write_status(base_status, return_time=last_time, source="manual")
                    self._ics_status_text = "✅ No active meetings"

            except Exception as e:
                self._ics_status_text = f"⚠ {e}"

            time.sleep(60)

    # ═══════════════════ PERIODIC TICK ════════════════════════

    def _tick(self):
        self._refresh_current_status()
        self._refresh_visitors()
        self._refresh_schedule_ui()
        if ICS_AVAILABLE:
            self.ics_status_lbl.configure(text=self._ics_status_text)
        self.after(5000, self._tick)

    # ═══════════════════ CLEANUP ══════════════════════════════

    def destroy(self):
        self._ics_running = False
        super().destroy()


if __name__ == "__main__":
    ControlApp().mainloop()
