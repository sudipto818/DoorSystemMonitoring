"""Control panel UI builder mixin for Door Sign."""
import csv
import threading
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

from db_manager import (
    get_visitors, save_timetable, get_timetable, clear_timetable,
    update_return_time, read_status, delete_visitor,
    write_status, get_priority
)
from file_store import _load_display_ip, _load_ntfy_topic, _save_display_ip, _save_ntfy_topic
from optional_deps import EXCEL_AVAILABLE, ICS_AVAILABLE, openpyxl
from ui_constants import (
    ACCENT, BORDER, BG_INPUT, BTN, FG_DIM, FG_MUTED, FG_PRIMARY,
    FONT_BODY, FONT_BTN, FONT_SECTION, FONT_SMALL, FONT_TITLE,
    FONT_TINY,
)
from ui_helpers import section_card, section_label


class ControlAppUI:
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

        if hasattr(self, 'return_entry'):
            db_time = row.get("return_time", "")
            current_ui_val = self.return_entry.get().strip()
            if self._last_known_source == "outlook" and src == "manual":
                self.return_entry.delete(0, "end")
                self.return_entry.insert(0, db_time)
            elif src == "outlook":
                if db_time and current_ui_val != db_time:
                    self.return_entry.delete(0, "end")
                    self.return_entry.insert(0, db_time)

        self._last_known_source = src

        base = row["current_status"].split(" — ")[0].strip()
        dot_colors = {
            "Available": "#2b8a3e", "Please Knock": "#e67700",
            "In a Meeting": "#c92a2a", "Do Not Disturb": "#7048e8",
            "Working remotely": "#4263eb", "Back soon": "#0c8599",
            "Busy": "#d9480f", "Out of office": "#868e96",
        }
        self.status_dot.configure(text_color=dot_colors.get(base, ACCENT))

    def _build_voice_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        section_label(card, "Voice Command", "\U0001F3A4")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=(0, 14))

        self.voice_btn = ctk.CTkButton(
            inner, text="\U0001F3A4  Hold to Record", font=FONT_BTN,
            fg_color=BTN["purple"][0], hover_color=BTN["purple"][1],
            text_color="#ffffff", height=44, corner_radius=8,
            width=200,
        )
        self.voice_btn.pack(side="left", padx=(0, 12))

        self.voice_btn.bind("<ButtonPress-1>", self._on_voice_press)
        self.voice_btn.bind("<ButtonRelease-1>", self._on_voice_release)

        self.voice_status = ctk.CTkLabel(
            inner, text="Press and hold to speak a command",
            font=FONT_SMALL, text_color=FG_DIM,
            wraplength=500, justify="left",
        )
        self.voice_status.pack(side="left", fill="x", expand=True)

        if not getattr(self, '_voice_processor', None):
            self.voice_btn.configure(state="disabled", fg_color=BTN["gray"][0])
            self.voice_status.configure(
                text="Install: pip install sounddevice numpy faster-whisper",
                text_color="#e67700",
            )

    def _on_voice_press(self, event=None):
        if not getattr(self, '_is_recording', False) and getattr(self, '_recorder', None):
            try:
                self._is_recording = True
                self._recorder.start()
                self.voice_btn.configure(
                    text="\U0001F534  Recording...",
                    fg_color=BTN["red"][0],
                    hover_color=BTN["red"][1],
                )
                self.voice_status.configure(
                    text="Listening... release to process",
                    text_color=FG_PRIMARY,
                )
            except Exception as e:
                self._is_recording = False
                self.voice_status.configure(text=f"Mic error: {e}", text_color="#c92a2a")

    def _on_voice_release(self, event=None):
        if not getattr(self, '_is_recording', False):
            return
        self._is_recording = False
        filepath = self._recorder.stop()

        self.voice_btn.configure(
            text="\u23F3  Processing...",
            fg_color=BTN["amber"][0],
            hover_color=BTN["amber"][1],
        )
        self.voice_status.configure(text="Transcribing audio...", text_color=FG_MUTED)

        threading.Thread(target=self._process_voice, args=(filepath,), daemon=True).start()

    def _process_voice(self, filepath: str):
        try:
            raw_text = self._voice_processor.transcribe(filepath)
            if not raw_text.strip():
                self.after(0, self._voice_result, "No speech detected. Try again.", None)
                return

            intent = self._voice_processor.parse_command(raw_text)
            if intent["action"] == "set_status":
                status_text = intent["status"]
                ret_time = intent.get("return_time", "")
                self.after(0, self._apply_voice_status, status_text, ret_time, raw_text)

            elif intent["action"] == "create_meeting":
                meeting = intent["meeting"]
                self._create_ics_meeting(
                    title=meeting["title"],
                    duration_minutes=meeting["duration_minutes"],
                    target_date=meeting.get("target_date"),
                )
                msg = f'Meeting created: {meeting["duration_minutes"]} min'
                self.after(0, self._voice_result, msg, raw_text, "", True)

            elif intent["action"] == "set_time":
                return_time = intent.get("return_time", "")
                if return_time:
                    update_return_time(return_time)
                    msg = f'Time set: available by {return_time}'
                    row = read_status()
                    self._send_network_update(row["current_status"], return_time, row["source"])
                else:
                    msg = "Couldn't extract a time. Try again."
                self.after(0, self._voice_result, msg, raw_text, return_time)

            else:
                self.after(0, self._voice_result, "Could not understand. Try again.", raw_text)

        except Exception as e:
            self.after(0, self._voice_result, f"Error: {e}", None)

    def _apply_voice_status(self, status_text: str, return_time: str, raw_text: str):
        """Apply voice-driven status via common manual path to preserve warning prompts."""
        applied = self._apply_status(status_text, return_time=return_time)

        if applied:
            msg = f'Set: "{status_text}"'
            if return_time:
                msg += f'  (back by {return_time})'
            self._voice_result(msg, raw_text, return_time)
            return

        self._voice_result(
            "Cancelled: Outlook meeting active, status not changed.",
            raw_text,
            "",
        )

    def _voice_result(self, message: str, raw_text: str = None,
                      return_time: str = "", refresh_schedule: bool = False):
        self.voice_btn.configure(
            text="\U0001F3A4  Hold to Record",
            fg_color=BTN["purple"][0],
            hover_color=BTN["purple"][1],
        )

        status_text = message
        if raw_text:
            status_text = f'{message}  |  Heard: "{raw_text}"'

        color = "#2b8a3e" if ("Set:" in message or "Meeting" in message or "Time set:" in message) else FG_DIM
        if "Error" in message or "understand" in message or "Couldn't" in message:
            color = "#c92a2a"

        self.voice_status.configure(text=status_text, text_color=color)

        if return_time:
            self.return_entry.delete(0, "end")
            self.return_entry.insert(0, return_time)
            update_return_time(return_time)

        if refresh_schedule and self._ics_running:
            url = self.ics_url_entry.get().strip()
            if url:
                self._ics_running = False
                if self._ics_thread and self._ics_thread.is_alive():
                    self._ics_thread.join(timeout=2)
                self._start_ics_thread(url)

        self._refresh_current_status()
        self._refresh_schedule_ui()
        self.update_idletasks()

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
        row = read_status()
        self._send_network_update(row["current_status"], new_time, row["source"])
        self._refresh_current_status()

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

    def _build_schedule_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            header, text="🗓  Upcoming Schedule (Next 7 Days)",
            font=FONT_SECTION, text_color=FG_MUTED
        ).pack(side="left")

        self.schedule_status = ctk.CTkLabel(
            header, text="No events fetched", font=FONT_TINY, text_color=FG_DIM
        )
        self.schedule_status.pack(side="right")

        tree_frame = ctk.CTkFrame(card, fg_color=BG_INPUT, corner_radius=8)
        tree_frame.pack(fill="x", padx=16, pady=(0, 14))

        columns = ("sn", "date", "title", "duration")
        self.sched_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Visitor.Treeview", selectmode="none", height=6
        )
        self.sched_tree.heading("sn", text="#")
        self.sched_tree.heading("date", text="Date")
        self.sched_tree.heading("title", text="Meeting Title")
        self.sched_tree.heading("duration", text="Duration")
        self.sched_tree.column("sn", width=50, minwidth=40, anchor="center")
        self.sched_tree.column("date", width=120, minwidth=100)
        self.sched_tree.column("title", width=330, minwidth=150)
        self.sched_tree.column("duration", width=250, minwidth=150)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical",
                                  command=self.sched_tree.yview)
        self.sched_tree.configure(yscrollcommand=scrollbar.set)
        self.sched_tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.pack(side="right", fill="y", pady=8, padx=(0, 8))

        self.upcoming_schedule_data = []

    def _refresh_schedule_ui(self):
        for item in self.sched_tree.get_children():
            self.sched_tree.delete(item)
        if hasattr(self, 'upcoming_schedule_data'):
            for i, event in enumerate(self.upcoming_schedule_data, 1):
                self.sched_tree.insert("", "end", values=(i, event['date'], event['title'], event['duration']))
            if self.upcoming_schedule_data:
                self.schedule_status.configure(text=f"Updated: {datetime.now().strftime('%I:%M %p')}")

    def _build_timetable_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            header, text="📅  Weekly Weekday Timetable",
            font=FONT_SECTION, text_color=FG_MUTED
        ).pack(side="left")

        self.timetable_status = ctk.CTkLabel(
            header, text="No timetable uploaded", font=FONT_TINY, text_color=FG_DIM
        )
        self.timetable_status.pack(side="right")

        filter_row = ctk.CTkFrame(card, fg_color="transparent")
        filter_row.pack(fill="x", padx=16, pady=(0, 10))

        self.day_filter_var = ctk.StringVar(value="ALL")
        days = ["ALL", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]

        self.day_filter = ctk.CTkSegmentedButton(
            filter_row, values=days,
            command=lambda v: self._refresh_timetable_ui(),
            variable=self.day_filter_var,
            font=FONT_TINY, height=28,
            selected_color=ACCENT, unselected_color=BG_INPUT,
            text_color=FG_PRIMARY, selected_hover_color=BTN["blue"][1]
        )
        self.day_filter.pack(side="left")

        tree_frame = ctk.CTkFrame(card, fg_color=BG_INPUT, corner_radius=8)
        tree_frame.pack(fill="x", padx=16, pady=(0, 14))

        columns = ("day", "name", "start", "end")
        self.time_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Visitor.Treeview", selectmode="none", height=6
        )
        self.time_tree.heading("day", text="Day")
        self.time_tree.heading("name", text="Event Name")
        self.time_tree.heading("start", text="Start")
        self.time_tree.heading("end", text="End")
        self.time_tree.column("day", width=120, minwidth=100)
        self.time_tree.column("name", width=280, minwidth=150)
        self.time_tree.column("start", width=100, minwidth=80, anchor="center")
        self.time_tree.column("end", width=100, minwidth=80, anchor="center")

        self.time_tree.pack(fill="both", expand=True, padx=8, pady=8)

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkButton(
            btn_row, text="Upload Weekly Matrix (CSV/XLSX)", font=FONT_BTN,
            fg_color=ACCENT, hover_color=BTN["blue"][1],
            text_color="#ffffff", height=34, corner_radius=8,
            command=self._upload_timetable
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row, text="Clear", font=FONT_BTN, width=80,
            fg_color=BTN["light"][0], hover_color=BTN["light"][1],
            text_color=FG_MUTED, height=34, corner_radius=8,
            command=self._clear_timetable
        ).pack(side="left")

        self._refresh_timetable_ui()

    def _refresh_timetable_ui(self):
        for item in self.time_tree.get_children():
            self.time_tree.delete(item)

        filter_day = self.day_filter_var.get()
        data = get_timetable(day=None if filter_day == "ALL" else filter_day)
        day_order = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY", "ALL"]
        data.sort(key=lambda x: (day_order.index(x["day"]) if x["day"] in day_order else 99, x["start"]))

        for item in data:
            self.time_tree.insert("", "end", values=(item["day"], item["name"], item["start"], item["end"]))
        else:
            self.timetable_status.configure(text="No timetable uploaded")

    def _upload_timetable(self):
        file_path = filedialog.askopenfilename(
            title="Upload Weekly Matrix Timetable",
            filetypes=[("Matrix files", "*.csv *.xlsx"), ("CSV", "*.csv"), ("Excel", "*.xlsx")]
        )
        if not file_path:
            return

        new_data = []
        try:
            if file_path.endswith(".xlsx"):
                if not EXCEL_AVAILABLE:
                    raise ImportError("Missing 'openpyxl' library. Run: pip install openpyxl")
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                rows = list(ws.values)
                if not rows:
                    raise ValueError("Excel file is empty.")

                header = [str(cell).strip().upper() if cell else "" for cell in rows[0]]
                data_rows = rows[1:]
                for r in data_rows:
                    if not any(r):
                        continue
                    time_val = str(r[0] or "").strip()
                    self._parse_matrix_row(time_val, r, header, new_data)

            elif file_path.endswith(".csv"):
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                    if not rows:
                        raise ValueError("CSV file is empty.")

                    header = [str(c).strip().upper() for c in rows[0]]
                    data_rows = rows[1:]
                    for r in data_rows:
                        if not any(r):
                            continue
                        time_val = str(r[0] or "").strip()
                        self._parse_matrix_row(time_val, r, header, new_data)
            else:
                raise ValueError("Unsupported file format.")

            if not new_data:
                raise ValueError("No valid events found. Ensure first column is 'Time' (HH:MM - HH:MM) and columns match days.")

            save_timetable(new_data)
            self._refresh_timetable_ui()
            messagebox.showinfo("Success", f"Weekly schedule loaded with {len(new_data)} events.")
        except Exception as e:
            messagebox.showerror("Upload Error", f"Failed to parse file:\n{e}")

    def _parse_matrix_row(self, time_str, row_values, header, out_list):
        if not time_str or " - " not in time_str:
            return
        try:
            start_t, end_t = time_str.split(" - ")
            start_t = start_t.strip()
            end_t = end_t.strip()
        except Exception:
            return

        days = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
        for i, val in enumerate(row_values):
            if i == 0:
                continue
            if i >= len(header):
                break
            col_name = header[i]
            if col_name in days and val and str(val).strip():
                out_list.append({
                    "day": col_name,
                    "name": str(val).strip(),
                    "start": start_t,
                    "end": end_t,
                })

    def _clear_timetable(self):
        if messagebox.askyesno("Clear", "Remove the daily timetable?"):
            clear_timetable()
            self._refresh_timetable_ui()

    def _build_outlook_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=(6, 18))

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

        ctk.CTkLabel(
            card,
            text="Paste your Outlook calendar ICS link below. "
                 "It auto-syncs every 60 seconds.",
            font=FONT_SMALL, text_color=FG_DIM, wraplength=700, justify="left"
        ).pack(anchor="w", padx=16, pady=(0, 8))

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

    def _build_network_card(self):
        card = section_card(self.scroll)
        card.pack(fill="x", padx=24, pady=6)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            header, text="🌐  Network Display",
            font=FONT_SECTION, text_color=FG_MUTED
        ).pack(side="left")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 14))

        self.display_ip_entry = ctk.CTkEntry(
            row, placeholder_text="e.g. 192.168.1.105",
            font=FONT_BODY, width=160, height=40,
            corner_radius=8, fg_color=BG_INPUT,
            border_color=BORDER, text_color=FG_PRIMARY
        )
        self.display_ip_entry.pack(side="left", padx=(0, 8))

        saved_ip = _load_display_ip()
        if saved_ip:
            self.display_ip_entry.insert(0, saved_ip)

        ctk.CTkButton(
            row, text="Connect", font=FONT_BTN, width=130,
            fg_color=BTN["blue"][0], hover_color=BTN["blue"][1],
            text_color="#ffffff", height=40, corner_radius=8,
            command=self._test_network_connection
        ).pack(side="left")

        self.net_status_lbl = ctk.CTkLabel(
            row, text="", font=FONT_SMALL, text_color=FG_DIM
        )
        self.net_status_lbl.pack(side="left", padx=(16, 0))

        ntfy_row = ctk.CTkFrame(card, fg_color="transparent")
        ntfy_row.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(
            ntfy_row, text="📱  ntfy.sh topic code:",
            font=FONT_SMALL, text_color=FG_MUTED
        ).pack(side="left")

        self.ntfy_topic_entry = ctk.CTkEntry(
            ntfy_row, placeholder_text="e.g. my_door_123",
            font=FONT_SMALL, width=220, height=36,
            corner_radius=8, fg_color=BG_INPUT,
            border_color=BORDER, text_color=FG_PRIMARY
        )
        self.ntfy_topic_entry.pack(side="left", padx=(8, 8))

        saved_topic = _load_ntfy_topic()
        if saved_topic:
            self.ntfy_topic_entry.insert(0, saved_topic)

        ctk.CTkButton(
            ntfy_row, text="Save", font=FONT_SMALL, width=60,
            fg_color=BTN["green"][0], hover_color=BTN["green"][1],
            text_color="#ffffff", height=36, corner_radius=8,
            command=self._save_ntfy_topic
        ).pack(side="left")

        self.ntfy_status_lbl = ctk.CTkLabel(
            ntfy_row, text="", font=FONT_SMALL, text_color=FG_DIM
        )
        self.ntfy_status_lbl.pack(side="left", padx=(8, 0))

    def _test_network_connection(self):
        ip = self.display_ip_entry.get().strip()
        if not ip:
            messagebox.showinfo("IP Missing", "Enter the Display PC's IP address first.")
            return

        _save_display_ip(ip)
        self.net_status_lbl.configure(text="Connecting...", text_color=FG_MUTED)
        self.update_idletasks()

        row = read_status()
        self._send_network_update(row["current_status"], row.get("return_time", ""), row["source"])

    def _save_ntfy_topic(self):
        topic = self.ntfy_topic_entry.get().strip()
        if topic:
            topic = ''.join(e for e in topic if e.isalnum() or e in '_-')
            _save_ntfy_topic(topic)
            self.ntfy_topic_entry.delete(0, 'end')
            self.ntfy_topic_entry.insert(0, topic)
            self.ntfy_status_lbl.configure(text="✓ Saved", text_color="#10b981")
        else:
            self.ntfy_status_lbl.configure(text="Enter a code", text_color="#ef4444")

    def _build_footer(self):
        footer = ctk.CTkFrame(self.scroll, fg_color="transparent")
        footer.pack(fill="x", padx=24, pady=(0, 18))

        ctk.CTkLabel(
            footer,
            text="Credits: Sudipto Ghosh, Kushagra Singhal",
            font=FONT_TINY,
            text_color=FG_DIM,
        ).pack(anchor="e")
