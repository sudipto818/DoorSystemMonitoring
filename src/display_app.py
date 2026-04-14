"""
display_app.py
─────────────────────────────────────────────────────
Always-on "door sign" that runs full-screen on a secondary
monitor (or falls back to a resizable window).
Reads the current status from the SQLite database every
1.5 seconds and renders it as a large, color-coded sign.
Also provides a "Leave a Message" button for visitors.

"""

import tkinter as tk
from tkinter import font as tkfont
import sys
import math
import time
from datetime import datetime

# ── third-party ──
try:
    from screeninfo import get_monitors
except ImportError:
    get_monitors = None  # graceful fallback

# ── project ──
from db_manager import init_db, read_status, write_status, add_visitor, visitor_count, get_timetable
import queue
try:
    from network_bridge import StatusServer, send_visitor_message
except ImportError:
    StatusServer = None
    send_visitor_message = None

from ui_constants import STATUS_ICONS, theme_for, blend


# ═══════════════════════════════════════════════════════════════════════
#  DISPLAY APPLICATION
# ═══════════════════════════════════════════════════════════════════════

class DisplayApp:
    POLL_MS = 1500

    def __init__(self):
        init_db()

        self.root = tk.Tk()
        self.root.title("Door Sign — Display")
        self.root.configure(bg="#0f172a")

        # ── position on the correct monitor ──
        self._place_window()

        # ── build widgets ──
        self._build_ui()

        # ── network server ──
        self.network_queue = queue.Queue()
        if StatusServer:
            self.server = StatusServer(callback=self._on_network_status)
            self.server.start()
            self.local_ip = self.server.get_local_ip()
            self.branding_label.config(text=f"Door Sign Status  •  v1.0  •  IP: {self.local_ip}")
        else:
            self.server = None
            self.local_ip = "Unknown"

        # ── state & loop ──
        self._active_event = None
        self._last_status = None
        self._pulse_phase = 0
        self._poll()
        self._check_network_queue()

    # ─────────────────── Monitor Detection ────────────────────

    def _place_window(self):
        if get_monitors is not None:
            monitors = get_monitors()
            if len(monitors) >= 2:
                m = monitors[1]
                self.root.geometry(f"{m.width}x{m.height}+{m.x}+{m.y}")
                self.root.overrideredirect(True)
                self.root.attributes("-topmost", True)
                return

        self.root.geometry("1200x750+80+60")
        self.root.minsize(900, 550)

    # ─────────────────────── UI Build ─────────────────────────

    def _build_ui(self):
        # ── Outer container ──
        self.outer = tk.Frame(self.root, bg="#0f172a")
        self.outer.pack(fill="both", expand=True)

        # ── Top bar — clock + badge ──
        top_bar = tk.Frame(self.outer, bg="#0f172a")
        top_bar.pack(fill="x", padx=40, pady=(20, 0))

        self.badge_label = tk.Label(
            top_bar, text="DOOR STATUS", font=("Segoe UI", 11, "bold"),
            bg="#0f172a", fg="#475569",
            padx=12, pady=4
        )
        self.badge_label.pack(side="left")

        self.time_label = tk.Label(
            top_bar, text="", font=("Segoe UI", 14),
            bg="#0f172a", fg="#64748b", anchor="e"
        )
        self.time_label.pack(side="right")

        self.visitor_badge = tk.Label(
            top_bar, text="", font=("Segoe UI", 12, "bold"),
            bg="#0f172a", fg="#94a3b8"
        )
        self.visitor_badge.pack(side="right", padx=20)

        # ── Main card ──
        self.card_outer = tk.Frame(self.outer, bg="#1e293b", bd=0)
        self.card_outer.pack(fill="both", expand=True, padx=40, pady=20)

        # Inner card with padding
        self.card = tk.Frame(self.card_outer, bg="#1e293b")
        self.card.pack(fill="both", expand=True, padx=3, pady=3)

        # ── Accent bar at top of card ──
        self.accent_bar = tk.Frame(self.card, bg="#64748b", height=5)
        self.accent_bar.pack(fill="x", side="top")

        # ── Content area ──
        content = tk.Frame(self.card, bg="#1e293b")
        content.pack(fill="both", expand=True)

        # Status icon (circle indicator)
        self.icon_frame = tk.Frame(content, bg="#1e293b")
        self.icon_frame.pack(pady=(50, 10))

        self.icon_canvas = tk.Canvas(
            self.icon_frame, width=80, height=80,
            bg="#1e293b", highlightthickness=0
        )
        self.icon_canvas.pack()
        self._dot = self.icon_canvas.create_oval(15, 15, 65, 65, fill="#64748b", outline="")

        # Status text — large and bold
        self.status_label = tk.Label(
            content, text="Loading…",
            font=("Segoe UI", 58, "bold"),
            bg="#1e293b", fg="#f1f5f9",
            wraplength=950, justify="center"
        )
        self.status_label.pack(expand=True, padx=50)

        # Subtitle
        self.sub_label = tk.Label(
            content, text="",
            font=("Segoe UI", 18),
            bg="#1e293b", fg="#94a3b8",
        )
        self.sub_label.pack(pady=(0, 10))

        # Divider
        self.divider = tk.Frame(content, bg="#334155", height=1)
        self.divider.pack(fill="x", padx=80, pady=(10, 20))

        # Time info row
        self.info_label = tk.Label(
            content, text="",
            font=("Segoe UI", 15),
            bg="#1e293b", fg="#64748b",
        )
        self.info_label.pack(pady=(0, 40))

        # ── Bottom bar ──
        bottom = tk.Frame(self.outer, bg="#0f172a")
        bottom.pack(fill="x", padx=40, pady=(0, 24))

        # Leave a message button — pill style
        self.msg_btn = tk.Button(
            bottom, text="✉   Leave a Message",
            font=("Segoe UI", 15, "bold"),
            bg="#1e293b", fg="#94a3b8",
            activebackground="#334155", activeforeground="#e2e8f0",
            relief="flat", padx=30, pady=10, cursor="hand2",
            bd=0, highlightthickness=0,
            command=self._open_visitor_popup
        )
        self.msg_btn.pack(side="left")

        # Branding
        self.branding_label = tk.Label(
            bottom, text="Door Sign Status  •  v1.0",
            font=("Segoe UI", 10, "bold"), bg="#0f172a", fg="#94a3b8"
        )
        self.branding_label.pack(side="right")

    # ─────────────────── Network Updates ────────────────────────

    def _on_network_status(self, status, return_time, source):
        """Called by network_bridge from a background thread."""
        # Persist to local DB so the polling loop stays in sync
        write_status(status, return_time=return_time, source=source)

        display_status = f"{status} — back by {return_time}" if return_time else status
        self.network_queue.put({
            "current_status": display_status,
            "return_time": return_time,
            "source": source,
            "last_updated": datetime.now().strftime("%I:%M %p")
        })

    def _check_network_queue(self):
        """Periodically check for queued network updates."""
        try:
            latest = None
            while True:
                latest = self.network_queue.get_nowait()
        except queue.Empty:
            pass

        if latest:
            self._apply_row(latest)
            
        self.root.after(500, self._check_network_queue)

    # ──────────────────── Polling Loop ────────────────────────

    def _poll(self):
        try:
            # 1. Check for Timetable Events first (Additive Layer)
            active_event = self._get_active_timetable_event()
            
            # 2. Get the underlying manual/outlook status
            row = read_status()
            
            # 3. Apply overrides if a timetable event is active
            if active_event:
                event_name = active_event["name"]
                # Map theme based on keywords
                lower_name = event_name.lower()
                effective_status = "In a Meeting" # Default
                
                if "class" in lower_name:
                    effective_status = "Do Not Disturb"
                    status_display = f"In a Class: {event_name}"
                elif "meeting" in lower_name:
                    effective_status = "In a Meeting"
                    status_display = event_name
                elif "available" in lower_name or "free" in lower_name:
                    effective_status = "Available"
                    status_display = "Available"
                else:
                    status_display = event_name
                
                # Create a synthetic row for the UI
                timetable_row = {
                    "current_status": status_display,
                    "return_time": active_event["end"],
                    "source": "timetable",
                    "theme_status": effective_status # used for theme lookup
                }
                self._apply_row(timetable_row)
            else:
                self._apply_row(row)
                
        except Exception as exc:
            self.status_label.config(text=f"⚠ Error:\n{exc}")

        self.root.after(self.POLL_MS, self._poll)

    def _get_active_timetable_event(self):
        """Find the most recently started active event in the timetable for today."""
        now_dt = datetime.now()
        now_time = now_dt.strftime("%H:%M")
        day_of_week = now_dt.strftime("%A").upper() # MONDAY, TUESDAY, etc.
        
        events = get_timetable(day=day_of_week)
        active_events = []
        for e in events:
            # Simple HH:MM comparison
            if e["start"] <= now_time < e["end"]:
                active_events.append(e)
        
        if not active_events:
            return None
            
        # Sort by start time (descending) to get the most recent one
        active_events.sort(key=lambda x: x["start"], reverse=True)
        return active_events[0]

    def _apply_row(self, row):
        status = row["current_status"]
        theme_ref = row.get("theme_status", status) # use mapped status or actual status

        # Update clock
        now = datetime.now()
        self.time_label.config(
            text=now.strftime("%A, %B %d   •   %I:%M %p")
        )

        # Update theme + text only if status or event changed
        if status != self._last_status:
            self._last_status = status
            bg_dark, bg_card, fg_primary, fg_muted, accent = theme_for(theme_ref)

            # ── Apply colour scheme ──
            self.outer.config(bg=bg_dark)
            self.root.config(bg=bg_dark)

            # Top bar and Bottom bar text
            for w in [self.badge_label, self.time_label, self.visitor_badge]:
                w.config(bg=bg_dark)
            self.badge_label.config(fg=fg_muted)
            self.time_label.config(fg=fg_muted)
            self.visitor_badge.config(fg=fg_muted)
            self.branding_label.config(fg=fg_muted)

            # Card
            self.card_outer.config(bg=self._blend(bg_card, "#000000", 0.3))
            self.card.config(bg=bg_card)
            self.accent_bar.config(bg=accent)

            # Content
            for w in self.card.winfo_children():
                if isinstance(w, tk.Frame) and w is not self.accent_bar:
                    w.config(bg=bg_card)
                    for child in w.winfo_children():
                        if isinstance(child, (tk.Label, tk.Frame)):
                            child.config(bg=bg_card)
                        if isinstance(child, tk.Canvas):
                            child.config(bg=bg_card)

            self.icon_frame.config(bg=bg_card)
            self.icon_canvas.config(bg=bg_card)
            self.icon_canvas.itemconfig(self._dot, fill=accent)

            # Base status (without return time suffix)
            base = status.split(" — ")[0].strip()

            self.status_label.config(bg=bg_card, fg=fg_primary, text=base)
            self.sub_label.config(bg=bg_card, fg=fg_muted)
            self.divider.config(bg=self._blend(fg_muted, bg_card, 0.6))
            self.info_label.config(bg=bg_card, fg=fg_muted)

            # Icon
            icon = STATUS_ICONS.get(base, "📌")
            self.sub_label.config(text=icon)

            # Return time / updated info
            if row.get("return_time"):
                self.info_label.config(
                    text=f"Expected Time Availability  •  {row['return_time']}"
                )
            else:
                updated = row.get("last_updated", "")
                self.info_label.config(text=f"Last updated  •  {updated}")

            # Bottom bar
            for w in self.outer.winfo_children():
                if isinstance(w, tk.Frame) and w is not self.card_outer:
                    w.config(bg=bg_dark)
                    for child in w.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg=bg_dark)
                        if isinstance(child, tk.Button):
                            child.config(bg=self._blend(bg_card, bg_dark, 0.5))

        # Visitor badge update
        vc = visitor_count()
        self.visitor_badge.config(
            text=f"📬 {vc} message{'s' if vc != 1 else ''}" if vc else ""
        )

    # ──────────────── Visitor Popup ───────────────────────────

    def _open_visitor_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("Leave a Message")
        popup.configure(bg="#0f172a")
        popup.geometry("540x480")
        popup.resizable(False, False)
        popup.grab_set()

        # Center
        popup.update_idletasks()
        pw, ph = 540, 480
        sx = (self.root.winfo_screenwidth() - pw) // 2
        sy = (self.root.winfo_screenheight() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        # Card
        card = tk.Frame(popup, bg="#1e293b")
        card.pack(fill="both", expand=True, padx=16, pady=16)

        # Accent bar
        tk.Frame(card, bg="#3b82f6", height=4).pack(fill="x")

        # Header
        tk.Label(
            card, text="Leave a Message",
            font=("Segoe UI", 22, "bold"),
            bg="#1e293b", fg="#f1f5f9"
        ).pack(pady=(28, 4))

        tk.Label(
            card, text="The occupant will see your message when they return.",
            font=("Segoe UI", 11),
            bg="#1e293b", fg="#64748b"
        ).pack(pady=(0, 20))

        # Name field
        tk.Label(card, text="YOUR NAME", font=("Segoe UI", 10, "bold"),
                 bg="#1e293b", fg="#94a3b8").pack(anchor="w", padx=36)
        name_entry = tk.Entry(card, font=("Segoe UI", 14), bg="#0f172a",
                              fg="#f1f5f9", insertbackground="#f1f5f9",
                              relief="flat", bd=0)
        name_entry.pack(fill="x", padx=36, pady=(6, 16), ipady=10)

        # Purpose field
        tk.Label(card, text="MESSAGE", font=("Segoe UI", 10, "bold"),
                 bg="#1e293b", fg="#94a3b8").pack(anchor="w", padx=36)
        purpose_text = tk.Text(card, font=("Segoe UI", 13), bg="#0f172a",
                               fg="#f1f5f9", insertbackground="#f1f5f9",
                               relief="flat", height=4, wrap="word", bd=0)
        purpose_text.pack(fill="x", padx=36, pady=(6, 16))

        feedback_lbl = tk.Label(card, text="", font=("Segoe UI", 11),
                                bg="#1e293b", fg="#ef4444")
        feedback_lbl.pack()

        def _submit():
            n = name_entry.get().strip()
            p = purpose_text.get("1.0", "end").strip()
            if not n or not p:
                feedback_lbl.config(text="Please fill in both fields.", fg="#ef4444")
                return
            add_visitor(n, p)

            # Forward visitor message to control PC over network
            if send_visitor_message and self.server and self.server.last_client_ip:
                import threading
                control_ip = self.server.last_client_ip
                threading.Thread(
                    target=send_visitor_message,
                    args=(control_ip, n, p),
                    daemon=True
                ).start()

            feedback_lbl.config(text="✓  Message saved — thank you!", fg="#10b981")
            name_entry.delete(0, "end")
            purpose_text.delete("1.0", "end")
            popup.after(1800, popup.destroy)

        btn_frame = tk.Frame(card, bg="#1e293b")
        btn_frame.pack(pady=(4, 20))

        tk.Button(
            btn_frame, text="Submit Message",
            font=("Segoe UI", 14, "bold"),
            bg="#3b82f6", fg="#ffffff",
            activebackground="#2563eb", activeforeground="#ffffff",
            relief="flat", padx=32, pady=10, cursor="hand2", bd=0,
            command=_submit
        ).pack(side="left", padx=4)

        tk.Button(
            btn_frame, text="Cancel",
            font=("Segoe UI", 14),
            bg="#334155", fg="#94a3b8",
            activebackground="#475569", activeforeground="#e2e8f0",
            relief="flat", padx=24, pady=10, cursor="hand2", bd=0,
            command=popup.destroy
        ).pack(side="left", padx=4)

    # ──────────────── Colour Utilities ────────────────────────

    @staticmethod
    def _blend(hex1: str, hex2: str, factor: float = 0.5) -> str:
        """Blend two hex colours. factor=0 → hex1, factor=1 → hex2."""
        h1, h2 = hex1.lstrip("#"), hex2.lstrip("#")
        r1, g1, b1 = (int(h1[i:i+2], 16) for i in (0, 2, 4))
        r2, g2, b2 = (int(h2[i:i+2], 16) for i in (0, 2, 4))
        r = int(r1 + (r2 - r1) * factor)
        g = int(g1 + (g2 - g1) * factor)
        b = int(b1 + (b2 - b1) * factor)
        return f"#{r:02x}{g:02x}{b:02x}"

    # ──────────────────── Run ─────────────────────────────────

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    DisplayApp().run()
