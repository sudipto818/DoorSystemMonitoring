"""
display_app.py
─────────────────────────────────────────────────────
Always-on "door sign" that runs full-screen on a secondary
monitor (or falls back to a resizable window).
Reads the current status from the SQLite database every
1.5 seconds and renders it as a large, color-coded sign.
Also provides a "Leave a Message" button for visitors.

PREMIUM REDESIGN — clean typography, layered card layout,
subtle animations, and refined colour palette.
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
from db_manager import init_db, read_status, add_visitor, visitor_count


# ═══════════════════════════════════════════════════════════════════════
#  COLOUR THEMES — (bg_dark, bg_card, fg_primary, fg_muted, accent)
# ═══════════════════════════════════════════════════════════════════════

THEMES = {
    "Available":                ("#064e3b", "#065f46", "#ecfdf5", "#6ee7b7", "#10b981"),
    "Please Knock":             ("#78350f", "#92400e", "#fffbeb", "#fcd34d", "#f59e0b"),
    "In a Meeting":             ("#7f1d1d", "#991b1b", "#fef2f2", "#fca5a5", "#ef4444"),
    "In a Meeting (Outlook)":   ("#7f1d1d", "#991b1b", "#fef2f2", "#fca5a5", "#dc2626"),
    "Do Not Disturb":           ("#581c87", "#6b21a8", "#faf5ff", "#d8b4fe", "#a855f7"),
    "Do Not Disturb (Outlook)": ("#581c87", "#6b21a8", "#faf5ff", "#d8b4fe", "#9333ea"),
    "Working remotely":         ("#1e3a5f", "#1e40af", "#eff6ff", "#93c5fd", "#3b82f6"),
    "Back soon":                ("#164e63", "#155e75", "#ecfeff", "#67e8f9", "#06b6d4"),
    "Out of office":            ("#1f2937", "#374151", "#f9fafb", "#d1d5db", "#9ca3af"),
    "Busy":                     ("#7c2d12", "#9a3412", "#fff7ed", "#fdba74", "#f97316"),
}
DEFAULT_THEME = ("#0f172a", "#1e293b", "#f1f5f9", "#94a3b8", "#64748b")


def theme_for(status_text: str):
    base = status_text.split(" — ")[0].strip()
    return THEMES.get(base, DEFAULT_THEME)


# ═══════════════════════════════════════════════════════════════════════
#  DISPLAY APPLICATION
# ═══════════════════════════════════════════════════════════════════════

class DisplayApp:
    POLL_MS = 1500

    STATUS_ICONS = {
        "Available":                "✓",
        "Please Knock":             "🚪",
        "In a Meeting":             "📅",
        "In a Meeting (Outlook)":   "📅",
        "Do Not Disturb":           "⛔",
        "Do Not Disturb (Outlook)": "⛔",
        "Working remotely":         "🏠",
        "Back soon":                "⏳",
        "Out of office":            "🚶",
        "Busy":                     "🔶",
    }

    def __init__(self):
        init_db()

        self.root = tk.Tk()
        self.root.title("Door Sign — Display")
        self.root.configure(bg="#0f172a")

        # ── position on the correct monitor ──
        self._place_window()

        # ── build widgets ──
        self._build_ui()

        # ── kick off the auto-update loop ──
        self._last_status = None
        self._pulse_phase = 0
        self._poll()

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
        tk.Label(
            bottom, text="Door Sign Status  •  v1.0",
            font=("Segoe UI", 10), bg="#0f172a", fg="#334155"
        ).pack(side="right")

    # ──────────────────── Polling Loop ────────────────────────

    def _poll(self):
        try:
            row = read_status()
            status = row["current_status"]

            # Update clock
            now = datetime.now()
            self.time_label.config(
                text=now.strftime("%A, %B %d   •   %I:%M %p")
            )

            # Update theme + text only if status changed
            if status != self._last_status:
                self._last_status = status
                bg_dark, bg_card, fg_primary, fg_muted, accent = theme_for(status)

                # ── Apply colour scheme ──
                self.outer.config(bg=bg_dark)
                self.root.config(bg=bg_dark)

                # Top bar
                for w in [self.badge_label, self.time_label, self.visitor_badge]:
                    w.config(bg=bg_dark)
                self.badge_label.config(fg=fg_muted)
                self.time_label.config(fg=fg_muted)
                self.visitor_badge.config(fg=fg_muted)

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
                icon = self.STATUS_ICONS.get(base, "📌")
                self.sub_label.config(text=icon)

                # Return time / updated info
                if row["return_time"]:
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

        except Exception as exc:
            self.status_label.config(text=f"⚠ Error:\n{exc}")

        self.root.after(self.POLL_MS, self._poll)

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
