"""Display UI builder mixin for the Door Sign display."""
import tkinter as tk
import threading
import queue
from datetime import datetime

try:
    from screeninfo import get_monitors
except ImportError:
    get_monitors = None

try:
    from network_bridge import send_visitor_message
except ImportError:
    send_visitor_message = None

from db_manager import add_visitor, write_status, visitor_count
from ui_constants import STATUS_ICONS, theme_for, blend


class DisplayAppUI:
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

    def _build_ui(self):
        self.outer = tk.Frame(self.root, bg="#0f172a")
        self.outer.pack(fill="both", expand=True)

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

        self.card_outer = tk.Frame(self.outer, bg="#1e293b", bd=0)
        self.card_outer.pack(fill="both", expand=True, padx=40, pady=20)

        self.card = tk.Frame(self.card_outer, bg="#1e293b")
        self.card.pack(fill="both", expand=True, padx=3, pady=3)

        self.accent_bar = tk.Frame(self.card, bg="#64748b", height=5)
        self.accent_bar.pack(fill="x", side="top")

        content = tk.Frame(self.card, bg="#1e293b")
        content.pack(fill="both", expand=True)

        self.icon_frame = tk.Frame(content, bg="#1e293b")
        self.icon_frame.pack(pady=(50, 10))

        self.icon_canvas = tk.Canvas(
            self.icon_frame, width=80, height=80,
            bg="#1e293b", highlightthickness=0
        )
        self.icon_canvas.pack()
        self._dot = self.icon_canvas.create_oval(15, 15, 65, 65, fill="#64748b", outline="")

        self.status_label = tk.Label(
            content, text="Loading…",
            font=("Segoe UI", 58, "bold"),
            bg="#1e293b", fg="#f1f5f9",
            wraplength=950, justify="center"
        )
        self.status_label.pack(expand=True, padx=50)

        self.sub_label = tk.Label(
            content, text="",
            font=("Segoe UI", 18),
            bg="#1e293b", fg="#94a3b8",
        )
        self.sub_label.pack(pady=(0, 10))

        self.divider = tk.Frame(content, bg="#334155", height=1)
        self.divider.pack(fill="x", padx=80, pady=(10, 20))

        self.info_label = tk.Label(
            content, text="",
            font=("Segoe UI", 15),
            bg="#1e293b", fg="#64748b",
        )
        self.info_label.pack(pady=(0, 40))

        bottom = tk.Frame(self.outer, bg="#0f172a")
        bottom.pack(fill="x", padx=40, pady=(0, 24))

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

        self.branding_label = tk.Label(
            bottom, text="Door Sign Status  •  v1.0",
            font=("Segoe UI", 10, "bold"), bg="#0f172a", fg="#94a3b8"
        )
        self.branding_label.pack(side="right")

    def _on_network_status(self, status, return_time, source):
        write_status(status, return_time=return_time, source=source)

        display_status = f"{status} — back by {return_time}" if return_time else status
        self.network_queue.put({
            "current_status": display_status,
            "return_time": return_time,
            "source": source,
            "last_updated": datetime.now().strftime("%I:%M %p")
        })

    def _check_network_queue(self):
        latest = None
        try:
            while True:
                latest = self.network_queue.get_nowait()
        except queue.Empty:
            pass

        if latest:
            self._apply_row(latest)
        self.root.after(500, self._check_network_queue)

    def _apply_row(self, row):
        status = row["current_status"]
        theme_ref = row.get("theme_status", status)

        now = datetime.now()
        self.time_label.config(text=now.strftime("%A, %B %d   •   %I:%M %p"))

        if status != self._last_status:
            self._last_status = status
            bg_dark, bg_card, fg_primary, fg_muted, accent = theme_for(theme_ref)
            self.outer.config(bg=bg_dark)
            self.root.config(bg=bg_dark)

            for w in [self.badge_label, self.time_label, self.visitor_badge]:
                w.config(bg=bg_dark)
            self.badge_label.config(fg=fg_muted)
            self.time_label.config(fg=fg_muted)
            self.visitor_badge.config(fg=fg_muted)
            self.branding_label.config(fg=fg_muted)

            self.card_outer.config(bg=blend(bg_card, "#000000", 0.3))
            self.card.config(bg=bg_card)
            self.accent_bar.config(bg=accent)

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

            base = status.split(" — ")[0].strip()
            self.status_label.config(bg=bg_card, fg=fg_primary, text=base)
            self.sub_label.config(bg=bg_card, fg=fg_muted)
            self.divider.config(bg=blend(fg_muted, bg_card, 0.6))
            self.info_label.config(bg=bg_card, fg=fg_muted)

            icon = STATUS_ICONS.get(base, "📌")
            self.sub_label.config(text=icon)

            if row.get("return_time"):
                self.info_label.config(text=f"Expected Time Availability  •  {row['return_time']}")
            else:
                updated = row.get("last_updated", "")
                self.info_label.config(text=f"Last updated  •  {updated}")

            for w in self.outer.winfo_children():
                if isinstance(w, tk.Frame) and w is not self.card_outer:
                    w.config(bg=bg_dark)

        vc = visitor_count()
        self.visitor_badge.config(
            text=f"📬 {vc} message{'s' if vc != 1 else ''}" if vc else ""
        )

    def _open_visitor_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("Leave a Message")
        popup.configure(bg="#0f172a")
        popup.geometry("540x480")
        popup.resizable(False, False)
        popup.grab_set()

        popup.update_idletasks()
        pw, ph = 540, 480
        sx = (self.root.winfo_screenwidth() - pw) // 2
        sy = (self.root.winfo_screenheight() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        card = tk.Frame(popup, bg="#1e293b")
        card.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Frame(card, bg="#3b82f6", height=4).pack(fill="x")

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

        tk.Label(card, text="YOUR NAME", font=("Segoe UI", 10, "bold"),
                 bg="#1e293b", fg="#94a3b8").pack(anchor="w", padx=36)
        name_entry = tk.Entry(card, font=("Segoe UI", 14), bg="#0f172a",
                              fg="#f1f5f9", insertbackground="#f1f5f9",
                              relief="flat", bd=0)
        name_entry.pack(fill="x", padx=36, pady=(6, 16), ipady=10)

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

            if send_visitor_message and getattr(self, 'server', None) and getattr(self.server, 'last_client_ip', None):
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
