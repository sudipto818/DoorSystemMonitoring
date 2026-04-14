"""
control_app.py
─────────────────────────────────────────────────────
Owner's control panel for the Door-Sign-Status system.
Can be opened / closed freely — the display_app.py
keeps running and reads from the shared SQLite database.

"""

import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import sys
from datetime import datetime, timedelta, timezone

from db_manager import (
    init_db, read_status, write_status, update_return_time,
    get_visitors, delete_visitor,
    get_priority, STATUS_PRIORITY,
    save_timetable, get_timetable, clear_timetable
)
import json
import csv

from file_store import (
    _load_display_ip, _load_ics_url,
    _load_ntfy_topic, _save_ics_url,
    _save_ntfy_topic,
)
from optional_deps import ICS_AVAILABLE, http_requests, recurring_ical_events, icalendar
from ui_constants import (
    ACCENT, BORDER, BG_BASE, BG_CARD, BG_INPUT,
    BTN, FONT_BODY, FONT_BTN, FONT_SECTION,
    FONT_SMALL, FONT_TITLE, FONT_TINY,
    FG_DIM, FG_MUTED, FG_PRIMARY,
)
from ui_helpers import section_card, section_label
from control_ui import ControlAppUI

try:
    from voice_command import AudioRecorder, VoiceProcessor, create_ics_meeting, AUDIO_AVAILABLE, WHISPER_AVAILABLE
    VOICE_AVAILABLE = AUDIO_AVAILABLE and WHISPER_AVAILABLE
except ImportError:
    VOICE_AVAILABLE = False

try:
    from network_bridge import send_status_update, StatusServer
except ImportError:
    send_status_update = None
    StatusServer = None

# ═══════════════════════════════════════════════════════════════════════
#  APPEARANCE 
# ═══════════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

#  CONTROL PANEL APPLICATION
# ═══════════════════════════════════════════════════════════════════════

class ControlApp(ctk.CTk, ControlAppUI):

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
        self._last_known_source = "manual"  # Track source changes to manage UI synchronization

        # Voice command state
        self._recorder = AudioRecorder() if VOICE_AVAILABLE else None
        self._voice_processor = VoiceProcessor() if VOICE_AVAILABLE else None
        self._is_recording = False

        # Scrollable container
        self.scroll = ctk.CTkScrollableFrame(
            self, fg_color=BG_BASE, corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=FG_DIM
        )
        self.scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # Network listener for incoming visitor messages
        self._visitor_queue = __import__('queue').Queue()
        if StatusServer:
            self._control_server = StatusServer(
                visitor_callback=self._on_visitor_received
            )
            self._control_server.start()
        else:
            self._control_server = None

        # Build all sections
        self._build_header()
        self._build_status_bar()
        self._build_voice_card()
        self._build_location_card()
        self._build_status_card()
        self._build_custom_card()
        self._build_time_card()
        self._build_visitor_card()
        self._build_schedule_card()
        self._build_timetable_card()
        self._build_outlook_card()
        self._build_network_card()

        # Initial data
        self._refresh_current_status()
        self._tick()
        self._check_visitor_queue()

        # Auto-connect if URL was previously saved
        saved_url = _load_ics_url()
        if ICS_AVAILABLE and saved_url:
            self.ics_url_entry.insert(0, saved_url)
            self._start_ics_thread(saved_url)

    # ═══════════════════════ HEADER ═══════════════════════════

    def _send_network_update(self, status: str, return_time: str, source: str):
        if not send_status_update:
            return
            
        ip = self.display_ip_entry.get().strip() if hasattr(self, 'display_ip_entry') else _load_display_ip()
        if not ip:
            return
            
        def _send():
            ok, err = send_status_update(ip, status, return_time, source=source)
            if hasattr(self, 'net_status_lbl'):
                if ok:
                    self.net_status_lbl.configure(text="✓ Connected", text_color="#10b981")
                else:
                    self.net_status_lbl.configure(text=f"⚠ Error", text_color="#ef4444")
        
        threading.Thread(target=_send, daemon=True).start()

    # ═══════════════ STATUS APPLICATION ═══════════════════════

    def _apply_status(self, status_text: str, return_time: str = ""):
        if not return_time:
            try:
                return_time = self.return_entry.get().strip()
            except Exception:
                return_time = ""

        new_priority = get_priority(status_text)
        current = read_status()

        # Bug Fix: If we are currently in an Outlook meeting, the UI box contains the meeting's end time.
        # If the user clicks a manual status button (like 'Available'), we don't want to carry
        # that meeting time into the manual state unless they specifically typed a different one.
        if current["source"] == "outlook" and return_time == current.get("return_time"):
            return_time = ""

        if self._ics_meeting_active:
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
        self._send_network_update(status_text, return_time, "manual")
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

                now = datetime.now(timezone.utc).astimezone()

                def to_local(dt):
                    if isinstance(dt, datetime):
                        return dt.astimezone() if dt.tzinfo else dt.replace(tzinfo=now.tzinfo)
                    return dt  # date object

                # 1. Fetch current status (for active meeting detection)
                now_utc = datetime.now(timezone.utc)
                start_win = now_utc - timedelta(minutes=1)
                end_win   = now_utc + timedelta(minutes=1)

                events_now = recurring_ical_events.of(cal).between(start_win, end_win)
                meeting_found = len(events_now) > 0
                meeting_subject = ""
                busy_status = 0
                meeting_end_time_str = ""
                latest_active_end_dt = None

                for event in events_now:
                    meeting_subject = str(event.get("SUMMARY", "Untitled"))
                    transp = str(event.get("TRANSP", "OPAQUE")).upper()
                    if transp != "TRANSPARENT":
                        busy_status = max(busy_status, 2)
                        
                        end_dt_raw = event.get("DTEND")
                        if end_dt_raw:
                            end_dt_local = to_local(end_dt_raw.dt)
                            if isinstance(end_dt_local, datetime):
                                if latest_active_end_dt is None or end_dt_local > latest_active_end_dt:
                                    latest_active_end_dt = end_dt_local
                                    meeting_end_time_str = end_dt_local.strftime("%I:%M %p").lstrip("0")

                # 2. Fetch full schedule for the NEXT 7 DAYS (local timezone)
                fetch_start = now
                fetch_end = now + timedelta(days=7)

                upcoming_events = recurring_ical_events.of(cal).between(fetch_start, fetch_end)

                sched_list = []

                for event in upcoming_events:
                    title = str(event.get("SUMMARY", "Untitled"))

                    start_dt = to_local(event.get("DTSTART").dt)
                    end_dt = to_local(event.get("DTEND").dt)

                    # 🔴 Skip past events
                    if isinstance(end_dt, datetime):
                        if end_dt < now:
                            continue
                    else:
                        if end_dt < now.date():
                            continue

                    # 🕒 Format helpers
                    def fmt_time(dt):
                        return dt.strftime("%I:%M %p") if isinstance(dt, datetime) else "All Day"

                    def fmt_date(dt):
                        return dt.strftime("%b %d")

                    duration = fmt_time(start_dt)
                    if isinstance(end_dt, datetime):
                        duration += f" - {fmt_time(end_dt)}"

                    sched_list.append({
                        "date": fmt_date(start_dt),
                        "title": title,
                        "duration": duration,
                        "start_dt": start_dt  # for sorting
                    })

                # Sort by actual datetime 
                sched_list.sort(key=lambda x: x["start_dt"])

                # Optional: limit rows (UI safe)
                sched_list = sched_list[:20]

                # Remove helper field before UI use
                for item in sched_list:
                    item.pop("start_dt", None)

                self.upcoming_schedule_data = sched_list

                # 3. Apply status logic
                self._ics_meeting_active = meeting_found
                if meeting_found and busy_status > 0:
                    current = read_status()
                    target_status = "In a Meeting (Outlook)"
                    target_priority = get_priority(target_status)
                    if target_priority > current["priority"] or current["source"] == "outlook":
                        write_status(target_status, return_time=meeting_end_time_str, source="outlook")
                        self._send_network_update(target_status, meeting_end_time_str, "outlook")
                    self._ics_status_text = f"✅ Active: \"{meeting_subject}\" ({len(sched_list)} upcoming)"
                else:
                    self._ics_meeting_active = False
                    current = read_status()
                    if current["source"] == "outlook":
                        last_status = current.get("last_manual_status", "Available")
                        last_time = current.get("last_manual_return_time", "")
                        base_status = last_status.split(" — ")[0].strip()
                        write_status(base_status, return_time=last_time, source="manual")
                        self._send_network_update(base_status, last_time, "manual")
                    self._ics_status_text = f"✅ Sync OK ({len(sched_list)} upcoming events)"

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

    # ═══════════════ VISITOR NETWORK + EMAIL ═══════════════════

    def _on_visitor_received(self, name, purpose, timestamp):
        """Called from network server thread when a visitor message arrives."""
        self._visitor_queue.put({"name": name, "purpose": purpose, "timestamp": timestamp})

    def _check_visitor_queue(self):
        """Check for incoming visitor messages from the network."""
        import queue as q
        try:
            while True:
                msg = self._visitor_queue.get_nowait()
                # Save to local DB
                from db_manager import add_visitor
                add_visitor(msg["name"], msg["purpose"])
                self._refresh_visitors()

                # Send phone push notification
                self._send_ntfy_notification(msg["name"], msg["purpose"])
        except q.Empty:
            pass
        self.after(2000, self._check_visitor_queue)

    def _save_ntfy_topic(self):
        topic = self.ntfy_topic_entry.get().strip()
        if topic:
            # topic must be alphanumeric roughly
            topic = ''.join(e for e in topic if e.isalnum() or e in '_-')
            _save_ntfy_topic(topic)
            self.ntfy_topic_entry.delete(0, 'end')
            self.ntfy_topic_entry.insert(0, topic)
            self.ntfy_status_lbl.configure(text="✓ Saved", text_color="#10b981")
        else:
            self.ntfy_status_lbl.configure(text="Enter a code", text_color="#ef4444")

    def _send_ntfy_notification(self, visitor_name: str, message: str):
        """Send push notification to the owner's phone via ntfy.sh."""
        topic = self.ntfy_topic_entry.get().strip() if hasattr(self, 'ntfy_topic_entry') else _load_ntfy_topic()
        if not topic:
            return

        def _send():
            try:
                import urllib.request
                import json
                
                req = urllib.request.Request(f"https://ntfy.sh/{topic}", method="POST")
                req.add_header("Title", f"Door Sign: {visitor_name}".encode('utf-8'))
                req.add_header("Tags", "door")
                req.add_header("Priority", "default")
                
                body = f"Message: {message}".encode('utf-8')
                urllib.request.urlopen(req, data=body, timeout=10)
                
                self.after(0, lambda: self.ntfy_status_lbl.configure(
                    text="✓ Notified", text_color="#10b981"))
            except Exception as e:
                self.after(0, lambda: self.ntfy_status_lbl.configure(
                    text=f"⚠ Failed", text_color="#ef4444"))

        threading.Thread(target=_send, daemon=True).start()

    # ═══════════════════ CLEANUP ══════════════════════════════

    def destroy(self):
        self._ics_running = False
        super().destroy()


if __name__ == "__main__":
    ControlApp().mainloop()
