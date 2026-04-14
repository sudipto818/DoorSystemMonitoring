"""
voice_command.py
─────────────────────────────────────────────────────
Local, offline voice-command module for the Door-Sign-Status system.
Handles audio recording, speech-to-text transcription (via faster-whisper),
and rule-based intent parsing.

Integrates with db_manager.write_status() and the control panel UI.
"""

import os
import sys
import re
import wave
import tempfile
import difflib
from datetime import datetime, timedelta, date as date_type
from pathlib import Path

# ═══════════════════════════════════════════════════════════════════════
#  Optional dependency imports — degrade gracefully if missing
# ═══════════════════════════════════════════════════════════════════════

try:
    import numpy as np
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except ImportError:
    AUDIO_AVAILABLE = False

try:
    from faster_whisper import WhisperModel
    WHISPER_AVAILABLE = True
except ImportError:
    WHISPER_AVAILABLE = False


def _get_base_dir() -> str:
    """Return the folder where the app lives (works for .py and .exe)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# ═══════════════════════════════════════════════════════════════════════
#  AUDIO RECORDER
# ═══════════════════════════════════════════════════════════════════════

class AudioRecorder:
    """Records audio from the default microphone and saves to a WAV file."""

    SAMPLE_RATE = 44100
    CHANNELS = 1
    DTYPE = "int16"

    def __init__(self):
        self._frames: list = []
        self._stream = None
        self._recording = False

    @property
    def is_recording(self) -> bool:
        return self._recording

    def start(self):
        """Begin capturing audio from the microphone."""
        if not AUDIO_AVAILABLE:
            raise RuntimeError(
                "Audio recording requires 'sounddevice' and 'numpy'. "
                "Install them with:  pip install sounddevice numpy"
            )
        self._frames = []
        self._recording = True
        self._stream = sd.InputStream(
            samplerate=self.SAMPLE_RATE,
            channels=self.CHANNELS,
            dtype=self.DTYPE,
            callback=self._audio_callback,
        )
        self._stream.start()

    def stop(self) -> str:
        """Stop recording and save audio to a WAV file. Returns the filepath."""
        if self._stream is not None:
            self._stream.stop()
            self._stream.close()
            self._stream = None
        self._recording = False

        if not self._frames:
            return ""

        # Combine all captured frames
        audio_data = np.concatenate(self._frames, axis=0)

        # Save to a temp WAV file in the project directory
        filepath = os.path.join(_get_base_dir(), "temp_audio.wav")
        with wave.open(filepath, "wb") as wf:
            wf.setnchannels(self.CHANNELS)
            wf.setsampwidth(2)  # 16-bit = 2 bytes
            wf.setframerate(self.SAMPLE_RATE)
            wf.writeframes(audio_data.tobytes())

        return filepath

    def _audio_callback(self, indata, frames, time_info, status):
        """Callback invoked by sounddevice for each audio block."""
        if self._recording:
            self._frames.append(indata.copy())


# ═══════════════════════════════════════════════════════════════════════
#  VOICE PROCESSOR — Transcription + Intent Parsing
# ═══════════════════════════════════════════════════════════════════════

class VoiceProcessor:
    """Offline speech-to-text using faster-whisper + rule-based parsing."""

    def __init__(self):
        self._model = None

    # ─────────────── Model Management ───────────────

    def load_model(self):
        """Lazily load the faster-whisper 'tiny.en' model for CPU inference."""
        if not WHISPER_AVAILABLE:
            raise RuntimeError(
                "Transcription requires 'faster-whisper'. "
                "Install with:  pip install faster-whisper"
            )
        if self._model is None:
            self._model = WhisperModel(
                "tiny.en",
                device="cpu",
                compute_type="int8",
            )
        return self._model

    @property
    def is_loaded(self) -> bool:
        return self._model is not None

    # ─────────────── Transcription ───────────────

    def transcribe(self, filepath: str) -> str:
        """Transcribe an audio file and return the raw text."""
        model = self.load_model()
        segments, _info = model.transcribe(filepath, beam_size=1)
        text_parts = [segment.text for segment in segments]
        return " ".join(text_parts).strip()

    # ─────────────── Fuzzy Match Helper ───────────────

    @staticmethod
    def _fuzzy_contains(text: str, phrase: str, cutoff: float = 0.75) -> bool:
        """
        Check if `phrase` appears in `text`, allowing for slight
        transcription inaccuracies via difflib.SequenceMatcher.

        First tries exact substring match (fast path), then falls back
        to sliding-window fuzzy comparison over word n-grams.
        """
        # Fast path: exact substring
        if phrase in text:
            return True

        # Fuzzy path: compare phrase against every n-gram window in text
        phrase_words = phrase.split()
        text_words = text.split()
        n = len(phrase_words)
        if n == 0 or len(text_words) < n:
            return False

        for i in range(len(text_words) - n + 1):
            window = " ".join(text_words[i:i + n])
            ratio = difflib.SequenceMatcher(None, phrase, window).ratio()
            if ratio >= cutoff:
                return True

        # Single-word phrases: also check individual words
        if n == 1:
            matches = difflib.get_close_matches(phrase, text_words, n=1, cutoff=cutoff)
            return len(matches) > 0

        return False

    # ─────────────── Date Extraction ───────────────

    @staticmethod
    def _extract_date(text: str):
        """
        Extract a target date from voice text.
        Returns a datetime.date or None.

        Handles: 'today', 'tomorrow', day names ('monday', 'tuesday', etc.)
        """
        today = datetime.now().date()

        if "today" in text:
            return today
        if "tomorrow" in text:
            return today + timedelta(days=1)

        # Day-of-week matching
        day_names = ["monday", "tuesday", "wednesday", "thursday",
                     "friday", "saturday", "sunday"]
        for i, day in enumerate(day_names):
            if day in text:
                current_weekday = today.weekday()  # Mon=0 ... Sun=6
                target_weekday = i
                days_ahead = (target_weekday - current_weekday) % 7
                if days_ahead == 0:
                    days_ahead = 7  # next week if same day
                return today + timedelta(days=days_ahead)

        return None  # no date found, will default to today

    # ─────────────── Intent Parsing ───────────────

    def parse_command(self, text: str) -> dict:
        """
        Parse transcribed text into a structured intent dictionary.
        Uses fuzzy matching (difflib) for robustness against transcription errors.

        Returns one of:
            {"action": "set_status",     "status": ..., "return_time": ..., "meeting": None}
            {"action": "create_meeting", "meeting": {...}}
            {"action": "unknown",        "status": "",  "return_time": ""}
        """
        lower = text.lower().strip()

        # ── 0. Quick check: "in a meeting" is a STATUS, not a creation intent ──
        # Must check this BEFORE meeting-creation phrases to avoid false positives
        in_meeting_phrases = ["in a meeting", "i'm in a meeting", "i am in a meeting"]
        is_status_meeting = any(self._fuzzy_contains(lower, p) for p in in_meeting_phrases)

        # ── 1. Check for meeting-creation intent (fuzzy) ──
        # Only trigger if the user explicitly wants to CREATE/SCHEDULE a meeting
        if not is_status_meeting:
            meeting_phrases = [
                "schedule a meeting", "schedule meeting", "schedule",
                "create a meeting", "create meeting",
                "new meeting", "set up a meeting", "book a meeting",
            ]
            for phrase in meeting_phrases:
                if self._fuzzy_contains(lower, phrase):
                    # Extract duration (e.g. "30 minute meeting")
                    duration = 30  # default
                    dur_match = re.search(r"(\d+)\s*(?:minute|min)", lower)
                    if dur_match:
                        duration = int(dur_match.group(1))

                    # Extract target date from voice
                    target_date = self._extract_date(lower)

                    return {
                        "action": "create_meeting",
                        "meeting": {
                            "title": "Voice Scheduled Meeting",
                            "duration_minutes": duration,
                            "target_date": target_date,  # None = today
                        },
                    }

        # ── 2. Map keywords to statuses (fuzzy) ──
        status_map = [
            (["knock"],                              "Please Knock"),
            (["available", "free", "open"],           "Available"),
            (["disturb", "dnd", "do not disturb"],    "Do Not Disturb"),
            (["meeting", "in a meeting"],             "In a Meeting"),
            (["remote", "remotely", "wfh",
              "work from home", "working remotely"],  "Working remotely"),
            (["back soon", "brb", "be right back"],   "Back soon"),
            (["out of office", "ooo"],                "Out of office"),
            (["busy"],                                "Busy"),
        ]

        matched_status = ""
        for keywords, target in status_map:
            for kw in keywords:
                if self._fuzzy_contains(lower, kw):
                    matched_status = target
                    break
            if matched_status:
                break

        # ── 3. Extract return / availability time ──
        return_time = self._extract_time(lower)

        # ── 3b. Check if user specifically wants to SET availability time ──
        # e.g. "set return time to 3 pm", "back by 4 pm", "I'll be back by 5"
        time_phrases = [
            "time of availability", "availability time",
            "set time", "set availability", "available at",
            "available by", "available until",
            "return time", "set return", "expected time",
            "back by", "will be back", "be back",
        ]
        is_time_intent = any(self._fuzzy_contains(lower, p) for p in time_phrases)

        # If user explicitly asked to set time AND we extracted a time,
        # return a 'set_time' action (no status change, just time update)
        if is_time_intent and return_time and not matched_status:
            return {
                "action": "set_time",
                "status": "",
                "return_time": return_time,
                "meeting": None,
            }

        # ── 4. Build response ──
        if matched_status:
            return {
                "action": "set_status",
                "status": matched_status,
                "return_time": return_time,
                "meeting": None,
            }

        # If we extracted a time but no status matched, treat it as set_time
        if return_time:
            return {
                "action": "set_time",
                "status": "",
                "return_time": return_time,
                "meeting": None,
            }

        # Fallback — truly couldn't understand
        return {
            "action": "unknown",
            "status": "",
            "return_time": "",
        }

    # ─────────────── Time Extraction Helpers ───────────────

    @staticmethod
    def _extract_time(text: str) -> str:
        """
        Extract a return time from text.

        Handles:
          - Absolute:  "back by 3 pm", "at 4:00", "until 2:30 pm"
          - Relative:  "in 30 minutes", "in 2 hours", "half an hour"
        """
        # ── Relative: "in X minutes" / "in X hours" ──
        rel_min = re.search(r"in\s+(\d+)\s*(?:minute|min)", text)
        if rel_min:
            delta = timedelta(minutes=int(rel_min.group(1)))
            return (datetime.now() + delta).strftime("%I:%M %p")

        rel_hr = re.search(r"in\s+(\d+)\s*(?:hour|hr)", text)
        if rel_hr:
            delta = timedelta(hours=int(rel_hr.group(1)))
            return (datetime.now() + delta).strftime("%I:%M %p")

        # "half an hour" / "half hour"
        if "half an hour" in text or "half hour" in text:
            delta = timedelta(minutes=30)
            return (datetime.now() + delta).strftime("%I:%M %p")

        # "an hour"
        if re.search(r"\ban\s+hour\b", text):
            delta = timedelta(hours=1)
            return (datetime.now() + delta).strftime("%I:%M %p")

        # ── Absolute: "back by 3 pm", "at 4:00 pm", "until 2:30", "to 3 pm" ──
        abs_match = re.search(
            r"(?:back\s+by|at|until|by|to)\s+"
            r"(\d{1,2})(?::(\d{2}))?\s*"
            r"(am|pm|a\.m\.|p\.m\.)?",
            text,
        )
        if abs_match:
            hour = int(abs_match.group(1))
            minute = int(abs_match.group(2) or 0)
            ampm = (abs_match.group(3) or "").replace(".", "").strip().lower()

            if ampm == "pm" and hour != 12:
                hour += 12
            elif ampm == "am" and hour == 12:
                hour = 0
            elif not ampm:
                # If no AM/PM given, assume PM for typical office hours
                if hour < 7:
                    hour += 12

            try:
                target = datetime.now().replace(
                    hour=hour, minute=minute, second=0, microsecond=0
                )
                return target.strftime("%I:%M %p")
            except ValueError:
                pass

        # ── Bare time: "3 pm", "4:30 pm" anywhere in text (last resort) ──
        bare_match = re.search(
            r"(\d{1,2})(?::(\d{2}))?\s*(am|pm|a\.m\.|p\.m\.)",
            text,
        )
        if bare_match:
            hour = int(bare_match.group(1))
            minute = int(bare_match.group(2) or 0)
            ampm = bare_match.group(3).replace(".", "").strip().lower()

            if ampm == "pm" and hour != 12:
                hour += 12
            elif ampm == "am" and hour == 12:
                hour = 0

            try:
                target = datetime.now().replace(
                    hour=hour, minute=minute, second=0, microsecond=0
                )
                return target.strftime("%I:%M %p")
            except ValueError:
                pass

        return ""


# ═══════════════════════════════════════════════════════════════════════
#  CALENDAR UTILITY — Create .ics Meeting File
# ═══════════════════════════════════════════════════════════════════════

def create_ics_meeting(title: str = "Voice Scheduled Meeting",
                       duration_minutes: int = 30,
                       target_date=None) -> str:
    """
    Generate a valid .ics calendar invitation and save it to disk.
    On Windows, automatically opens the file in the default calendar app.

    Meeting start = target_date at (current_time + 30 minutes).
    If target_date is None, defaults to today.

    Returns the absolute path of the created .ics file.
    """
    now = datetime.now()

    # +30 minute rule: meeting starts 30 minutes from now
    start_time = now + timedelta(minutes=30)

    # Apply to target date if specified, otherwise use today
    if target_date is not None:
        # Combine the target date with the +30min time-of-day
        start = datetime.combine(target_date, start_time.time())
    else:
        start = start_time

    end = start + timedelta(minutes=duration_minutes)

    # Unique filename based on timestamp
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"meeting_{timestamp}.ics"
    filepath = os.path.join(_get_base_dir(), filename)

    # Format timestamps for iCal (YYYYMMDDTHHMMSS)
    fmt = "%Y%m%dT%H%M%S"
    uid = f"doorsign-{now.strftime('%Y%m%d%H%M%S')}@localhost"

    ics_content = (
        "BEGIN:VCALENDAR\r\n"
        "VERSION:2.0\r\n"
        "PRODID:-//DoorSignStatus//VoiceCommand//EN\r\n"
        "METHOD:REQUEST\r\n"
        "CALSCALE:GREGORIAN\r\n"
        "BEGIN:VEVENT\r\n"
        f"UID:{uid}\r\n"
        f"DTSTART:{start.strftime(fmt)}\r\n"
        f"DTEND:{end.strftime(fmt)}\r\n"
        f"SUMMARY:{title}\r\n"
        "DESCRIPTION:Meeting created via Door Sign voice command.\r\n"
        "STATUS:CONFIRMED\r\n"
        f"DTSTAMP:{now.strftime(fmt)}\r\n"
        f"CREATED:{now.strftime(fmt)}\r\n"
        "TRANSP:OPAQUE\r\n"
        "END:VEVENT\r\n"
        "END:VCALENDAR\r\n"
    )

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(ics_content)

    # On Windows, open in Outlook using subprocess for reliability
    if os.name == "nt":
        try:
            import subprocess
            subprocess.Popen(['cmd', '/c', 'start', '', filepath], shell=False)
        except OSError:
            try:
                os.startfile(filepath)  # fallback
            except OSError:
                pass

    return filepath


# ═══════════════════════════════════════════════════════════════════════
#  QUICK TEST (run this file directly to test)
# ═══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    vp = VoiceProcessor()

    test_phrases = [
        "Set my status to available",
        "I'm in a meeting back by 3 pm",
        "Do not disturb for the next hour",
        "Set status to busy in 30 minutes",
        "Working remotely until 5 pm",
        "Schedule a new meeting",
        "I'll be back soon",
        "Out of office",
        "Please knock back by 2:30 pm",
        "Create a 45 minute meeting",
        # Fuzzy matching tests (simulated transcription errors)
        "Sett my stattus to availabel",
        "Scedule a new meting for tomorrow",
        "I'm werking remotly",
        "Crate a meeting for monday",
    ]

    print("-" * 60)
    print("  Voice Command Parser - Test Results")
    print("-" * 60)

    for phrase in test_phrases:
        result = vp.parse_command(phrase)
        print(f"\n  Input:  \"{phrase}\"")
        print(f"  Result: {result}")

    print("\n" + "-" * 60)
