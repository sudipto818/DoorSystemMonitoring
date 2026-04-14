"""Shared UI constants for Door Sign apps."""

FONT_TITLE = ("Segoe UI", 22, "bold")
FONT_SECTION = ("Segoe UI", 13, "bold")
FONT_BODY = ("Segoe UI", 13)
FONT_SMALL = ("Segoe UI", 11)
FONT_BTN = ("Segoe UI", 12, "bold")
FONT_TINY = ("Segoe UI", 10)

BG_BASE = "#f8f9fa"
BG_CARD = "#ffffff"
BG_INPUT = "#f1f3f5"

FG_PRIMARY = "#212529"
FG_MUTED = "#6c757d"
FG_DIM = "#adb5bd"

ACCENT = "#4263eb"
BORDER = "#dee2e6"

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


def theme_for(status_text: str):
    base = status_text.split(" — ")[0].strip()
    return THEMES.get(base, DEFAULT_THEME)


def blend(hex1: str, hex2: str, factor: float = 0.5) -> str:
    h1, h2 = hex1.lstrip("#"), hex2.lstrip("#")
    r1, g1, b1 = (int(h1[i:i+2], 16) for i in (0, 2, 4))
    r2, g2, b2 = (int(h2[i:i+2], 16) for i in (0, 2, 4))
    r = int(r1 + (r2 - r1) * factor)
    g = int(g1 + (g2 - g1) * factor)
    b = int(b1 + (b2 - b1) * factor)
    return f"#{r:02x}{g:02x}{b:02x}"
