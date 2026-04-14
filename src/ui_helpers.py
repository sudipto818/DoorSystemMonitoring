"""Shared helper widgets for Door Sign apps."""
import customtkinter as ctk
from ui_constants import BG_CARD, BORDER, FONT_SECTION, FG_MUTED


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
