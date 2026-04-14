"""File helpers for Door Sign storage files."""
import os
import sys


def _get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _load_file(filename: str) -> str:
    path = os.path.join(_get_base_dir(), filename)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""


def _save_file(filename: str, value: str):
    path = os.path.join(_get_base_dir(), filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(value.strip())


def _load_display_ip() -> str:
    return _load_file("display_ip.txt")


def _save_display_ip(ip: str):
    _save_file("display_ip.txt", ip)


def _load_ics_url() -> str:
    return _load_file("outlook_ics_url.txt")


def _save_ics_url(url: str):
    _save_file("outlook_ics_url.txt", url)


def _load_ntfy_topic() -> str:
    return _load_file("ntfy_topic.txt")


def _save_ntfy_topic(topic: str):
    _save_file("ntfy_topic.txt", topic)
