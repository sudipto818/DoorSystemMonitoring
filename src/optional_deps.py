"""Optional dependency detection for Door Sign.

These values are imported by both app modules so optional package
availability is centralized and there is no circular import.
"""

try:
    import requests as http_requests
    import recurring_ical_events
    import icalendar
    ICS_AVAILABLE = True
except ImportError:
    http_requests = None
    recurring_ical_events = None
    icalendar = None
    ICS_AVAILABLE = False

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    openpyxl = None
    EXCEL_AVAILABLE = False
