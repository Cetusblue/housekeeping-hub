from datetime import date, timedelta


def compute_run_date(today: date, template_day: str) -> date:
    template_day = template_day.upper()

    if template_day in ("TUE", "ANNEX_TUE"):
        target_weekday = 1  # Tuesday
    elif template_day in ("FRI", "ANNEX_FRI"):
        target_weekday = 4  # Friday
    else:
        raise ValueError("template_day must be 'TUE', 'FRI', 'ANNEX_TUE', or 'ANNEX_FRI'")
    
    days_ahead = target_weekday - today.weekday()

    if days_ahead < 0:
        days_ahead += 7
    elif days_ahead == 0:
        # If created on the same weekday, roll to next week
        days_ahead = 7

    return today + timedelta(days=days_ahead)