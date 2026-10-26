"""Recipient routing rules for OPS HUB toilet service requests.

This module decides *who* should be notified; it does not send messages.
Telegram eligibility (active account and chat ID) is handled separately.
"""


def select_recipients(toilet, when, users):
    """Return eligible user dictionaries selected by location and Singapore time.

    The caller supplies eligible users with 'username' and 'group' keys.
    Time boundaries use hour and minute, matching the established rules.
    Demo-zone toilets route to their configured group at any time.
    """
    tower = str(toilet.get('tower') or '').strip().upper().replace('TOWER ', '')
    group = str(toilet.get('supervisor_group') or '').strip().upper()
    zone = str(toilet.get('zone') or '').strip().lower()
    minute = when.hour * 60 + when.minute

    if zone == 'demo':
        return [u for u in users if str(u.get('group') or '').strip().upper() == group]

    if minute >= 21 * 60 + 1 or minute < 7 * 60:
        names = {'sangkele'}
    elif tower == 'A' and minute >= 16 * 60 + 1:
        names = {'myrna'}
    elif tower == 'A' and group == 'AB2-B1':
        names = {'zawhtet'}
    else:
        return [u for u in users if str(u.get('group') or '').strip().upper() == group]

    return [u for u in users if str(u.get('username') or '').strip().lower() in names]
