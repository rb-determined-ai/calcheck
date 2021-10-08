# `calcheck.py`

Scriptable alerts for Linux users of OWA.

## Requirements

- `python3` for `calcheck.py`.  There are no additional dependencies.

- `systemd` for time-based triggering.

- An outlook account, to provide the `.ics`.  The embedded ics parser is pretty
  limited, so `calcheck.py` isn't likely to work with other providers.

## Configuring

1. Share your outlook calendar with an email address you control, then find a
   url in that email (possibly in an xml attachment) looking like:
   `https://outlook.office365.com/.../reachcalendar.ics`; set that as your URL
   in `calcheck.py`.

1. Update the rest of the `CONFIG` section of `calcheck.py` according to how
   you want to be notified.

1. Double-check the `calcheck.timer` script in `install-systemd.sh` to make
   sure the triggering is suitable to your needs.

1. Run `sh ./install-systemd.sh` to configure a user.

1. Run `systemd --user status calcheck.time` to verify your installation.
