#!/usr/bin/env python3

import datetime
import os
import re
import sys
import time
import urllib.request

try:
    import zoneinfo
except:
    # pre-python 3.9 requires an extra packages
    from backports import zoneinfo


### BEGIN CONFIG SECTION

# URL is where the calendar is fetched from
URL = (
    "https://outlook.office365.com/.../reachcalendar.ics"
    # or os.environ["OUTLOOK_URL"]
)

# WINDOW is the time window for which you want to detect events when the script runs.
# This must play nicely with the trigger times set up in calcheck.timer.
WINDOW = 6*60

# on_event is called for every future event within the WINDOW
def on_event(epoch, summary):
    import subprocess
    cmd = notify_command(f"upcoming event: {summary}")
    subprocess.Popen(cmd).wait()

def on_failure(msg):
    import subprocess
    subprocess.Popen(notify_command(f"calcheck failed: {msg}")).wait()

def notify_command(message):
    if sys.platform.startswith("linux"):
        return ["notify-send", "-u", "critical", message]
    elif sys.platform == "darwin":
        return ["osascript", "-e", f'display notification "{message}" with title "calcheck"']

    elif sys.platform == "win32":
        raise ValueError("Unsupported operating system")

### END CONFIG SECTION


def epoch_with_zone(dtstart, iana_zone):
    # use time.strptime to read the fields as a named tuple
    ambiguous = time.strptime(dtstart, "%Y%m%dT%H%M%S")
    # pass what was parsed into a datetime with zone info
    zoned_time = datetime.datetime(*ambiguous[:6], tzinfo=zoneinfo.ZoneInfo(iana_zone))
    # return epoch time
    return int(zoned_time.astimezone().strftime("%s"))


def to_utf8(lines):
    for line in lines:
        if isinstance(line, bytes):
            line = line.decode('utf8')
        yield line


def vevents(f):
    lines = iter(f)
    while True:
        # unmached state
        for line in lines:
            if "BEGIN:VEVENT" in line:
                break
        else:
            break

        # matched state
        vevent = []
        full_line = ""
        for line in lines:
            if not line.startswith(" "):
                # not folding whitespace, emit what we've gathered
                if full_line:
                    vevent.append(full_line.split(":", 1))
                    full_line = ""
                full_line += line.strip()
            else:
                full_line += line[1:].rstrip("\n")
            if "END:VEVENT" in line:
                yield vevent
                break
        else:
            # break the while loop
            break

def read_key(vevent, key):
    for k, v in vevent:
        if k.split(";")[0] == key:
            return v
    return None

def key_params(vevent, key):
    for k, v in vevent:
        if k.split(";")[0] == key:
            params = {}
            for pair in k.split(";")[1:]:
                kk, vv = pair.split("=", 1)
                params[kk] = vv
            return params
    return None

def dup_vevent(vevent, **replacements):
    out = []
    for k, v in vevent:
        if k.split(";")[0] in replacements:
            out.append(replacements[k.split(";")[0]])
        else:
            out.append((k, v))
    return out


def day_to_num(day):
    dmap = {
        "MO": 1,
        "TU": 2,
        "WE": 3,
        "TH": 4,
        "FR": 5,
        "SA": 6,
        "SU": 7,
    }
    return dmap[day]

def recur(vevent, rrule):
    params = {}
    for pair in rrule.split(";"):
        k, v = pair.split("=", 1)
        params[k] = v
    summary = read_key(vevent, "SUMMARY")
    dtstart = read_key(vevent, "DTSTART")
    dtstart_params = key_params(vevent, "DTSTART")

    # Daily recurrance logic.
    if params["FREQ"] == "DAILY":
        assert "INTERVAL" in params, f"RRULE:FREQ=DAILY has no INTERVAL: {rrule}"
        assert "UNTIL" in params, f"RRULE:FREQ=DAILY has no UNTIL: {rrule}"
        until = datetime.datetime.strptime(params["UNTIL"], "%Y%m%dT%H%M%SZ")
        interval = int(params["INTERVAL"])
        end = datetime.datetime.now() + datetime.timedelta(days=14)
        if until < end:
            end = until
        tm = datetime.datetime.strptime(dtstart, "%Y%m%dT%H%M%S")
        tm += datetime.timedelta(days=interval)
        while tm < end:
            new_dtstart_val = tm.strftime("%Y%m%dT%H%M%S")
            new_dtstart_key = "DTSTART;" + ";".join(f"{k}={v}" for k,v in dtstart_params.items())
            yield dup_vevent(vevent, DTSTART=(new_dtstart_key, new_dtstart_val))
            tm += datetime.timedelta(days=interval)
        return

    # Weekly recurrance logic.
    if params["FREQ"] == "WEEKLY":
        assert "BYDAY" in params, f"RRULE:FREQ=WEEKLY has no BYDAY: {rrule}"
        assert "UNTIL" in params, f"RRULE:FREQ=WEEKLY has no UNTIL: {rrule}"
        days = [day_to_num(d) for d in params["BYDAY"].split(",")]
        until = datetime.datetime.strptime(params["UNTIL"], "%Y%m%dT%H%M%SZ")
        interval = int(params.get("INTERVAL", "1"))
        tm = datetime.datetime.strptime(dtstart, "%Y%m%dT%H%M%S")
        end = datetime.datetime.now() + datetime.timedelta(days=14)
        if until < end:
            end = until
        # repeatedly add days, and check if they land in the BYDAY spec
        tm += datetime.timedelta(days=1)
        while tm < end:
            if tm.isoweekday() in days:
                new_dtstart_val = tm.strftime("%Y%m%dT%H%M%S")
                new_dtstart_key = "DTSTART;" + ";".join(f"{k}={v}" for k,v in dtstart_params.items())
                yield dup_vevent(vevent, DTSTART=(new_dtstart_key, new_dtstart_val))
            tm += datetime.timedelta(days=1)
            # handle intervals: skip from sunday to sunday
            if interval > 1 and tm.isoweekday() == 7:
                for _ in range(1, interval):
                    tm += datetime.timedelta(days=7)
        return

    if params["FREQ"] == "MONTHLY":
        assert "BYDAY" in params, f"RRULE:FREQ=MONTHLY has no BYDAY"
        assert "UNTIL" in params, f"RRULE:FREQ=WEEKLY has no UNTIL: {rrule}"
        # BYDAY examples:
        #    1TH: "first thursday of the month"
        #    2WE: "second wednesday of the month"
        #    -1FR: "last friday of the month"
        #    -2MO: "sencond-to-last monday of the month"
        interval = int(params.get("INTERVAL", "1"))
        until = datetime.datetime.strptime(params["UNTIL"], "%Y%m%dT%H%M%SZ")
        byday = params["BYDAY"]
        nth = int(byday[0:-2])
        day_num = day_to_num(byday[-2:])
        start = datetime.datetime.strptime(dtstart, "%Y%m%dT%H%M%S")
        tm = datetime.datetime.strptime(dtstart, "%Y%m%dT%H%M%S")
        end = datetime.datetime.now() + datetime.timedelta(days=14)
        if until < end:
            end = until

        def month_matches(tm):
            # Match INTERVAL
            monthdiff = (tm.year-start.year)*12 + tm.month - start.month
            return monthdiff % interval == 0

        def day_matches(tm):
            # Match the day of the week.
            if tm.isoweekday() != day_num:
                return False
            # Match the nth spec.
            if nth > 0:
                # Example: nth=2; (tm - 1weeks) should be the same month, but not (tm - 2weeks)
                test1 = tm - datetime.timedelta(weeks=nth-1)
                test2 = tm - datetime.timedelta(weeks=nth)
            else:
                # Example: nth=-2; (tm + 1weeks) should be the same month, but not (tm + 2weeks)
                test1 = tm + datetime.timedelta(weeks=(-nth)-1)
                test2 = tm + datetime.timedelta(weeks=-nth)
            return test1.month == tm.month and test2.month != tm.month

        # repeatedly add days, and check if they match the BYDAY spec
        tm += datetime.timedelta(days=1)
        while tm < end:
            if month_matches(tm) and day_matches(tm):
                new_dtstart_val = tm.strftime("%Y%m%dT%H%M%S")
                new_dtstart_key = "DTSTART;" + ";".join(f"{k}={v}" for k,v in dtstart_params.items())
                yield dup_vevent(vevent, DTSTART=(new_dtstart_key, new_dtstart_val))
            tm += datetime.timedelta(days=1)
        return

    raise ValueError(f"unsupported FREQ: '{params['FREQ']}'")


def recurrance(vevents):
    for vevent in vevents:
        summary = read_key(vevent, "SUMMARY")
        yield vevent
        rrule = read_key(vevent, "RRULE")
        if rrule:
            yield from recur(vevent, rrule)


def ignore_all_day_events(vevents):
    for vevent in vevents:
        allday = read_key(vevent, "X-MICROSOFT-CDO-ALLDAYEVENT")
        if allday and allday.lower() == "true":
            continue
        yield vevent


def detect_upcoming_events(url, window, hook, now, win2iana):
    with urllib.request.urlopen(url) as f:
        vevs = vevents(to_utf8(f))
        vevs = ignore_all_day_events(vevs)
        vevs = recurrance(vevs)
        for vevent in vevs:
            for obj in vevent:
                assert len(obj) == 2
            dtstart = None
            tzid = None
            summary = None
            for key, val in vevent:
                if key.startswith("DTSTART"):
                    dtstart = val
                    tzid = key[13:]
                if key == "SUMMARY":
                    summary = val
            epoch = epoch_with_zone(dtstart, win2iana[tzid])

            diff = epoch - now
            print(diff, dtstart, summary)
            if diff < 0 or diff > window:
                continue
            print("HOOK!", epoch, summary)
            hook(epoch, summary)


def windows_to_iana_timezones():
    file = os.path.join(os.path.dirname(__file__), "windowsZones.xml")
    # Periodically download the latest version of the lookup table.
    if not os.path.exists(file) or os.stat(file).st_mtime + 60*60*30 < time.time():
        url = (
            "https://raw.githubusercontent.com/unicode-org/cldr/"
            "main/common/supplemental/windowsZones.xml"
        )
        with urllib.request.urlopen(url) as f:
            text = f.read().decode("utf8")
        with open(file, "w") as f:
            f.write(text)
    else:
        with open(file, "r") as f:
            text = f.read()
    # Python XML parsers are not safe against malicious documents, and this document is so easy
    # to parse with a regex, (which is safe), we'll just do it that way.
    #
    # Example:
    #
    #    <mapZone other="Hawaiian Standard Time" territory="001" type="Pacific/Honolulu"/>
    table = {}
    pattern = re.compile('<mapZone other="([^"]*)".* type="([^"]*)"')
    for line in text.splitlines():
        match = pattern.search(line)
        if not match:
            continue
        table[match[1]] = match[2]
    return table

if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            now = int(sys.argv[1])
        else:
            now = time.time()
        win2iana = windows_to_iana_timezones()
        detect_upcoming_events(URL, WINDOW, on_event, now, win2iana)
    except Exception as e:
        on_failure(str(e))
        raise
