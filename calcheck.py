#!/usr/bin/env python3

import time
import datetime
import urllib.request
import sys

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



# Outlook uses their own custom timezone name database, which was dumped (on Windows) with:
#
#    tzutil /l > tzutil_-l.txt
#
# and subsequently processed into the following dictionary with:
#
#    python process_calendar_input.py tzutil_-l.txt
#
tz_offsets = {
    "Dateline Standard Time": -43200,
    "UTC-11": -39600,
    "Aleutian Standard Time": -36000,
    "Hawaiian Standard Time": -36000,
    "Marquesas Standard Time": -34200,
    "Alaskan Standard Time": -32400,
    "UTC-09": -32400,
    "Pacific Standard Time (Mexico)": -28800,
    "UTC-08": -28800,
    "Pacific Standard Time": -28800,
    "US Mountain Standard Time": -25200,
    "Mountain Standard Time (Mexico)": -25200,
    "Mountain Standard Time": -25200,
    "Yukon Standard Time": -25200,
    "Central America Standard Time": -21600,
    "Central Standard Time": -21600,
    "Easter Island Standard Time": -21600,
    "Central Standard Time (Mexico)": -21600,
    "Canada Central Standard Time": -21600,
    "SA Pacific Standard Time": -18000,
    "Eastern Standard Time (Mexico)": -18000,
    "Eastern Standard Time": -18000,
    "Haiti Standard Time": -18000,
    "Cuba Standard Time": -18000,
    "US Eastern Standard Time": -18000,
    "Turks And Caicos Standard Time": -18000,
    "Paraguay Standard Time": -14400,
    "Atlantic Standard Time": -14400,
    "Venezuela Standard Time": -14400,
    "Central Brazilian Standard Time": -14400,
    "SA Western Standard Time": -14400,
    "Pacific SA Standard Time": -14400,
    "Newfoundland Standard Time": -12600,
    "Tocantins Standard Time": -10800,
    "E. South America Standard Time": -10800,
    "SA Eastern Standard Time": -10800,
    "Argentina Standard Time": -10800,
    "Greenland Standard Time": -10800,
    "Montevideo Standard Time": -10800,
    "Magallanes Standard Time": -10800,
    "Saint Pierre Standard Time": -10800,
    "Bahia Standard Time": -10800,
    "UTC-02": -7200,
    "Azores Standard Time": -3600,
    "Cape Verde Standard Time": -3600,
    "UTC": 0,
    "Morocco Standard Time": 0,
    "GMT Standard Time": 0,
    "Greenwich Standard Time": 0,
    "Sao Tome Standard Time": 3600,
    "W. Europe Standard Time": 3600,
    "Central Europe Standard Time": 3600,
    "Romance Standard Time": 3600,
    "Central European Standard Time": 3600,
    "W. Central Africa Standard Time": 3600,
    "Jordan Standard Time": 7200,
    "GTB Standard Time": 7200,
    "Middle East Standard Time": 7200,
    "Egypt Standard Time": 7200,
    "E. Europe Standard Time": 7200,
    "Syria Standard Time": 7200,
    "West Bank Standard Time": 7200,
    "South Africa Standard Time": 7200,
    "FLE Standard Time": 7200,
    "Israel Standard Time": 7200,
    "South Sudan Standard Time": 7200,
    "Kaliningrad Standard Time": 7200,
    "Sudan Standard Time": 7200,
    "Libya Standard Time": 7200,
    "Namibia Standard Time": 7200,
    "Arabic Standard Time": 10800,
    "Turkey Standard Time": 10800,
    "Arab Standard Time": 10800,
    "Belarus Standard Time": 10800,
    "Russian Standard Time": 10800,
    "E. Africa Standard Time": 10800,
    "Volgograd Standard Time": 10800,
    "Iran Standard Time": 12600,
    "Arabian Standard Time": 14400,
    "Astrakhan Standard Time": 14400,
    "Azerbaijan Standard Time": 14400,
    "Russia Time Zone 3": 14400,
    "Mauritius Standard Time": 14400,
    "Saratov Standard Time": 14400,
    "Georgian Standard Time": 14400,
    "Caucasus Standard Time": 14400,
    "Afghanistan Standard Time": 16200,
    "West Asia Standard Time": 18000,
    "Ekaterinburg Standard Time": 18000,
    "Pakistan Standard Time": 18000,
    "Qyzylorda Standard Time": 18000,
    "India Standard Time": 19800,
    "Sri Lanka Standard Time": 19800,
    "Nepal Standard Time": 20700,
    "Central Asia Standard Time": 21600,
    "Bangladesh Standard Time": 21600,
    "Omsk Standard Time": 21600,
    "Myanmar Standard Time": 23400,
    "SE Asia Standard Time": 25200,
    "Altai Standard Time": 25200,
    "W. Mongolia Standard Time": 25200,
    "North Asia Standard Time": 25200,
    "N. Central Asia Standard Time": 25200,
    "Tomsk Standard Time": 25200,
    "China Standard Time": 28800,
    "North Asia East Standard Time": 28800,
    "Singapore Standard Time": 28800,
    "W. Australia Standard Time": 28800,
    "Taipei Standard Time": 28800,
    "Ulaanbaatar Standard Time": 28800,
    "Aus Central W. Standard Time": 31500,
    "North Korea Standard Time": 30600,
    "Transbaikal Standard Time": 32400,
    "Tokyo Standard Time": 32400,
    "Korea Standard Time": 32400,
    "Yakutsk Standard Time": 32400,
    "Cen. Australia Standard Time": 34200,
    "AUS Central Standard Time": 34200,
    "E. Australia Standard Time": 36000,
    "AUS Eastern Standard Time": 36000,
    "West Pacific Standard Time": 36000,
    "Tasmania Standard Time": 36000,
    "Vladivostok Standard Time": 36000,
    "Lord Howe Standard Time": 37800,
    "Bougainville Standard Time": 39600,
    "Russia Time Zone 10": 39600,
    "Magadan Standard Time": 39600,
    "Norfolk Standard Time": 39600,
    "Sakhalin Standard Time": 39600,
    "Central Pacific Standard Time": 39600,
    "Russia Time Zone 11": 43200,
    "New Zealand Standard Time": 43200,
    "UTC+12": 43200,
    "Fiji Standard Time": 43200,
    "Chatham Islands Standard Time": 45900,
    "UTC+13": 46800,
    "Tonga Standard Time": 46800,
    "Samoa Standard Time": 46800,
    "Line Islands Standard Time": 50400,
}


def epoch_with_tzid(dtstart, tzid):
    # local_tm will implicitly be in our current timezone
    local_tm = time.strptime(dtstart, "%Y%m%dT%H%M%S")
    local_epoch = int(time.strftime("%s", local_tm))
    # adjust for our current timezone, to get epoch time, if timestr were gmt
    gmt_epoch = local_epoch - time.timezone
    # now the actual epoch time is just an offset away from what we've calculated
    epoch = gmt_epoch - tz_offsets[tzid]
    return epoch


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
    tzid = dtstart_params["TZID"]

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


def detect_upcoming_events(url, window, hook, now):
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
            epoch = epoch_with_tzid(dtstart, tzid)

            diff = epoch - now
            print(diff, dtstart, summary)
            if diff < 0 or diff > window:
                continue
            print("HOOK!", epoch, summary)
            hook(epoch, summary)


if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            now = int(sys.argv[1])
        else:
            now = time.time()
        detect_upcoming_events(URL, WINDOW, on_event, now)
    except Exception as e:
        on_failure(str(e))
        raise
