import re
import sys

if __name__ == "__main__":
    tzids = {}
    with open(sys.argv[1]) as f:
        lines = iter(f)
        while True:
            try:
                display_name = next(lines).strip()
                match = re.match(r"\(UTC([+-])([0-9][0-9]):([0-9][0-9])\)", display_name)
                if match is None:
                    # there should only be one non-match
                    assert display_name.startswith("(UTC)")
                    offset = 0
                else:
                    # hour
                    offset = 60*60 * int(match[2])
                    # min
                    offset += 60 * int(match[3])
                    # +/-
                    offset *= 1 if match[1] == '+' else -1
                tzid = next(lines).strip()
                _ = next(lines)
                tzids[tzid] = offset
            except StopIteration:
                break

    # generate some python code
    print("tz_offsets = {")
    for tzid, offset in tzids.items():
        print("    \"%s\": %d,"%(tzid, offset))
    print("}")
