#!/bin/sh

set -e

if [ "$UID" -eq 0 ] ; then
    echo "$0 is not meant to be run as root" >&2
    exit 1
fi

calcheck="$HOME/Library/LaunchAgents/com.calcheck.timer.plist"
dir=$(pwd)
stdout="$dir/log.out"
stderr="$dir/log.err"

if [ -e $calcheck ]; then
    echo "Existing plist file found. Unloading $calcheck"
    launchctl unload $calcheck
fi

touch $stdout
touch $stderr
echo "writing $calcheck:"
sudo tee "$calcheck" << EOF | sed -e 's/^/    /'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.calcheck.timer</string>

  <key>ProgramArguments</key>
  <array>
        <string>/usr/bin/python3</string>
        <string>$dir/calcheck.py</string>
  </array>

  <key>StartInterval</key>
  <integer>120</integer>

  <key>RunAtLoad</key>
  <true/>

  <key>StandardErrorPath</key>
  <string>$dir/log.err</string>
  <key>StandardOutPath</key>
  <string>$dir/log.out</string>

</dict>
</plist>

EOF
echo ""

set -x
launchctl load $calcheck
