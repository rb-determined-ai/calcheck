#!/bin/sh

set -e

if [ "$UID" -eq 0 ] ; then
    echo "$0 is not meant to be run as root" >&2
    exit 1
fi

timer="/etc/systemd/user/calcheck.timer"
echo "writing $timer:"
sudo tee "$timer" << EOF | sed -e 's/^/    /'
[Unit]
Description=calcheck.py timer

[Timer]
# assume events start on 5-minute increments, and catch them 2 minutes before
OnCalendar=*-*-* *:3,8,13,18,23,28,33,38,43,48,53,58:0
Unit=calcheck.service

[Install]
WantedBy=timers.target
EOF
echo ""

service="/etc/systemd/user/calcheck.service"
echo "writing $service:"
sudo tee "$service" << EOF | sed -e 's/^/    /'
[Unit]
Description=calcheck service

[Service]
Type=oneshot
ExecStart=$PWD/calcheck.py
TimeoutSec=60
EOF
echo ""

set -x
systemctl --user daemon-reload
systemctl --user enable --now calcheck.timer
