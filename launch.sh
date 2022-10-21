#!/bin/sh

set -e
dir="$(dirname $0)"
log="$dir/output.log"
script="$dir/calcheck.py"
echo "$script"
echo "$log"
date >> $log
python3 $script >> $log
echo "--------------------------" >> $log