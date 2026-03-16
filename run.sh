#!/bin/bash
cd "$(dirname "$0")"
echo ""
echo " delaware Timesheet Automator"
echo " ─────────────────────────────"
echo ""
python3 timesheet.py
if [ $? -ne 0 ]; then
  echo ""
  echo " Something went wrong. See error above."
  read -p " Press Enter to continue..."
fi
