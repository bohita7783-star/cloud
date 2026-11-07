#!/data/data/com.termux/files/usr/bin/bash
cd "$(dirname "$0")" || exit 1
python homebox_convert.py "$@"
status=$?
echo ""
echo "終了コード: $status"
if [ -t 0 ]; then read -r -p "Enterを押すと戻ります..." unused; fi
exit "$status"
