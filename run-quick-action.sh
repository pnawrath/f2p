#!/bin/bash
# Called by the Finder Quick Action with the selected folder(s) / .wsp file(s) as arguments.
PY="$HOME/.f2p/venv/bin/python"
for f in "$@"; do
  if out=$("$PY" "$HOME/.f2p/f2p.py" "$f" 2>&1); then
    last=$(echo "$out" | tail -n1)
    name=$(basename "${last%%  (*}")
    info=$(echo "$last" | sed 's/.*  (//; s/)$//')
    osascript -e "display notification \"$info\" with title \"f2p\" subtitle \"$name\""
  elif echo "$out" | grep -q "^Cancelled"; then
    :
  else
    msg=$(echo "$out" | tail -n3 | tr '"' "'")
    osascript -e "display alert \"f2p failed\" message \"$msg\" as critical"
  fi
done
