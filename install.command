#!/bin/bash
# Double-click in Finder to install f2p (script, Finder Quick Action, "f2p" terminal command).
set -e
cd "$(dirname "$0")"
DEST="$HOME/.f2p"
echo "Installing f2p into $DEST ..."
if ! /usr/bin/python3 -c "import sys" >/dev/null 2>&1; then
  echo "Python 3 is missing. macOS will offer to install the Command Line Tools; run this installer again afterwards."
  xcode-select --install || true
  exit 1
fi
mkdir -p "$DEST"
cp f2p.py run-quick-action.sh icon.png settings_default.ini "$DEST/"
chmod +x "$DEST/f2p.py" "$DEST/run-quick-action.sh"
[ -x "$DEST/venv/bin/python" ] || /usr/bin/python3 -m venv "$DEST/venv"
"$DEST/venv/bin/python" -m pip install --quiet --upgrade pip openpyxl
mkdir -p "$HOME/Library/Services"
rm -rf "$HOME/Library/Services/f2p.workflow"
cp -R f2p.workflow "$HOME/Library/Services/"
/System/Library/CoreServices/pbs -flush >/dev/null 2>&1 || true
LINE="alias f2p='\"$DEST/venv/bin/python\" \"$DEST/f2p.py\"'"
grep -qs "alias f2p=" "$HOME/.zshrc" || echo "$LINE" >> "$HOME/.zshrc"
echo
echo "Done. Right-click a folder or .wsp file in Finder > Quick Actions > 'f2p'."
echo "In a new terminal window you can also run:  f2p /path/to/folder"
