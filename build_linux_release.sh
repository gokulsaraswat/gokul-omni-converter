#!/usr/bin/env bash
# Build GO Omni Pro for Linux and package it as a .deb.
#
# NOTE: not executed during development (authored on macOS, where dpkg-deb is
# unavailable). Verify on a Debian/Ubuntu machine before shipping.
#
# Requires: python3, pyinstaller, dpkg-deb.
set -euo pipefail
cd "$(dirname "$0")"

VERSION="2.4.0"
PKG="go-omni-pro"
STAGE="dist/${PKG}_${VERSION}_amd64"

echo "==> Generating icons"
python3 installer/generate_icons.py

echo "==> Running release validation"
python3 release_validation.py
python3 app.py --validate-release

echo "==> Building with PyInstaller"
pyinstaller --clean --noconfirm installer/GOOmniPro.spec

command -v dpkg-deb >/dev/null 2>&1 || {
  echo "dpkg-deb not found — the PyInstaller build is in dist/, but no .deb was produced." >&2
  exit 1
}

echo "==> Staging .deb tree"
rm -rf "$STAGE"
install -d "$STAGE/DEBIAN" \
           "$STAGE/opt/${PKG}" \
           "$STAGE/usr/share/applications" \
           "$STAGE/usr/share/icons/hicolor/256x256/apps" \
           "$STAGE/usr/share/icons/hicolor/48x48/apps"

cp -r "dist/GO Omni Pro/." "$STAGE/opt/${PKG}/"
cp installer/go-omni-pro.desktop "$STAGE/usr/share/applications/${PKG}.desktop"
cp assets/logo/go_omni_pro_256.png "$STAGE/usr/share/icons/hicolor/256x256/apps/${PKG}.png"
cp assets/logo/go_omni_pro_48.png  "$STAGE/usr/share/icons/hicolor/48x48/apps/${PKG}.png"

cat > "$STAGE/DEBIAN/control" <<EOF
Package: ${PKG}
Version: ${VERSION}
Section: utils
Priority: optional
Architecture: amd64
Maintainer: Gokul Saraswat <gokulsaraswat07@gmail.com>
Homepage: https://gokulsaraswat.com
Depends: libc6, libtk8.6 | libtk9.0
Recommends: tesseract-ocr, libreoffice
Description: GO Omni Pro - local-first document conversion toolkit
 Convert PDFs, documents, spreadsheets, presentations, images, HTML and
 Markdown entirely offline. Includes PDF tools, OCR, a page organizer,
 archive tools, batch renaming, and automation presets.
EOF

# Refresh the icon and desktop caches so the launcher appears immediately.
cat > "$STAGE/DEBIAN/postinst" <<'EOF'
#!/bin/sh
set -e
if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database -q /usr/share/applications || true
fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then
  gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor || true
fi
EOF
chmod 0755 "$STAGE/DEBIAN/postinst"

echo "==> Building .deb"
dpkg-deb --build --root-owner-group "$STAGE"
echo "Done: ${STAGE}.deb"
echo "Install with: sudo dpkg -i ${STAGE}.deb"
