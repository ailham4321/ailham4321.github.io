#!/usr/bin/env bash
set -euo pipefail
URL=${1:?"Usage: fetch_cv_pdf.sh <onedrive_share_url>"}

UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
mkdir -p documents

resolve_url() {
  local url="$1"
  local resolved
  resolved=$(curl -A "$UA" -sS -L -o /dev/null -w '%{url_effective}' "$url" || true)
  echo "$resolved"
}

get_direct_from_url() {
  local url="$1"
  local resid authkey
  resid=$(printf '%s' "$url" | sed -n 's/.*[?&]resid=\([^&]*\).*/\1/p')
  authkey=$(printf '%s' "$url" | sed -n 's/.*[?&]authkey=\([^&]*\).*/\1/p')
  if [[ -n "$resid" && -n "$authkey" ]]; then
    echo "https://onedrive.live.com/download?resid=${resid}&authkey=${authkey}"
  else
    echo ""
  fi
}

is_docx() {
  local f="$1"
  if [[ ! -f "$f" ]]; then return 1; fi
  if command -v file >/dev/null 2>&1; then
    file "$f" | grep -qi "zip archive" && return 0 || return 1
  else
    head -c 4 "$f" | grep -q "PK" && return 0 || return 1
  fi
}

fetch_docx() {
  local url="$1"
  curl -A "$UA" -L --fail -o cv.docx "$url" || return 1
  is_docx cv.docx
}

# 1) Try direct/download link
resolved=$(resolve_url "$URL")
direct=$(get_direct_from_url "$resolved")
if [[ -n "$direct" ]]; then
  echo "Using derived direct download URL"
  if fetch_docx "$direct"; then
    echo "Fetched DOCX via direct URL"
  else
    echo "Direct download failed"
  fi
fi

# 2) Try original URL if not yet fetched
if ! is_docx cv.docx; then
  echo "Trying original URL"
  fetch_docx "$URL" || true
fi

# 3) Try scraping HTML for download link via Python
if ! is_docx cv.docx; then
  echo "Resolving via Python helper"
  python3 - "$resolved" <<'PY' > .direct_url 2>/dev/null || true
import sys, re, urllib.request, urllib.parse
UA = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
url = sys.argv[1]

def build_direct(u: str):
    parsed = urllib.parse.urlparse(u)
    q = urllib.parse.parse_qs(parsed.query)
    resid = q.get('resid', [None])[0]
    auth = q.get('authkey', [None])[0]
    if resid and auth:
        return f"https://onedrive.live.com/download?resid={urllib.parse.quote(resid)}&authkey={urllib.parse.quote(auth)}"
    return None

try:
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with urllib.request.urlopen(req, timeout=20) as resp:
        final_url = resp.geturl()
        direct = build_direct(final_url)
        if direct:
            print(direct)
        else:
            body = resp.read(256*1024).decode('utf-8', errors='ignore')
            m = re.search(r'download\?resid=([^&"\\]+)&authkey=([^&"\\]+)', body)
            if m:
                resid, auth = m.group(1), m.group(2)
                print(f"https://onedrive.live.com/download?resid={resid}&authkey={auth}")
except Exception:
    pass
PY
  if [[ -s .direct_url ]]; then
    direct=$(cat .direct_url)
    echo "Retrying via scraped direct URL"
    fetch_docx "$direct" || true
  fi
fi

# 3.5) Try OneDrive Public Shares API (works with public share links)
if ! is_docx cv.docx; then
  echo "Trying OneDrive public shares API"
  # Base64url encode the original share URL
  share_id=$(python3 - <<'PY'
import sys, base64
u=sys.stdin.read().strip().encode('utf-8')
print('u!' + base64.urlsafe_b64encode(u).decode('ascii').rstrip('='))
PY
  <<<"$URL")
  curl -A "$UA" -L --fail -o cv.docx \
    "https://api.onedrive.com/v1.0/shares/${share_id}/driveItem/content" || true
fi

# 4) Convert if we have a DOCX
if is_docx cv.docx; then
  echo "Converting with LibreOffice"
  sudo apt-get update
  sudo apt-get install -y libreoffice
  soffice --headless --convert-to pdf --outdir documents cv.docx
  out=$(ls -1 documents/*.pdf 2>/dev/null | head -n 1 || true)
  if [[ -z "$out" ]]; then
    echo "ERROR: No PDF produced by LibreOffice" >&2
    exit 1
  fi
  mv -f "$out" documents/Abdillah-Ilham-CV-Indonesia.pdf
  exit 0
fi

# 5) Fallback to Office export service
echo "Falling back to Office export service"
encoded=$(python3 -c "import sys, urllib.parse; print(urllib.parse.quote(sys.argv[1], safe=''))" "${direct:-$URL}")
set +e
curl -A "$UA" -L --fail --retry 3 --retry-delay 2 -sS \
  -o documents/Abdillah-Ilham-CV-Indonesia.pdf \
  "https://export.word.officeapps.live.com/export?format=pdf&url=${encoded}"
rc=$?
set -e
if [[ $rc -ne 0 || ! -s documents/Abdillah-Ilham-CV-Indonesia.pdf ]]; then
  echo "ERROR: Unable to obtain PDF from OneDrive link (curl exit $rc)." >&2
  exit 1
fi
