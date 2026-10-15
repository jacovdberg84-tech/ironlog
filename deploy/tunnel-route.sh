#!/usr/bin/env bash
# deploy/tunnel-route.sh — publishes a company instance through the server's Cloudflare Tunnel.
#
# The "ironlog" tunnel is locally managed: its routes live in cloudflared's
# config file on this server, not in the Cloudflare dashboard. This adds one
#   - hostname: <web address>
#     service: http://localhost:<port>
# rule just above the final catch-all rule, after copying the file to
# <config>.bak-<date>. The new file is checked with "cloudflared ingress
# validate" before cloudflared restarts; if the check or restart fails the
# copy is put back. Then it points the web address at the tunnel in DNS.
#
#   add <web address> <port>
set -euo pipefail

say() { printf '%s\n' "$*"; }
fail() { printf 'Stopped: %s\n' "$*" >&2; exit 1; }

find_config() {
  local unit_cfg
  unit_cfg="$(systemctl cat cloudflared 2>/dev/null | grep -oE -- '--config[= ][^ ]+' | head -n 1 | sed -E 's/--config[= ]//' || true)"
  for f in "$unit_cfg" /etc/cloudflared/config.yml /etc/cloudflared/config.yaml "${HOME:-/root}/.cloudflared/config.yml" "${HOME:-/root}/.cloudflared/config.yaml"; do
    [ -n "$f" ] && [ -f "$f" ] && { printf '%s\n' "$f"; return 0; }
  done
  return 1
}

add() {
  local host="${1:-}" port="${2:-}"
  [[ "$host" =~ ^[a-z0-9-]+(\.[a-z0-9-]+)+$ ]] || fail "web address must look like pilot.ironlogafrica.com."
  [[ "$port" =~ ^31[0-9][0-9]$ ]] || fail "port must be between 3100 and 3199 (AML's 3001 is never changed)."
  command -v cloudflared > /dev/null || fail "cloudflared is not installed on this server."
  local cfg; cfg="$(find_config)" || fail "cloudflared config file not found."
  say "Config file: $cfg"

  local backup="$cfg.bak-$(date +%Y%m%d-%H%M%S)"
  cp -p "$cfg" "$backup"
  say "Copy saved: $backup"

  # Prints "added", "present" or an error; edits the file in place.
  local result
  result="$(node "$(dirname "$0")/tunnel-route.cjs" "$cfg" "$host" "$port")" || { cp -p "$backup" "$cfg"; fail "could not add the route; file left as it was."; }
  say "Route: $result"

  if ! cloudflared tunnel --config "$cfg" ingress validate; then
    cp -p "$backup" "$cfg"
    fail "cloudflared did not accept the new file; the copy was put back. Nothing restarted."
  fi

  if [ "$result" = "added" ]; then
    systemctl restart cloudflared || { cp -p "$backup" "$cfg"; systemctl restart cloudflared || true; fail "cloudflared did not restart; the copy was put back."; }
    sleep 5
    if ! systemctl is-active --quiet cloudflared; then
      cp -p "$backup" "$cfg"; systemctl restart cloudflared || true
      fail "cloudflared stopped after the change; the copy was put back and it was restarted."
    fi
    say "cloudflared restarted and running."
  fi

  local tunnel; tunnel="$(sed -n -E 's/^tunnel:[[:space:]]*"?([^"[:space:]]+)"?.*/\1/p' "$cfg" | head -n 1)"
  if cloudflared tunnel route dns "$tunnel" "$host" 2>&1; then
    say "DNS: $host points at the tunnel."
  else
    say "DNS could not be set from the server. In Cloudflare DNS add: CNAME $(printf '%s' "$host" | cut -d. -f1) -> ${tunnel}.cfargotunnel.com (Proxied)."
  fi
  say "== rules now"; cloudflared tunnel --config "$cfg" ingress validate > /dev/null && grep -nE '^\s*-?\s*(hostname|service):' "$cfg"
}

case "${1:-}" in
  add) shift; add "$@" ;;
  *) say "Usage: $0 add <web address> <port>"; exit 2 ;;
esac
