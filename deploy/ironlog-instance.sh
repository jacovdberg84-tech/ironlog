#!/usr/bin/env bash
# deploy/ironlog-instance.sh — runs another company's IronLog on this server, kept apart from AML.
#
# Every instance uses the same code (/data/ironlog) but has its own folder
# under /data/ironlog-instances/<name>/ holding its own database, uploads,
# backups, secrets and settings file (instance.env). It is started by PM2 as
# ironlog-<name>, listening only on this machine (127.0.0.1); Cloudflare
# Tunnel publishes it on its own web address. It never reads AML's api/.env
# (api/loadEnv.js), never issues licences, and has no AML keys or mailboxes.
#
#   provision <name> <port> <time zone> [web address] [admin username] [admin full name]
#       Creates the folder and settings (once; secrets are never replaced),
#       starts or restarts the instance, and adds its first administrator if
#       it has none yet (prints a one-time setup code).
#   restart-all   Restarts every instance on the new code (run by deploy.yml).
#   list          Shows the instances and whether they answer.
set -euo pipefail

CODE_DIR="${IRONLOG_CODE_DIR:-/data/ironlog}"
BASE_DIR="${IRONLOG_INSTANCES_DIR:-/data/ironlog-instances}"
AML_PORT=3001

say() { printf '%s\n' "$*"; }
fail() { printf 'Stopped: %s\n' "$*" >&2; exit 1; }

# PM2 gets a clean environment, so the deploy's own keys (AML's AI key, GitHub token) never reach an instance.
pm2_clean() { env -i PATH="$PATH" HOME="${HOME:-/root}" PM2_HOME="${PM2_HOME:-${HOME:-/root}/.pm2}" pm2 "$@"; }

setting() { sed -n "s/^$2=//p" "$1" | head -n 1; }

healthy() {
  local port="$1" i
  for i in 1 2 3 4 5 6 7 8; do
    if curl -fsS "http://127.0.0.1:${port}/health" > /dev/null 2>&1; then return 0; fi
    sleep 4
  done
  return 1
}

start_or_restart() {
  local name="$1" dir="$BASE_DIR/$1"
  if pm2_clean jlist 2>/dev/null | grep -q "\"name\":\"ironlog-$name\""; then
    pm2_clean restart "ironlog-$name" > /dev/null
  else
    env -i PATH="$PATH" HOME="${HOME:-/root}" PM2_HOME="${PM2_HOME:-${HOME:-/root}/.pm2}" \
      IRONLOG_ENV_FILE="$dir/instance.env" \
      pm2 start "$CODE_DIR/api/index.js" --name "ironlog-$name" --cwd "$dir" --time > /dev/null
  fi
}

provision() {
  local name="${1:-}" port="${2:-}" tz="${3:-}" host="${4:-}" admin="${5:-}" admin_name="${6:-}"
  [[ "$name" =~ ^[a-z][a-z0-9-]{1,30}$ ]] || fail "instance name must be lower-case letters, numbers or dashes (e.g. pilot)."
  [[ "$port" =~ ^31[0-9][0-9]$ ]] || fail "port must be between 3100 and 3199."
  [ "$port" != "$AML_PORT" ] || fail "port $AML_PORT belongs to AML."
  [[ "$tz" =~ ^[A-Za-z_]+(/[A-Za-z0-9_+-]+)*$ ]] || fail "time zone must look like Africa/Maputo."
  [ -z "$host" ] || [[ "$host" =~ ^[a-z0-9.-]+$ ]] || fail "web address must look like pilot.ironlogafrica.com."
  [ -d "$CODE_DIR/api/node_modules" ] || fail "IronLog code not found in $CODE_DIR."

  local dir="$BASE_DIR/$name" env_file="$BASE_DIR/$name/instance.env" other
  for other in "$BASE_DIR"/*/instance.env; do
    [ -f "$other" ] && [ "$other" != "$env_file" ] || continue
    [ "$(setting "$other" PORT)" != "$port" ] || fail "port $port is already used by $(basename "$(dirname "$other")")."
  done

  umask 077
  mkdir -p "$BASE_DIR" "$dir/db/backups" "$dir/data" "$dir/secrets"
  chmod 700 "$dir" "$dir/secrets"

  if [ -f "$env_file" ]; then
    say "Settings already exist for $name; keeping them (secrets are never replaced)."
    [ "$(setting "$env_file" PORT)" = "$port" ] || say "Note: $name keeps its port $(setting "$env_file" PORT)."
  else
    {
      printf '# IronLog company instance "%s", created %s. Private: keep chmod 600.\n' "$name" "$(date -u +%Y-%m-%d)"
      printf 'PORT=%s\nHOST=127.0.0.1\n' "$port"
      printf 'DB_PATH=%s/db/ironlog.db\n' "$dir"
      printf 'IRONLOG_DATA_DIR=%s/data\n' "$dir"
      printf 'IRONLOG_DB_BACKUP_DIR=%s/db/backups\n' "$dir"
      printf 'IRONLOG_SECRETS_DIR=%s/secrets\n' "$dir"
      printf 'IRONLOG_LICENCE_ISSUER=0\n'
      printf 'IRONLOG_TIMEZONE=%s\n' "$tz"
      printf 'IRONLOG_AUTH_REQUIRED=true\n'
      printf 'IRONLOG_AUTH_SECRET=%s\n' "$(openssl rand -base64 48 | tr -d '\n')"
      printf 'JUNE_CONNECTOR_ENCRYPTION_SECRET=%s\n' "$(openssl rand -base64 48 | tr -d '\n')"
      [ -z "$host" ] || printf 'IRONLOG_PUBLIC_BASE_URL=https://%s\n' "$host"
    } > "$env_file"
    say "Created settings for $name in $env_file."
  fi
  chmod 600 "$env_file"

  start_or_restart "$name"
  port="$(setting "$env_file" PORT)"
  if healthy "$port"; then say "ironlog-$name is running on 127.0.0.1:$port."; else fail "ironlog-$name did not answer on port $port. See: pm2 logs ironlog-$name"; fi
  pm2_clean save > /dev/null

  if [ -n "$admin" ]; then
    (cd "$CODE_DIR/api" && env -i PATH="$PATH" IRONLOG_ENV_FILE="$env_file" node scripts/first-admin.js "$admin" "$admin_name")
  fi
}

restart_all() {
  local env_file name port failed=0
  for env_file in "$BASE_DIR"/*/instance.env; do
    [ -f "$env_file" ] || continue
    name="$(basename "$(dirname "$env_file")")"
    port="$(setting "$env_file" PORT)"
    start_or_restart "$name"
    if healthy "$port"; then say "ironlog-$name restarted and healthy."; else say "ironlog-$name did not answer on port $port."; failed=1; fi
  done
  pm2_clean save > /dev/null 2>&1 || true
  return "$failed"
}

list() {
  local env_file name port
  for env_file in "$BASE_DIR"/*/instance.env; do
    [ -f "$env_file" ] || continue
    name="$(basename "$(dirname "$env_file")")"
    port="$(setting "$env_file" PORT)"
    if curl -fsS "http://127.0.0.1:${port}/health" > /dev/null 2>&1; then say "$name  port $port  running"; else say "$name  port $port  NOT answering"; fi
  done
}

case "${1:-}" in
  provision) shift; provision "$@" ;;
  restart-all) restart_all ;;
  list) list ;;
  *) say "Usage: $0 provision <name> <port> <time zone> [web address] [admin username] [admin full name] | restart-all | list"; exit 2 ;;
esac
