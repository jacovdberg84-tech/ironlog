#!/usr/bin/env bash
set -euo pipefail

DRY_RUN=0
YES=0
PM2_APP="${PM2_APP:-ironlog-api}"
HEALTH_URL="${IRONLOG_HEALTH_URL:-http://127.0.0.1:3001/health}"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    --yes|-y)
      YES=1
      shift
      ;;
    --pm2-app)
      PM2_APP="${2:-}"
      shift 2
      ;;
    --help|-h)
      cat <<'EOF'
Usage: restore-latest-backup.sh [--dry-run] [--yes] [--pm2-app <name>]

Restores SQLite from the newest auto-backup in db/backups, validates integrity,
restarts PM2 app, and checks health endpoint.

Options:
  --dry-run       Print planned actions without making changes
  --yes, -y       Skip confirmation prompt
  --pm2-app NAME  PM2 process name (default: ironlog-api)
  --help, -h      Show this help
EOF
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      exit 1
      ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
API_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DB_DIR="${IRONLOG_DB_DIR:-$API_DIR/db}"
BACKUP_DIR="${IRONLOG_DB_BACKUP_DIR:-$DB_DIR/backups}"
RESTORE_DIR="$DB_DIR/manual_restore"

DB_FILE="$DB_DIR/ironlog.db"
ROOT_DB_FILE="$API_DIR/ironlog.db"
WAL_FILE="$DB_DIR/ironlog.db-wal"
SHM_FILE="$DB_DIR/ironlog.db-shm"

LATEST_BACKUP="$(ls -1t "$BACKUP_DIR"/ironlog_*.db 2>/dev/null | head -n 1 || true)"
if [[ -z "$LATEST_BACKUP" ]]; then
  echo "No backup files found in $BACKUP_DIR" >&2
  exit 1
fi

TS="$(date +%Y%m%d_%H%M%S)"

echo "Restore source: $LATEST_BACKUP"
echo "Target DB: $DB_FILE"
echo "PM2 app: $PM2_APP"
echo "Health URL: $HEALTH_URL"

run() {
  if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "[dry-run] $*"
  else
    eval "$@"
  fi
}

if [[ "$DRY_RUN" -eq 0 && "$YES" -eq 0 ]]; then
  read -r -p "Proceed with restore? This will replace the active DB. [y/N]: " answer
  if [[ ! "$answer" =~ ^[Yy]$ ]]; then
    echo "Cancelled."
    exit 1
  fi
fi

run "mkdir -p '$RESTORE_DIR'"
run "cp -f '$DB_FILE' '$RESTORE_DIR/ironlog.db.before_restore.$TS'"
run "if [[ -f '$WAL_FILE' ]]; then cp -f '$WAL_FILE' '$RESTORE_DIR/ironlog.db-wal.before_restore.$TS'; fi"
run "if [[ -f '$SHM_FILE' ]]; then cp -f '$SHM_FILE' '$RESTORE_DIR/ironlog.db-shm.before_restore.$TS'; fi"

run "cp -f '$LATEST_BACKUP' '$DB_FILE'"
run "rm -f '$WAL_FILE' '$SHM_FILE'"

run "sqlite3 '$DB_FILE' 'PRAGMA quick_check;'"
run "cp -f '$DB_FILE' '$ROOT_DB_FILE'"
run "pm2 restart '$PM2_APP' --update-env"
run "sleep 2"
run "curl -fsS '$HEALTH_URL' >/dev/null"

echo "Restore completed successfully."
