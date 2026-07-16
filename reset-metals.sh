#!/usr/bin/env bash
# ── Reset Metals / Bloop index — восстанавливает навигацию по исходникам ──
#
# Когда клик по имени класса не ведёт к его исходнику — значит индекс Metals
# (H2 .mv.db) или конфиг Bloop устарел. Скрипт убивает процессы, чистит кеши
# и генерирует конфигурацию заново.
#
# Использование:
#   ./reset-metals.sh          # только backend
#   ./reset-metals.sh --all    # включая sbt и frontend target

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$SCRIPT_DIR/backend"
FRONTEND_DIR="$SCRIPT_DIR/frontend"

echo "═══════════════════════════════════════════"
echo "  Reset Metals / Bloop index"
echo "═══════════════════════════════════════════"

# ── 1. Kill lingering processes ──

echo ""
echo "── Killing Metals / Bloop / sbt processes ──"

for PROC in "metals" "bloop" "sbt"; do
  if pkill -f "$PROC" 2>/dev/null; then
    echo "  ✅ killed $PROC"
  else
    echo "  ℹ️  no $PROC running"
  fi
done

# Wait for file locks to be released
sleep 1

# ── 2. Clean metals / bloop / bsp cache ──

echo ""
echo "── Cleaning Metals / Bloop / BSP cache ──"

clean_dir() {
  local dir="$1"
  local label="$2"
  if [ -d "$dir" ]; then
    rm -rf "$dir"
    echo "  ✅ removed $label"
  else
    echo "  ℹ️  $label — not found, skipping"
  fi
}

# Clean root-level Metals database
clean_dir "$SCRIPT_DIR/.metals" ".metals/ (root)"

# Clean backend-level caches
clean_dir "$BACKEND_DIR/.bloop"   "backend/.bloop/"
clean_dir "$BACKEND_DIR/.bsp"     "backend/.bsp/"
clean_dir "$BACKEND_DIR/.metals"  "backend/.metals/"

# Clean sbt target (auto-generated classfiles etc.)
clean_dir "$BACKEND_DIR/target"            "backend/target/"
clean_dir "$BACKEND_DIR/project/target"    "backend/project/target/"
clean_dir "$BACKEND_DIR/project/project"   "backend/project/project/"

# ── 3. Optional: clean frontend ──

if [ "${1:-}" = "--all" ]; then
  echo ""
  echo "── (--all) Cleaning frontend ──"
  clean_dir "$FRONTEND_DIR/node_modules" "frontend/node_modules/"
  clean_dir "$FRONTEND_DIR/dist"         "frontend/dist/"
  clean_dir "$FRONTEND/src-tauri/target" "frontend/src-tauri/target/"
fi

# ── 4. Generate Bloop config ──

echo ""
echo "── Generating Bloop config (sbt bloopInstall) ──"

cd "$BACKEND_DIR"
sbt "bloopInstall"
cd "$SCRIPT_DIR"

echo ""
echo "  ✅ Bloop config regenerated"

# ── 5. Optional: reinstall frontend ──

if [ "${1:-}" = "--all" ]; then
  echo ""
  echo "── (--all) Reinstalling frontend dependencies ──"
  cd "$FRONTEND_DIR"
  npm install
  cd "$SCRIPT_DIR"
  echo "  ✅ Frontend dependencies reinstalled"
fi

# ── 6. Done ──

echo ""
echo "═══════════════════════════════════════════"
echo "  ✅ Done! Restart Metals in VS Code:"
echo ""
echo "    1. Cmd+Shift+P → 'Metals: Restart server'"
echo "    2. Wait for import to finish"
echo "═══════════════════════════════════════════"
