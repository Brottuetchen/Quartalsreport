#!/bin/sh
# entrypoint.sh – Restore any persisted OTA update, then start the app.
# The webapp_override directory lives on the persistent data volume so updates
# survive container restarts even though the base image is unchanged.

OVERRIDE_DIR="/app/data/webapp_override"
WEBAPP_DIR="/app/webapp"

if [ -d "$OVERRIDE_DIR" ] && [ "$(ls -A "$OVERRIDE_DIR" 2>/dev/null)" ]; then
  echo "[entrypoint] Restoring OTA update from $OVERRIDE_DIR ..."
  cp -r "$OVERRIDE_DIR/." "$WEBAPP_DIR/"
  echo "[entrypoint] Done."
fi

exec uvicorn webapp.server:app \
  --host 0.0.0.0 \
  --port 9999 \
  --reload \
  --reload-dir "$WEBAPP_DIR"
