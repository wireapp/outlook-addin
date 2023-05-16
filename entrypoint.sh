#!/usr/bin/env bash
set -euo pipefail

# envsubst coming with the nginx container unfortunately doesn't support
# -no-unset and -no-empty parameters, so make sure to keep this list in sync!

if [ -z ${BASE_URL:-} ]; then echo "BASE_URL is unset"; exit 1; fi
if [ -z ${WIRE_API_BASE_URL:-} ]; then echo "WIRE_API_BASE_URL is unset"; exit 1; fi
if [ -z ${WIRE_AUTHORIZATION_ENDPOINT:-} ]; then echo "WIRE_AUTHORIZATION_ENDPOINT is unset"; exit 1; fi
if [ -z ${CLIENT_ID:-} ]; then echo "CLIENT_ID is unset"; exit 1; fi

envsubst  < /usr/share/nginx/html/manifest.xml.template > /usr/share/nginx/html/manifest.xml
envsubst  < /usr/share/nginx/html/config.js.template > /usr/share/nginx/html/config.js

nginx -g "daemon off;"
