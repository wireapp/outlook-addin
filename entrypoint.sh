#!/usr/bin/env bash
set -euo pipefail

# envsubst coming with the nginx container unfortunately doesn't support
# -no-unset and -no-empty parameters, so make sure to keep this list in sync!

if [ -z ${ADDIN_HOST:-} ]; then echo "ADDIN_HOST is unset"; exit 1; fi
if [ -z ${API_HOST:-} ]; then echo "API_HOST is unset"; exit 1; fi
if [ -z ${AUTHORIZE_HOST:-} ]; then echo "AUTHORIZE_HOST is unset"; exit 1; fi
if [ -z ${CLIENT_ID:-} ]; then echo "CLIENT_ID is unset"; exit 1; fi

envsubst  < /usr/share/nginx/html/manifest.xml.template > /usr/share/nginx/html/manifest.xml
envsubst  < /usr/share/nginx/html/config.js.template > /usr/share/nginx/html/config.js

nginx -g "daemon off;"
