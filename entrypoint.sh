#!/bin/sh

envsubst < /usr/share/nginx/html/manifest.template.xml > /usr/share/nginx/html/manifest.xml
envsubst < /usr/share/nginx/html/config.template.js > /usr/share/nginx/html/config.js
envsubst < /etc/nginx/conf.d/default.template.conf > /etc/nginx/conf.d/default.conf

nginx -g "daemon off;"
