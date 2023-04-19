#!/bin/sh

envsubst < /usr/share/nginx/html/manifest.xml.template > /usr/share/nginx/html/manifest.xml
envsubst < /usr/share/nginx/html/config.js.template > /usr/share/nginx/html/config.js
envsubst < /etc/nginx/conf.d/default.conf.template > /etc/nginx/conf.d/default.conf

nginx -g "daemon off;"
