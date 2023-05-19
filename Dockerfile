FROM node as build
WORKDIR /app
COPY ./package.json ./package-lock.json /app/
RUN npm ci
COPY . /app
COPY ./src/config.js.template /app/src/config.js
RUN npm run build

FROM nginx:latest
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY entrypoint.sh /entrypoint.sh

RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]