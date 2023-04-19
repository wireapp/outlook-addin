FROM node as build
WORKDIR /app
COPY ./package.json ./package-lock.json /app/
COPY ./src/config.template.js /app/config.js
RUN npm ci
COPY . /app
RUN npm run build

FROM nginx:latest
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.template.conf /etc/nginx/conf.d/default.template.conf
COPY entrypoint.sh /entrypoint.sh

RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]