FROM node as build
WORKDIR /app
COPY ./package.json /app/
RUN npm install
COPY . /app
RUN npm run build

FROM node as build-client
WORKDIR /app/client
COPY ./client/package.json /app/client/
RUN npm install
COPY ./client /app/client
RUN npm run build

FROM nginx
COPY --from=build /app/dist /usr/share/nginx/html
COPY --from=build-client /app/client/dist /usr/share/nginx/html/client
RUN rm /etc/nginx/conf.d/default.conf
COPY nginx/nginx.conf /etc/nginx/conf.d/
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]