FROM node as build
WORKDIR /app
COPY ./package.json /app/
RUN npm install
COPY . /app
COPY ./certs /certs
RUN npm run build

EXPOSE 4000
CMD ["node", "server.js"]