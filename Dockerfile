FROM node:boron

# Create app directory
RUN mkdir /app
WORKDIR /app

COPY package.json /app/package.json
COPY index.js /app/index.js
COPY google.js /app/google.js
COPY client_secret.json /app/client_secret.json

RUN npm install
EXPOSE 3000
CMD ["node", "index.js"]
