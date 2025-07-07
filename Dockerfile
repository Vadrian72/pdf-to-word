FROM node:18-alpine

WORKDIR /app

COPY package*.json ./

RUN npm install

COPY . .

RUN mkdir -p uploads output

EXPOSE 3000

CMD ["npm", "start"]
