FROM node:20-alpine

WORKDIR /app

# Install dependencies
COPY package*.json ./
RUN npm install

# Install client dependencies and build
COPY client/package*.json ./client/
RUN cd client && npm install

# Copy source
COPY . .

# Build React client
RUN cd client && npm run build

EXPOSE 3001
ENV NODE_ENV=production

CMD ["node", "server/index.js"]
