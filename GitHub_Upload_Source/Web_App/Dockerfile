FROM ghcr.io/puppeteer/puppeteer:21.5.0

USER root
WORKDIR /usr/src/app

# Copy package files
COPY package*.json ./

# Install dependencies (Skip chromium download since we use the one in base image)
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true
RUN npm install

# Copy app source
COPY . .

# Expose port
EXPOSE 3000

# Start command
CMD [ "node", "server.js" ]
