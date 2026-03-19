# Use lightweight Node image
FROM node:18-alpine

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --production

# Copy app code
COPY . .

# Create output directory
# RUN mkdir -p output
RUN apk add --no-cache tzdata \
 && cp /usr/share/zoneinfo/Asia/Kolkata /etc/localtime \
 && echo "Asia/Kolkata" > /etc/timezone

# Set environment variables (optional fallback)
ENV NODE_ENV=production

# Start app
CMD ["node", "index.js"]

# Health check to ensure the container is running properly
HEALTHCHECK CMD node -e "console.log('healthy')" || exit 1