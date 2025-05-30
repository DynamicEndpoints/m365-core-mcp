# Generated by https://smithery.ai. See: https://smithery.ai/docs/config#dockerfile
FROM node:lts-alpine

WORKDIR /app

# Install dependencies
COPY package*.json ./
RUN npm install --ignore-scripts

# Copy source code
COPY . .

# Build the project
RUN npm run build

# Expose standard port if needed (optional)
# EXPOSE 3000

# Start the MCP server
CMD ["npm", "start"]
