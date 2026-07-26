const { join } = require('path');

/**
 * @type {import("puppeteer").Configuration}
 */
module.exports = {
  // Changes the cache location for Puppeteer to be within the project directory.
  // This ensures Render copies the Chromium binary into the runtime container.
  cacheDirectory: join(__dirname, '.cache', 'puppeteer'),
};
