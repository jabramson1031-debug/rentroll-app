/** @type {import('next').NextConfig} */
const nextConfig = {
  // For Electron production build, export as static site
  output: process.env.ELECTRON_BUILD ? 'export' : undefined,
  trailingSlash: true,
  images: {
    unoptimized: true,
  },
}

module.exports = nextConfig
