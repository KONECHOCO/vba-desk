import { mkdirSync } from 'node:fs'
import puppeteer from 'puppeteer-core'

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
const BASE_URL = 'http://localhost:4174'
const OUT_DIR = 'docs/appstore-screenshots'
// CSS logical size of an iPhone 11 Pro Max / XS Max (matches the app's mobile
// breakpoint at 780px), rendered at deviceScaleFactor 3 to land exactly on
// Apple's required 1242x2688 raster for App Store screenshots.
const CSS_WIDTH = 414
const CSS_HEIGHT = 896
const SCALE = 3

mkdirSync(OUT_DIR, { recursive: true })

const browser = await puppeteer.launch({
  executablePath: CHROME_PATH,
  headless: true,
  defaultViewport: { width: CSS_WIDTH, height: CSS_HEIGHT, deviceScaleFactor: SCALE },
})

async function shot(name, path, { theme = 'light' } = {}) {
  const page = await browser.newPage()
  await page.setViewport({ width: CSS_WIDTH, height: CSS_HEIGHT, deviceScaleFactor: SCALE, isMobile: true, hasTouch: true })
  await page.evaluateOnNewDocument((t) => {
    localStorage.setItem('vba-desk-lang', 'it')
    localStorage.setItem('vba-desk-theme', t)
  }, theme)
  await page.goto(BASE_URL + path, { waitUntil: 'networkidle0' })
  await new Promise((r) => setTimeout(r, 500))
  await page.screenshot({ path: `${OUT_DIR}/${name}.png` })
  await page.close()
  console.log('saved', name)
}

await shot('01-home', '/')
await shot('02-catalogo-excel', '/#/cat/xl-range')
await shot('03-comando-msgbox', '/#/c/msgbox')
await shot('04-comando-range', '/#/c/xl-range-value')
await shot('05-home-dark', '/', { theme: 'dark' })

await browser.close()
console.log('done')
