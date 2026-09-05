import { mkdirSync } from 'node:fs'
import { chromium } from 'playwright'

mkdirSync('docs/screenshots', { recursive: true })

const browser = await chromium.launch({ headless: true })

async function shot(name, fn, opts = {}) {
  const context = await browser.newContext({
    viewport: opts.viewport ?? { width: 1440, height: 900 },
    deviceScaleFactor: 2,
    locale: 'it-IT',
  })
  await context.addInitScript((theme) => {
    localStorage.setItem('vba-desk-lang', 'it')
    localStorage.setItem('vba-desk-theme', theme)
  }, opts.theme ?? 'light')
  const page = await context.newPage()
  await page.goto('http://localhost:5175/', { waitUntil: 'networkidle' })
  await page.waitForTimeout(700)
  await fn(page)
  await page.screenshot({
    path: `docs/screenshots/${name}.png`,
    fullPage: Boolean(opts.fullPage),
  })
  await context.close()
}

await shot('01-home', async () => {})
await shot('01-home-dark', async () => {}, { theme: 'dark' })
await shot('02-catalogo-excel', async (page) => {
  await page.locator('.scope-tile[data-scope="excel"]').click()
  await page.waitForTimeout(700)
})
await shot('03-categoria-range', async (page) => {
  await page.goto('http://localhost:5175/#/cat/xl-range', { waitUntil: 'networkidle' })
  await page.waitForTimeout(700)
})
await shot('04-comando-msgbox', async (page) => {
  await page.goto('http://localhost:5175/#/c/msgbox', { waitUntil: 'networkidle' })
  await page.waitForTimeout(800)
})
await shot('05-comando-range', async (page) => {
  await page.goto('http://localhost:5175/#/c/xl-range-value', { waitUntil: 'networkidle' })
  await page.waitForTimeout(800)
})
await shot('06-mobile', async () => {}, { viewport: { width: 390, height: 844 } })

await browser.close()
console.log('screenshots saved to docs/screenshots')
