import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()
        # example chp url for Coca Cola 1.5L barcode 7290112490189 or some popular item
        # wait, let's just search for milk or sugar 7290000000000 probably isn't a thing, but 7290000000000
        await page.goto("https://chp.co.il/%D7%98%D7%99%D7%A8%D7%AA%20%D7%9B%D7%A8%D7%9E%D7%9C/9000/2100/7290112490189/0")
        await asyncio.sleep(2)
        
        elements = await page.query_selector_all("h1, h2, h3, h4, h5, table")
        for el in elements:
            tag = await el.evaluate("e => e.tagName.toLowerCase()")
            if tag != "table":
                text = await el.inner_text()
                print(f"HEADER {tag}: {text}")
            else:
                headers = await el.query_selector_all("th")
                texts = [await h.inner_text() for h in headers]
                print(f"TABLE with headers: {texts}")
        
        await browser.close()

asyncio.run(main())
