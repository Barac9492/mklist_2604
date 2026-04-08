import asyncio
from playwright.async_api import async_playwright
import re
import os

data_dir = os.path.join(os.path.dirname(__file__), 'data')
os.makedirs(data_dir, exist_ok=True)

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )
        page = await context.new_page()

        article_urls = []
        for page_num in range(1, 6): # Get 50 articles
            url = f"https://search.mk.co.kr/search?word=%EC%8B%A0%EC%84%A4%EB%B2%95%EC%9D%B8&docType=news&pn={page_num}"
            print(f"Fetching search page {page_num}...")
            await page.goto(url, wait_until="domcontentloaded")
            await page.wait_for_timeout(3000)
            
            # Find all links that contain '신설법인' in their text context
            links = await page.locator('a:has-text("신설법인")').evaluate_all(
                "elements => elements.map(el => el.href)"
            )
            # Filter valid news links
            links = [l for l in links if 'mk.co.kr/news/business/' in l]
            article_urls.extend(links)
        
        # Deduplicate
        article_urls = list(set(article_urls))
        print(f"Found {len(article_urls)} articles.")

        for url in article_urls:
            print(f"Processing {url}...")
            try:
                await page.goto(url, wait_until="domcontentloaded")
                await page.wait_for_timeout(2000)
            except Exception as e:
                print(f"  Could not load article page: {e}")
                continue
            
            # extract the date from the title to make a nice filename
            title = await page.title()
            
            article_id = url.split('/')[-1]
            
            m = re.search(r'(\d{1,2})월\s*(\d{1,2})일\s*~\s*(\d{1,2})월\s*(\d{1,2})일', title)
            if m:
                m1, d1, m2, d2 = [int(x) for x in m.groups()]
                filename = f"week_{m1:02d}{d1:02d}_{m2:02d}{d2:02d}.xls"
            else:
                filename = f"week_{article_id}.xls"
                
            out_path = os.path.join(data_dir, filename)
            if os.path.exists(out_path):
                print(f"  Already exists: {filename}")
                continue

            try:
                # Find download links. 
                download_loc = page.locator('a', has_text=re.compile(r'xls|xlsx|첨부|파일', re.IGNORECASE))
                count = await download_loc.count()
                
                # Check for instances of window.open() wrappers if not standard anchor
                if count > 0:
                    for i in range(count):
                        dl_elem = download_loc.nth(i)
                        # MK attaches files in a specific element often called "btn_file" or "file_down"
                        # We try to trigger the download event
                        try:
                            async with page.expect_download(timeout=5000) as download_info:
                                await dl_elem.click()
                            download = await download_info.value
                            await download.save_as(out_path)
                            print(f"  Saved {filename}")
                            break
                        except Exception as inner_e:
                            if i == count - 1:
                                print(f"  Failed download: {inner_e}")
                else:
                    print(f"  No download link found in {url}")
            except Exception as e:
                print(f"  Exception finding download link: {e}")
                
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
