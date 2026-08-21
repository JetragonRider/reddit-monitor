#!/usr/bin/env python3
"""Debug: Test RSS + Playwright post page scraping."""
import urllib.request, ssl, xml.etree.ElementTree as ET, re, time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# 1. Get RSS posts
url = "https://www.reddit.com/r/Palworld/hot/.rss?limit=3"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 Chrome/131.0.0.0"})
with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
    content = resp.read()

root = ET.fromstring(content)
atom_ns = "http://www.w3.org/2005/Atom"
entries = root.findall(f"{{{atom_ns}}}entry")
print(f"RSS: {len(entries)} entries")

post_urls = []
for entry in entries:
    link = entry.find(f"{{{atom_ns}}}link")
    if link is not None:
        post_urls.append(link.get("href", ""))

print(f"Post URLs: {post_urls}")

# 2. Visit each post page with Playwright and get scores
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
    page = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"
    ).new_page()
    
    for i, post_url in enumerate(post_urls):
        old_url = post_url.replace("www.reddit.com", "old.reddit.com")
        print(f"\nPost {i+1}: {old_url[:60]}")
        
        try:
            page.goto(old_url, timeout=10000, wait_until="domcontentloaded")
            time.sleep(1)
            html = page.content()
            print(f"  Page size: {len(html)}")
            
            # Score
            score_match = re.search(r'class="score[^"]*"[^>]*>([\d,.]+)', html)
            if score_match:
                print(f"  Score: {score_match.group(1)}")
            else:
                print(f"  Score: NOT FOUND")
                # Try other patterns
                score_data = re.search(r'data-score="(\d+)"', html)
                if score_data:
                    print(f"  Score (data): {score_data.group(1)}")
            
            # Comments
            comment_match = re.search(r'class="comments[^"]*"[^>]*>([\d,.]+)\s*comment', html, re.IGNORECASE)
            if comment_match:
                print(f"  Comments: {comment_match.group(1)}")
            else:
                print(f"  Comments: NOT FOUND")
                # Try other patterns
                comment_data = re.search(r'data-num-comments="(\d+)"', html)
                if comment_data:
                    print(f"  Comments (data): {comment_data.group(1)}")
            
            # Print first 500 chars of page
            print(f"  Page preview: {html[:300]}")
            
        except Exception as e:
            print(f"  Error: {e}")
        
        time.sleep(1)
    
    browser.close()

print("\nDone!")
