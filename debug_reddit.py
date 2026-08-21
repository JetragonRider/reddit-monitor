#!/usr/bin/env python3
"""Debug: Dump full HTML content from RSS entry."""
import urllib.request, ssl, xml.etree.ElementTree as ET, re

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

url = "https://www.reddit.com/r/Palworld/hot/.rss?limit=2"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 Chrome/131.0.0.0"})
with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
    content = resp.read()

root = ET.fromstring(content)
atom_ns = "http://www.w3.org/2005/Atom"
entries = root.findall(f"{{{atom_ns}}}entry")

for i, entry in enumerate(entries):
    print(f"\n=== Entry {i+1} ===")
    content_el = entry.find(f"{{{atom_ns}}}content")
    if content_el is not None and content_el.text:
        html = content_el.text
        print(f"Content HTML ({len(html)} chars):")
        print(html[:3000])
        print("\n--- Parsing ---")
        
        # Extract score - Reddit RSS uses <span class="score">123</span> or data attributes
        score_match = re.findall(r'(?:score|upvote|point|liked)[^<]*?(\d+)', html, re.IGNORECASE)
        print(f"Score matches: {score_match}")
        
        # Extract comment count - <a href=".../comments/...">123 comments</a>
        comment_matches = re.findall(r'(\d+)\s*(?:comment|reply|response)', html, re.IGNORECASE)
        print(f"Comment matches: {comment_matches}")
        
        # Find all numbers in the HTML
        numbers = re.findall(r'>(\d+)<', html)
        print(f"All numbers in tags: {numbers}")
        
        # Find all links
        links = re.findall(r'href="([^"]*)"', html)
        for l in links:
            if "comment" in l.lower():
                print(f"Comment link: {l}")
        
        # Find span content
        spans = re.findall(r'<span[^>]*>([^<]+)</span>', html)
        print(f"Spans: {spans}")
        
        # Find all text content between tags
        texts = re.findall(r'>([^<]+)<', html)
        texts = [t.strip() for t in texts if t.strip()]
        print(f"All text content: {texts}")
