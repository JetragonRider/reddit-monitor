#!/usr/bin/env python3
"""Debug: Dump full RSS XML and check all attributes."""
import urllib.request, ssl, xml.etree.ElementTree as ET

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

url = "https://www.reddit.com/r/Palworld/hot/.rss?limit=3"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 Chrome/131.0.0.0"})
with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
    content = resp.read()

# Parse XML
root = ET.fromstring(content)

# Print all namespaces
print("=== Namespaces ===")
for prefix, uri in root.attrib.items():
    if "xmlns" in prefix.lower() or ":" in prefix:
        print(f"  {prefix} = {uri}")
# Also check for xmlns declarations
import re
xml_str = content.decode("utf-8")
ns_matches = re.findall(r'xmlns:?(\w*)="([^"]+)"', xml_str)
for prefix, uri in ns_matches:
    print(f"  xmlns:{prefix} = {uri}")

# Print all entries with ALL attributes
atom_ns = "http://www.w3.org/2005/Atom"
thr_ns = "http://purl.org/syndication/thread/1.0"

entries = root.findall(f"{{{atom_ns}}}entry")
print(f"\n=== {len(entries)} entries ===")

for i, entry in enumerate(entries):
    print(f"\n--- Entry {i+1} ---")
    # Print all child elements
    for child in entry:
        tag = child.tag
        text = (child.text or "")[:100]
        attrs = dict(child.attrib)
        print(f"  {tag}: text={text[:60]} attrs={attrs}")
    
    # Specifically look for link with replies
    links = entry.findall(f"{{{atom_ns}}}link")
    for link in links:
        href = link.get("href", "")
        rel = link.get("rel", "")
        thr_count = link.get(f"{{{thr_ns}}}count")
        print(f"  LINK: rel={rel} href={href[:60]} thr:count={thr_count}")
    
    # Look for thr: elements
    for child in entry:
        if "thread" in child.tag.lower() or "thr" in child.tag.lower():
            print(f"  THR ELEMENT: {child.tag} text={child.text} attrs={dict(child.attrib)}")
    
    # Check content for score-related data
    content_el = entry.find(f"{{{atom_ns}}}content")
    if content_el is not None and content_el.text:
        html = content_el.text
        # Look for score-related patterns
        import re
        score_patterns = re.findall(r'(?:score|upvote|point|like)[^<]*', html, re.IGNORECASE)
        if score_patterns:
            print(f"  SCORE PATTERNS: {score_patterns[:5]}")
        
        # Look for data-score attributes
        data_score = re.findall(r'data-[a-z]*score[^=]*=["\']([^"\']+)', html, re.IGNORECASE)
        if data_score:
            print(f"  DATA-SCORE: {data_score}")
        
        # Look for comment count patterns
        comment_patterns = re.findall(r'(\d+)\s*comment', html, re.IGNORECASE)
        if comment_patterns:
            print(f"  COMMENT COUNTS: {comment_patterns}")
