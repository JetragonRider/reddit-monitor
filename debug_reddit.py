#!/usr/bin/env python3
"""Debug: Check what Reddit RSS/JSON actually returns."""
import urllib.request, ssl, json, sys

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# 1. Check RSS content
print("=== 1. RSS Feed ===")
url = "https://www.reddit.com/r/Palworld/hot/.rss?limit=3"
try:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 Chrome/131.0.0.0"})
    with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
        content = resp.read()
        print(f"Status: {resp.status}, Size: {len(content)}")
        text = content.decode("utf-8", errors="replace")
        # Print full XML
        print(text[:5000])
        # Check for score/comment keywords
        if "score" in text.lower():
            print("\n>>> FOUND 'score' in RSS!")
        if "comment" in text.lower():
            print(">>> FOUND 'comment' in RSS!")
        if "points" in text.lower():
            print(">>> FOUND 'points' in RSS!")
        if "thr:" in text.lower():
            print(">>> FOUND thr: namespace in RSS!")
except Exception as e:
    print(f"RSS error: {e}")

# 2. Try JSON with different User-Agents
for ua_name, ua in [
    ("Desktop Chrome", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"),
    ("Mobile Safari", "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0) AppleWebKit/605.1.15 Mobile/15E148 Safari/604.1"),
    ("Bot", "python-requests/2.31.0"),
    ("Reddit Bot", "CommunityMonitor/1.0"),
]:
    print(f"\n=== 2. JSON API ({ua_name}) ===")
    json_url = "https://www.reddit.com/r/Palworld/hot.json?limit=3"
    try:
        req = urllib.request.Request(json_url, headers={"User-Agent": ua, "Accept": "application/json"})
        with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
            print(f"Status: {resp.status}")
            data = resp.read()
            j = json.loads(data.decode("utf-8"))
            children = j.get("data", {}).get("children", [])
            print(f"Posts: {len(children)}")
            for c in children[:3]:
                d = c["data"]
                print(f"  {d.get('title','')[:40]} | score={d.get('score',0)} | comments={d.get('num_comments',0)}")
            break  # Success, stop trying
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code}: {e.reason}")
    except Exception as e:
        print(f"Error: {str(e)[:80]}")

# 3. Try old.reddit.com
print("\n=== 3. old.reddit.com JSON ===")
try:
    req = urllib.request.Request("https://old.reddit.com/r/Palworld/hot.json?limit=3",
                                 headers={"User-Agent": "Mozilla/5.0 Chrome/131.0.0.0", "Accept": "application/json"})
    with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
        print(f"Status: {resp.status}")
        data = resp.read()
        j = json.loads(data.decode("utf-8"))
        children = j.get("data", {}).get("children", [])
        print(f"Posts: {len(children)}")
        for c in children[:3]:
            d = c["data"]
            print(f"  {d.get('title','')[:40]} | score={d.get('score',0)} | comments={d.get('num_comments',0)}")
except urllib.error.HTTPError as e:
    print(f"HTTP {e.code}: {e.reason}")
except Exception as e:
    print(f"Error: {str(e)[:80]}")

# 4. Try Reddit API v3
print("\n=== 4. Reddit API v3 ===")
try:
    req = urllib.request.Request("https://api.reddit.com/r/Palworld/hot?limit=3",
                                 headers={"User-Agent": "Mozilla/5.0 Chrome/131.0.0.0", "Accept": "application/json"})
    with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
        print(f"Status: {resp.status}")
        data = resp.read()
        j = json.loads(data.decode("utf-8"))
        children = j.get("data", {}).get("children", [])
        print(f"Posts: {len(children)}")
        for c in children[:3]:
            d = c["data"]
            print(f"  {d.get('title','')[:40]} | score={d.get('score',0)} | comments={d.get('num_comments',0)}")
except urllib.error.HTTPError as e:
    print(f"HTTP {e.code}: {e.reason}")
except Exception as e:
    print(f"Error: {str(e)[:80]}")
