#!/usr/bin/env python3
"""
FB/X Scraper using Playwright - searches by keywords and topics.
Directly browses Facebook search and X search pages.
Bypasses search engine IP blocking entirely.
"""
import json
import os
import re
import time
import sys

# ============================================================
# Facebook 搜索关键词（按话题/关键词搜索）
# ============================================================
FB_SEARCH_KEYWORDS = {
    "Palworld": [
        "Palworld breeding guide",
        "Palworld best base",
        "Palworld pal tier list",
        "Palworld update",
        "Palworld patch notes",
    ],
    "CS2": [
        "CS2 best crosshair",
        "CS2 weapon stats",
        "CS2 new update",
        "Counter-Strike 2 patch",
        "CS2 pro settings",
    ],
    "Valorant": [
        "Valorant agent tier list",
        "Valorant best skin",
        "Valorant patch notes",
        "Valorant new agent",
        "Valorant rank distribution",
    ],
    "LOL": [
        "League of Legends patch notes",
        "LOL best build",
        "League of Legends tier list",
        "LOL ARAM guide",
        "League of Legends esports",
    ],
    "DeltaForce": [
        "Delta Force weapon build",
        "Delta Force Hawk Ops",
        "Delta Force best loadout",
        "Delta Force update",
        "Delta Force tips",
    ],
    "TFT": [
        "TFT best comp",
        "Teamfight Tactics tier list",
        "TFT patch notes",
        "TFT augment guide",
        "TFT meta build",
    ],
}

# ============================================================
# X/Twitter 搜索关键词（按话题/关键词搜索）
# ============================================================
X_SEARCH_KEYWORDS = {
    "Palworld": [
        "Palworld breeding",
        "Palworld update",
        "Palworld patch",
        "#Palworld",
        "Palworld new pal",
    ],
    "CS2": [
        "CS2 update",
        "CS2 patch notes",
        "CS2 Major",
        "#CS2",
        "Counter-Strike 2 news",
    ],
    "Valorant": [
        "Valorant patch notes",
        "Valorant new agent",
        "Valorant skin",
        "#Valorant",
        "VCT 2026",
    ],
    "LOL": [
        "League of Legends patch",
        "LOL tier list",
        "League of Legends Worlds",
        "#LeagueOfLegends",
        "LoL esports",
    ],
    "DeltaForce": [
        "Delta Force Hawk Ops",
        "Delta Force update",
        "Delta Force weapon",
        "#DeltaForce",
        "Delta Force gameplay",
    ],
    "TFT": [
        "TFT best comp",
        "TFT patch notes",
        "TFT tier list",
        "#TFT",
        "Teamfight Tactics meta",
    ],
}

# Also visit official pages as fallback
FB_PAGES = {
    "Palworld": "PalworldOfficial",
    "CS2": "CounterStrike",
    "Valorant": "VALORANT",
    "LOL": "LeagueOfLegends",
    "DeltaForce": "DeltaForceGame",
    "TFT": "TeamfightTactics",
}

# Classification keywords
TOOL_KEYWORDS = {
    "DeltaForce": ["weapon build", "loadout", "attachment", "blueprint", "modify", "best build"],
    "Valorant": ["skin", "agent", "stats", "win rate", "pick rate", "agent guide"],
    "Palworld": ["breeding", "pair", "combination", "paldex", "species", "passive skill"],
    "LOL": ["aram", "build", "rune", "item", "counter", "tier list", "matchup"],
    "TFT": ["comp", "team comp", "tier list", "meta", "augment"],
    "CS2": ["stats", "crosshair", "sensitivity", "weapon stats", "accuracy"],
}

PAIN_KEYWORDS = [
    "broken", "bug", "terrible", "frustrating", "awful", "worst", "hate", "annoying",
    "wish there was", "is there a tool", "why is there no", "need a way to",
    "wrong data", "incorrect", "outdated", "no solution", "can't find", "still stuck",
]


def classify_post(post, game):
    """Classify a post as A (tool opportunity) or B (pain point)."""
    title = post.get("title", "").lower()
    selftext = post.get("selftext", "").lower()
    combined = f"{title} {selftext}"
    
    tool_kws = TOOL_KEYWORDS.get(game, [])
    matched_tool = [kw for kw in tool_kws if kw in combined]
    if matched_tool:
        post["is_tool_opportunity"] = True
        post["classification"] = "A-工具引流机会"
        post["classification_note"] = f"匹配: {', '.join(matched_tool[:3])} → {game}工具"
        return
    
    matched_pain = [kw for kw in PAIN_KEYWORDS if kw in combined]
    if matched_pain:
        post["is_pain_point"] = True
        post["classification"] = "B-用户痛点"
        post["classification_note"] = f"匹配: {', '.join(matched_pain[:3])}"
        return
    
    post["is_tool_opportunity"] = False
    post["is_pain_point"] = False
    post["classification"] = ""
    post["classification_note"] = ""


def scrape_facebook_search(game, keywords, limit=15):
    """Search Facebook posts using Playwright - by keywords/topics."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  Playwright not installed")
        return []
    
    posts = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 720},
            locale="en-US",
        )
        page = context.new_page()
        
        for kw in keywords:
            # Facebook search URL for posts
            search_url = f"https://www.facebook.com/search/posts/?q={kw.replace(' ', '%20')}"
            print(f"  FB Search: '{kw}' -> {search_url[:80]}")
            
            try:
                page.goto(search_url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(5)
                
                # Extract post elements
                post_elements = page.query_selector_all('[role="article"]')
                if not post_elements:
                    post_elements = page.query_selector_all('div[data-ad-comet-preview]')
                
                for el in post_elements[:limit]:
                    try:
                        # Get post text
                        text_els = el.query_selector_all('div[dir="auto"]')
                        text = " ".join([t.inner_text() for t in text_els if t.inner_text().strip()])
                        
                        if not text or len(text) < 10:
                            continue
                        
                        # Get author
                        author_el = el.query_selector('span a span') or el.query_selector('a span span')
                        author = author_el.inner_text() if author_el else "Facebook User"
                        
                        # Get link
                        link_el = el.query_selector('a[href*="/posts/"]') or el.query_selector('a[href*="/permalink/"]')
                        post_url = link_el.get_attribute('href') if link_el else search_url
                        
                        # Get timestamp
                        time_el = el.query_selector('abbr') or el.query_selector('[data-utime]')
                        timestamp = time_el.inner_text() if time_el else ""
                        
                        post = {
                            "title": text[:200],
                            "url": post_url,
                            "author": author,
                            "score": 0,
                            "num_comments": 0,
                            "selftext": text[:500],
                            "link_flair_text": "Facebook",
                            "subreddit": "facebook",
                            "sort_type": "facebook",
                            "created_utc": 0,
                            "search_keyword": kw,
                        }
                        classify_post(post, game)
                        
                        # Dedup by title
                        if post["title"] not in [p["title"] for p in posts]:
                            posts.append(post)
                    except Exception:
                        continue
                
                print(f"    Found {len([p for p in posts if p.get('search_keyword') == kw])} posts for '{kw}'")
                time.sleep(2)
                
            except Exception as e:
                print(f"    FB search error for '{kw}': {str(e)[:60]}")
                continue
        
        # If no results from search, try official page
        if not posts:
            fb_page = FB_PAGES.get(game, game)
            page_url = f"https://www.facebook.com/{fb_page}/posts/"
            print(f"  FB Page fallback: {page_url}")
            try:
                page.goto(page_url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(5)
                
                post_elements = page.query_selector_all('[role="article"]')
                for el in post_elements[:limit]:
                    try:
                        text_els = el.query_selector_all('div[dir="auto"]')
                        text = " ".join([t.inner_text() for t in text_els if t.inner_text().strip()])
                        if not text or len(text) < 10:
                            continue
                        author_el = el.query_selector('span a span')
                        author = author_el.inner_text() if author_el else fb_page
                        
                        post = {
                            "title": text[:200],
                            "url": page_url,
                            "author": author,
                            "score": 0, "num_comments": 0,
                            "selftext": text[:500],
                            "link_flair_text": "Facebook",
                            "subreddit": "facebook",
                            "sort_type": "facebook",
                            "created_utc": 0,
                            "search_keyword": f"page:{fb_page}",
                        }
                        classify_post(post, game)
                        if post["title"] not in [p["title"] for p in posts]:
                            posts.append(post)
                    except Exception:
                        continue
            except Exception as e:
                print(f"  FB page error: {str(e)[:60]}")
        
        browser.close()
    
    print(f"  FB Total: {len(posts)} posts for {game}")
    return posts


def scrape_x_search(game, keywords, limit=15):
    """Search X/Twitter posts using Playwright - by keywords/hashtags/topics."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  Playwright not installed")
        return []
    
    posts = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 720},
            locale="en-US",
        )
        page = context.new_page()
        
        for kw in keywords:
            # X search URL
            search_url = f"https://x.com/search?q={kw.replace(' ', '%20').replace('#', '%23')}&src=typed_query&f=live"
            print(f"  X Search: '{kw}' -> {search_url[:80]}")
            
            try:
                page.goto(search_url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(5)
                
                # Extract tweets
                tweets = page.query_selector_all('article[data-testid="tweet"]')
                if not tweets:
                    tweets = page.query_selector_all('div[data-testid="cellInnerDiv"]')
                
                for tweet in tweets[:limit]:
                    try:
                        # Get tweet text
                        text_el = tweet.query_selector('[data-testid="tweetText"]')
                        text = text_el.inner_text() if text_el else ""
                        
                        if not text_el:
                            text_el = tweet.query_selector('div[lang]')
                            text = text_el.inner_text() if text_el else ""
                        
                        if not text or len(text) < 5:
                            continue
                        
                        # Get tweet URL
                        link_el = tweet.query_selector('a[href*="/status/"]')
                        tweet_url = f"https://x.com{link_el.get_attribute('href')}" if link_el else search_url
                        
                        # Get author
                        author_el = tweet.query_selector('[data-testid="User-Name"]')
                        author = author_el.inner_text().split('\n')[0] if author_el else "X User"
                        
                        # Get timestamp
                        time_el = tweet.query_selector('time')
                        timestamp = time_el.get_attribute('datetime') if time_el else ""
                        
                        post = {
                            "title": text[:200],
                            "url": tweet_url,
                            "author": author,
                            "score": 0,
                            "num_comments": 0,
                            "selftext": text[:500],
                            "link_flair_text": "X/Twitter",
                            "subreddit": "x_twitter",
                            "sort_type": "x",
                            "created_utc": 0,
                            "search_keyword": kw,
                        }
                        classify_post(post, game)
                        
                        if post["title"] not in [p["title"] for p in posts]:
                            posts.append(post)
                    except Exception:
                        continue
                
                print(f"    Found {len([p for p in posts if p.get('search_keyword') == kw])} posts for '{kw}'")
                time.sleep(2)
                
            except Exception as e:
                print(f"    X search error for '{kw}': {str(e)[:60]}")
                continue
        
        browser.close()
    
    print(f"  X Total: {len(posts)} posts for {game}")
    return posts


def main():
    print("=== FB/X Playwright Scraper (Keyword Search) ===")
    
    all_data = {}
    
    for game in FB_SEARCH_KEYWORDS:
        print(f"\n{'='*60}")
        print(f"--- {game} ---")
        
        # Facebook search by keywords
        fb_keywords = FB_SEARCH_KEYWORDS.get(game, [game])
        print(f"\nFB keywords: {fb_keywords}")
        fb_posts = scrape_facebook_search(game, fb_keywords, limit=15)
        
        # X search by keywords/hashtags
        x_keywords = X_SEARCH_KEYWORDS.get(game, [game])
        print(f"\nX keywords: {x_keywords}")
        x_posts = scrape_x_search(game, x_keywords, limit=15)
        
        all_data[game] = {
            "fb_posts": fb_posts,
            "x_posts": x_posts,
        }
        
        print(f"\n{game}: FB={len(fb_posts)} X={len(x_posts)}")
        time.sleep(3)
    
    # Save
    output_file = os.environ.get("OUTPUT_FILE", "/tmp/fb_x_results.json")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    
    total_fb = sum(len(d["fb_posts"]) for d in all_data.values())
    total_x = sum(len(d["x_posts"]) for d in all_data.values())
    print(f"\n{'='*60}")
    print(f"Total: FB={total_fb} X={total_x}")
    print(f"Results saved to {output_file}")


if __name__ == "__main__":
    main()
