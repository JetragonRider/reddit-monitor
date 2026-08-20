#!/usr/bin/env python3
"""
FB/X Scraper using Playwright - directly browses Facebook and X.com
Bypasses search engine IP blocking entirely.
"""
import json
import os
import re
import time
import sys

# Game-specific search terms and official accounts
FB_GROUPS = {
    "Palworld": "Palworld",
    "CS2": "Counter-Strike",
    "Valorant": "VALORANT",
    "LOL": "League of Legends",
    "DeltaForce": "Delta Force",
    "TFT": "Teamfight Tactics",
}

X_ACCOUNTS = {
    "Palworld": ["Palworld_EN", "Pocketpair"],
    "CS2": ["csgo", "CS2"],
    "Valorant": ["PlayVALORANT"],
    "LOL": ["LeagueOfLegends", "lolesports"],
    "DeltaForce": ["DeltaForceGame"],
    "TFT": ["TFT", "TFTesports"],
}

# Classification keywords (same as reddit_monitor.py)
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


def scrape_x_with_playwright(game, accounts, limit=15):
    """Scrape X/Twitter public posts using Playwright."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  Playwright not installed, skipping X scrape")
        return []
    
    posts = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 720},
        )
        page = context.new_page()
        
        for account in accounts:
            url = f"https://x.com/{account}"
            print(f"  X Playwright: {url}")
            try:
                page.goto(url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(3)  # Wait for content to load
                
                # Try to extract tweets from page
                tweets = page.query_selector_all('article[data-testid="tweet"]')
                
                if not tweets:
                    # Try alternative selectors
                    tweets = page.query_selector_all('div[data-testid="cellInnerDiv"]')
                
                for tweet in tweets[:limit]:
                    try:
                        # Get tweet text
                        text_el = tweet.query_selector('[data-testid="tweetText"]')
                        text = text_el.inner_text() if text_el else ""
                        
                        if not text:
                            text_el = tweet.query_selector('div[lang]')
                            text = text_el.inner_text() if text_el else ""
                        
                        if not text or len(text) < 5:
                            continue
                        
                        # Get tweet link
                        link_el = tweet.query_selector('a[href*="/status/"]')
                        tweet_url = f"https://x.com{link_el.get_attribute('href')}" if link_el else url
                        
                        # Get author
                        author_el = tweet.query_selector('[data-testid="User-Name"]')
                        author = author_el.inner_text().split('\n')[0] if author_el else f"@{account}"
                        
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
                        }
                        classify_post(post, game)
                        posts.append(post)
                    except Exception as e:
                        continue
                
                print(f"  X: Found {len([p for p in posts if p['author'].startswith('@')])} posts from @{account}")
                time.sleep(2)
                
            except Exception as e:
                print(f"  X error for @{account}: {str(e)[:60]}")
                continue
        
        browser.close()
    
    print(f"  X (Playwright): Found {len(posts)} posts for {game}")
    return posts


def scrape_facebook_with_playwright(game, search_term, limit=15):
    """Scrape Facebook public posts using Playwright."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  Playwright not installed, skipping FB scrape")
        return []
    
    posts = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 720},
        )
        page = context.new_page()
        
        # Search Facebook for public posts about the game
        search_url = f"https://www.facebook.com/search/posts/?q={search_term.replace(' ', '%20')}"
        print(f"  FB Playwright: {search_url}")
        
        try:
            page.goto(search_url, timeout=20000, wait_until="domcontentloaded")
            time.sleep(5)  # Wait for JS to render
            
            # Try to extract posts from search results
            post_elements = page.query_selector_all('div[data-ad-comet-preview=""]')
            
            if not post_elements:
                # Try alternative selectors
                post_elements = page.query_selector_all('[role="article"]')
            
            if not post_elements:
                post_elements = page.query_selector_all('div.x1yztbd')
            
            for el in post_elements[:limit]:
                try:
                    # Get post text
                    text_el = el.query_selector('div[dir="auto"]')
                    text = text_el.inner_text() if text_el else ""
                    
                    if not text or len(text) < 10:
                        continue
                    
                    # Get author
                    author_el = el.query_selector('span a span')
                    author = author_el.inner_text() if author_el else "Facebook User"
                    
                    # Get link
                    link_el = el.query_selector('a[href*="/posts/"]')
                    post_url = link_el.get_attribute('href') if link_el else ""
                    
                    post = {
                        "title": text[:200],
                        "url": post_url or search_url,
                        "author": author,
                        "score": 0,
                        "num_comments": 0,
                        "selftext": text[:500],
                        "link_flair_text": "Facebook",
                        "subreddit": "facebook",
                        "sort_type": "facebook",
                        "created_utc": 0,
                    }
                    classify_post(post, game)
                    posts.append(post)
                except Exception:
                    continue
            
            # If search didn't work, try official game pages
            if not posts:
                page_pages = {
                    "Palworld": "PalworldOfficial",
                    "CS2": "CounterStrikeOfficial",
                    "Valorant": "VALORANT",
                    "LOL": "LeagueOfLegends",
                    "DeltaForce": "DeltaForceGame",
                    "TFT": "TeamfightTactics",
                }
                
                fb_page = page_pages.get(game, game)
                page_url = f"https://www.facebook.com/{fb_page}/posts/"
                print(f"  FB Page: {page_url}")
                
                page.goto(page_url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(5)
                
                post_elements = page.query_selector_all('[role="article"]')
                
                for el in post_elements[:limit]:
                    try:
                        text_el = el.query_selector('div[dir="auto"]')
                        text = text_el.inner_text() if text_el else ""
                        
                        if not text or len(text) < 10:
                            continue
                        
                        author_el = el.query_selector('span a span')
                        author = author_el.inner_text() if author_el else fb_page
                        
                        post = {
                            "title": text[:200],
                            "url": page_url,
                            "author": author,
                            "score": 0,
                            "num_comments": 0,
                            "selftext": text[:500],
                            "link_flair_text": "Facebook",
                            "subreddit": "facebook",
                            "sort_type": "facebook",
                            "created_utc": 0,
                        }
                        classify_post(post, game)
                        posts.append(post)
                    except Exception:
                        continue
            
            print(f"  FB (Playwright): Found {len(posts)} posts for {game}")
            
        except Exception as e:
            print(f"  FB error: {str(e)[:80]}")
        
        browser.close()
    
    return posts


def main():
    print("=== FB/X Playwright Scraper ===")
    
    all_data = {}
    
    for game in FB_GROUPS:
        print(f"\n--- {game} ---")
        
        # Scrape X
        x_accounts = X_ACCOUNTS.get(game, [game])
        x_posts = scrape_x_with_playwright(game, x_accounts, limit=15)
        
        # Scrape Facebook
        fb_search = FB_GROUPS.get(game, game)
        fb_posts = scrape_facebook_with_playwright(game, fb_search, limit=15)
        
        all_data[game] = {
            "fb_posts": fb_posts,
            "x_posts": x_posts,
        }
        
        print(f"  {game}: FB={len(fb_posts)} X={len(x_posts)}")
        time.sleep(3)
    
    # Save results
    output_file = os.environ.get("OUTPUT_FILE", "/tmp/fb_x_playwright_results.json")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    
    print(f"\nResults saved to {output_file}")
    print(f"Total: {sum(len(d['fb_posts']) for d in all_data.values())} FB posts, {sum(len(d['x_posts']) for d in all_data.values())} X posts")


if __name__ == "__main__":
    main()
