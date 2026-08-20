#!/usr/bin/env python3
"""
FB/X Scraper using Playwright - searches by keywords and topics.
Directly browses Facebook search and X search pages.
Captures real engagement metrics (likes, comments, timestamps).
Only keeps posts from the current monitoring period.
"""
import json
import os
import re
import time
import sys
from datetime import datetime, timezone, timedelta

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

# Time filter: only keep posts from last N hours
MAX_HOURS_OLD = int(os.environ.get("MAX_HOURS_OLD", "24"))


def parse_engagement(text):
    """Parse engagement number from text like '1.2K', '345', '12K'."""
    if not text:
        return 0
    text = text.strip().upper()
    multipliers = {"K": 1000, "M": 1000000, "B": 1000000000}
    for suffix, mult in multipliers.items():
        if text.endswith(suffix):
            try:
                return int(float(text[:-1]) * mult)
            except:
                return 0
    try:
        return int(text.replace(",", "").replace(".", ""))
    except:
        return 0


def parse_timestamp(text):
    """Parse various timestamp formats to ISO format string."""
    if not text:
        return ""
    now = datetime.now(timezone.utc)

    # ISO format (from X): 2026-08-20T14:30:00.000Z
    if re.match(r'\d{4}-\d{2}-\d{2}T', text):
        try:
            return text
        except:
            pass

    # Facebook relative time: "2h", "5h", "Just now", "1d", "3d", "2w"
    if "just now" in text.lower():
        return now.isoformat()

    m = re.match(r'(\d+)\s*(s|m|h|d|w)\b', text.lower())
    if m:
        num = int(m.group(1))
        unit = m.group(2)
        if unit == "s":
            delta = timedelta(seconds=num)
        elif unit == "m":
            delta = timedelta(minutes=num)
        elif unit == "h":
            delta = timedelta(hours=num)
        elif unit == "d":
            delta = timedelta(days=num)
        elif unit == "w":
            delta = timedelta(weeks=num)
        else:
            delta = timedelta(hours=1)
        dt = now - delta
        return dt.isoformat()

    # Facebook date format: "August 20", "Aug 20"
    try:
        dt = datetime.strptime(text, "%B %d")
        dt = dt.replace(year=now.year, tzinfo=timezone.utc)
        if dt > now:
            dt = dt.replace(year=now.year - 1)
        return dt.isoformat()
    except:
        pass

    return ""


def is_recent(created_iso, max_hours=MAX_HOURS_OLD):
    """Check if post is within the time window."""
    if not created_iso:
        return True  # If no timestamp, keep it (better to include than miss)
    try:
        dt = datetime.fromisoformat(created_iso.replace("Z", "+00:00"))
        now = datetime.now(timezone.utc)
        age = now - dt
        return age.total_seconds() < max_hours * 3600
    except:
        return True


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
    """Search Facebook posts using Playwright - by keywords/topics.
    Captures real likes, comments, and timestamps."""
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
            search_url = f"https://www.facebook.com/search/posts/?q={kw.replace(' ', '%20')}"
            print(f"  FB Search: '{kw}' -> {search_url[:80]}")

            try:
                page.goto(search_url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(5)

                # Extract post elements
                post_elements = page.query_selector_all('[role="article"]')
                if not post_elements:
                    post_elements = page.query_selector_all('div[data-ad-comet-preview]')

                kw_posts = 0
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
                        if post_url and not post_url.startswith("http"):
                            post_url = f"https://www.facebook.com{post_url}"

                        # Get timestamp - FB uses various elements
                        time_el = el.query_selector('abbr') or el.query_selector('[data-utime]') or el.query_selector('a[href*="posts"] span[id]')
                        time_text = ""
                        if time_el:
                            time_text = time_el.inner_text().strip()
                        # Also try to find time in any span
                        if not time_text:
                            time_spans = el.query_selector_all('span')
                            for sp in time_spans:
                                sp_text = sp.inner_text().strip()
                                if re.match(r'\d+\s*(s|m|h|d|w)\b', sp_text.lower()) or "just now" in sp_text.lower() or "now" == sp_text.lower():
                                    time_text = sp_text
                                    break

                        created_iso = parse_timestamp(time_text) if time_text else ""

                        # Get likes/reactions - FB shows reaction count
                        likes = 0
                        comments = 0

                        # Try to find reaction/like count
                        # FB uses aria-label with "X reactions" or "X likes"
                        reaction_els = el.query_selector_all('[aria-label*="reaction"], [aria-label*="like"], [aria-label*="comment"]')
                        for r_el in reaction_els:
                            label = r_el.get_attribute("aria-label") or ""
                            count_text = ""
                            if "reaction" in label.lower() or "like" in label.lower():
                                count_text = re.search(r'([\d,.]+[KMB]?)', label)
                                if count_text:
                                    likes = max(likes, parse_engagement(count_text.group(1)))
                            elif "comment" in label.lower():
                                count_text = re.search(r'([\d,.]+[KMB]?)', label)
                                if count_text:
                                    comments = max(comments, parse_engagement(count_text.group(1)))

                        # Also try text-based counts
                        all_text = el.inner_text()
                        comment_match = re.search(r'(\d+)\s*comment', all_text, re.IGNORECASE)
                        if comment_match and comments == 0:
                            comments = parse_engagement(comment_match.group(1))

                        reaction_match = re.search(r'([\d,.]+[KMB]?)\s*(?:reaction|like)', all_text, re.IGNORECASE)
                        if reaction_match and likes == 0:
                            likes = parse_engagement(reaction_match.group(1))

                        post = {
                            "title": text[:200],
                            "url": post_url,
                            "author": author,
                            "score": likes,
                            "num_comments": comments,
                            "selftext": text[:500],
                            "link_flair_text": "Facebook",
                            "subreddit": "facebook",
                            "sort_type": "facebook",
                            "created_utc": created_iso,
                            "search_keyword": kw,
                            "time_text": time_text,
                        }
                        classify_post(post, game)

                        # Filter: only keep recent posts
                        if is_recent(created_iso):
                            if post["title"] not in [pp["title"] for pp in posts]:
                                posts.append(post)
                                kw_posts += 1
                    except Exception:
                        continue

                print(f"    Found {kw_posts} recent posts for '{kw}' (likes/comments captured)")
                time.sleep(2)

            except Exception as e:
                print(f"    FB search error for '{kw}': {str(e)[:80]}")
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

                        # Get timestamp
                        time_text = ""
                        time_spans = el.query_selector_all('span')
                        for sp in time_spans:
                            sp_text = sp.inner_text().strip()
                            if re.match(r'\d+\s*(s|m|h|d|w)\b', sp_text.lower()) or "just now" in sp_text.lower():
                                time_text = sp_text
                                break
                        created_iso = parse_timestamp(time_text) if time_text else ""

                        # Get engagement
                        likes, comments = 0, 0
                        all_text = el.inner_text()
                        comment_match = re.search(r'(\d+)\s*comment', all_text, re.IGNORECASE)
                        if comment_match:
                            comments = parse_engagement(comment_match.group(1))
                        reaction_match = re.search(r'([\d,.]+[KMB]?)\s*(?:reaction|like)', all_text, re.IGNORECASE)
                        if reaction_match:
                            likes = parse_engagement(reaction_match.group(1))

                        post = {
                            "title": text[:200],
                            "url": page_url,
                            "author": author,
                            "score": likes,
                            "num_comments": comments,
                            "selftext": text[:500],
                            "link_flair_text": "Facebook",
                            "subreddit": "facebook",
                            "sort_type": "facebook",
                            "created_utc": created_iso,
                            "search_keyword": f"page:{fb_page}",
                            "time_text": time_text,
                        }
                        classify_post(post, game)
                        if is_recent(created_iso):
                            if post["title"] not in [pp["title"] for pp in posts]:
                                posts.append(post)
                    except Exception:
                        continue
            except Exception as e:
                print(f"  FB page error: {str(e)[:80]}")

        browser.close()

    print(f"  FB Total: {len(posts)} posts for {game} (filtered: last {MAX_HOURS_OLD}h)")
    return posts


def scrape_x_search(game, keywords, limit=15):
    """Search X/Twitter posts using Playwright - by keywords/hashtags/topics.
    Captures real likes, retweets, replies, and timestamps."""
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
            search_url = f"https://x.com/search?q={kw.replace(' ', '%20').replace('#', '%23')}&src=typed_query&f=live"
            print(f"  X Search: '{kw}' -> {search_url[:80]}")

            try:
                page.goto(search_url, timeout=20000, wait_until="domcontentloaded")
                time.sleep(5)

                # Extract tweets
                tweets = page.query_selector_all('article[data-testid="tweet"]')
                if not tweets:
                    tweets = page.query_selector_all('div[data-testid="cellInnerDiv"]')

                kw_posts = 0
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
                        author = ""
                        if author_el:
                            author_text = author_el.inner_text()
                            # Extract @username
                            at_match = re.search(r'@(\w+)', author_text)
                            author = f"@{at_match.group(1)}" if at_match else author_text.split('\n')[0]
                        if not author:
                            author = "X User"

                        # Get timestamp - X uses <time datetime="2026-08-20T14:30:00.000Z">
                        time_el = tweet.query_selector('time')
                        created_iso = ""
                        if time_el:
                            created_iso = time_el.get_attribute('datetime') or ""
                            time_text = time_el.get_attribute('title') or time_el.inner_text()
                        else:
                            time_text = ""

                        # Get engagement metrics - X uses aria-label
                        likes = 0
                        comments = 0
                        retweets = 0

                        # X shows: [Like count] [Reply count] [Repost count] [View count]
                        # They use aria-label like "123 Likes", "45 Replies"
                        action_els = tweet.query_selector_all('button[role="button"]')
                        for btn in action_els:
                            aria_label = btn.get_attribute("aria-label") or ""
                            if "like" in aria_label.lower():
                                m = re.search(r'([\d,.]+[KMB]?)', aria_label)
                                if m:
                                    likes = parse_engagement(m.group(1))
                            elif "reply" in aria_label.lower():
                                m = re.search(r'([\d,.]+[KMB]?)', aria_label)
                                if m:
                                    comments = parse_engagement(m.group(1))
                            elif "repost" in aria_label.lower() or "retweet" in aria_label.lower():
                                m = re.search(r'([\d,.]+[KMB]?)', aria_label)
                                if m:
                                    retweets = parse_engagement(m.group(1))

                        # Fallback: parse from text
                        if likes == 0:
                            all_text = tweet.inner_text()
                            # X bookmark bar shows: "123" "456" "789" "1.2K"
                            numbers = re.findall(r'(?:^|\s)(\d+(?:\.\d+)?[KMB]?)(?:\s|$)', all_text)
                            if len(numbers) >= 3:
                                comments = max(comments, parse_engagement(numbers[0]))
                                retweets = max(retweets, parse_engagement(numbers[1]))
                                likes = max(likes, parse_engagement(numbers[2]))

                        post = {
                            "title": text[:200],
                            "url": tweet_url,
                            "author": author,
                            "score": likes,
                            "num_comments": comments + retweets,
                            "selftext": text[:500],
                            "link_flair_text": "X/Twitter",
                            "subreddit": "x_twitter",
                            "sort_type": "x",
                            "created_utc": created_iso,
                            "search_keyword": kw,
                            "retweets": retweets,
                            "time_text": time_text if 'time_text' in dir() else "",
                        }
                        classify_post(post, game)

                        # Filter: only keep recent posts
                        if is_recent(created_iso):
                            if post["title"] not in [pp["title"] for pp in posts]:
                                posts.append(post)
                                kw_posts += 1
                    except Exception:
                        continue

                print(f"    Found {kw_posts} recent posts for '{kw}' (engagement captured)")
                time.sleep(2)

            except Exception as e:
                print(f"    X search error for '{kw}': {str(e)[:80]}")
                continue

        browser.close()

    print(f"  X Total: {len(posts)} posts for {game} (filtered: last {MAX_HOURS_OLD}h)")
    return posts


def main():
    print("=== FB/X Playwright Scraper (Keyword Search + Engagement Metrics) ===")
    print(f"Time filter: only posts from last {MAX_HOURS_OLD} hours")

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

        # Print engagement summary
        fb_with_likes = sum(1 for p in fb_posts if p.get("score", 0) > 0 or p.get("num_comments", 0) > 0)
        x_with_likes = sum(1 for p in x_posts if p.get("score", 0) > 0 or p.get("num_comments", 0) > 0)
        print(f"\n{game}: FB={len(fb_posts)} ({fb_with_likes} with engagement) X={len(x_posts)} ({x_with_likes} with engagement)")
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
