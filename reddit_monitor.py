#!/usr/bin/env python3
"""
Reddit Community Monitor - Daily Patrol Report Generator
Uses Reddit RSS feeds (less likely to be blocked) + JSON API for comments.
Falls back to JSON API if RSS fails.
"""

import json
import os
import sys
import time
import datetime
import urllib.request
import urllib.error
import ssl
import xml.etree.ElementTree as ET
import re

# For Excel generation
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Game subreddits to monitor (with fallbacks)
SUBREDDITS = {
    "Palworld": ["Palworld"],
    "CS2": ["GlobalOffensive"],
    "Valorant": ["Valorant"],
    "LOL": ["leagueoflegends"],
    "DeltaForce": ["DeltaForce"],
    "TFT": ["TeamfightTactics"],
}

POSTS_PER_SUB = 20
COMMENTS_PER_POST = 10

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
}


def fetch_url(url, retries=3, delay=5):
    """Fetch URL content with retries. Returns (status_code, content_bytes)."""
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
                return resp.status, resp.read()
        except urllib.error.HTTPError as e:
            print(f"  HTTP {e.code}: {e.reason} (attempt {attempt+1}/{retries})", file=sys.stderr)
            if e.code in (429, 503):
                wait = delay * (attempt + 1) * 2
                print(f"  Rate limited, waiting {wait}s...", file=sys.stderr)
                time.sleep(wait)
            elif attempt < retries - 1:
                time.sleep(delay)
        except Exception as e:
            print(f"  Error: {e} (attempt {attempt+1}/{retries})", file=sys.stderr)
            if attempt < retries - 1:
                time.sleep(delay)
    return None, None


def fetch_json(url, retries=3, delay=5):
    """Fetch JSON from URL with retries."""
    status, content = fetch_url(url, retries, delay)
    if content:
        try:
            return json.loads(content.decode("utf-8"))
        except json.JSONDecodeError as e:
            print(f"  JSON parse error: {e}", file=sys.stderr)
    return None


def fetch_rss(subreddit, limit=POSTS_PER_SUB):
    """Fetch posts from Reddit RSS feed (less likely to be blocked)."""
    url = f"https://www.reddit.com/r/{subreddit}/hot/.rss?limit={limit}"
    print(f"Fetching RSS: {url}")
    status, content = fetch_url(url, retries=3, delay=5)
    if not content:
        print(f"  RSS fetch failed for r/{subreddit}")
        return []

    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        print(f"  RSS parse error: {e}")
        return []

    # Reddit RSS uses Atom feed format
    ns = {
        "atom": "http://www.w3.org/2005/Atom",
        "media": "http://search.yahoo.com/mrss/",
    }

    entries = root.findall("atom:entry", ns)
    if not entries:
        # Try without namespace
        entries = root.findall("{http://www.w3.org/2005/Atom}entry")
    
    if not entries:
        # Try plain RSS format
        entries = root.findall(".//item")
        ns = {}

    print(f"  Found {len(entries)} entries in RSS")
    posts = []

    for entry in entries[:limit]:
        post = {"subreddit": subreddit}

        # Atom format
        title = entry.find("atom:title", ns)
        if title is None:
            title = entry.find("{http://www.w3.org/2005/Atom}title")
        post["title"] = title.text if title is not None else ""

        # Link
        link = entry.find("atom:link", ns)
        if link is None:
            link = entry.find("{http://www.w3.org/2005/Atom}link")
        post["url"] = link.get("href", "") if link is not None else ""

        # Author
        author = entry.find("atom:author/atom:name", ns)
        if author is None:
            author = entry.find("{http://www.w3.org/2005/Atom}author/{http://www.w3.org/2005/Atom}name")
        post["author"] = author.text if author is not None else ""

        # Content/summary
        content_elem = entry.find("atom:content", ns)
        if content_elem is None:
            content_elem = entry.find("{http://www.w3.org/2005/Atom}content")
        summary = entry.find("atom:summary", ns)
        if summary is None:
            summary = entry.find("{http://www.w3.org/2005/Atom}summary")
        
        text = ""
        if content_elem is not None and content_elem.text:
            text = content_elem.text
        elif summary is not None and summary.text:
            text = summary.text
        
        # Clean HTML tags
        text = re.sub(r'<[^>]+>', '', text)[:500]
        post["selftext"] = text

        # Published date
        published = entry.find("atom:published", ns)
        if published is None:
            published = entry.find("{http://www.w3.org/2005/Atom}published")
        post["created_utc"] = 0
        if published is not None:
            try:
                dt = datetime.datetime.fromisoformat(published.text.replace("Z", "+00:00"))
                post["created_utc"] = dt.timestamp()
            except Exception:
                pass

        # Category/flair
        category = entry.find("atom:category", ns)
        if category is None:
            category = entry.find("{http://www.w3.org/2005/Atom}category")
        post["link_flair_text"] = category.get("term", "") if category is not None else ""

        # Try to extract score from content (Reddit includes it in the feed)
        post["score"] = 0
        post["num_comments"] = 0
        post["upvote_ratio"] = 0

        # Try to find score in the content text
        score_match = re.search(r'(\d+)\s*points?', text, re.IGNORECASE)
        if score_match:
            post["score"] = int(score_match.group(1))
        comment_match = re.search(r'(\d+)\s*comments?', text, re.IGNORECASE)
        if comment_match:
            post["num_comments"] = int(comment_match.group(1))

        posts.append(post)

    return posts


def get_hot_posts(subreddit_names, limit=POSTS_PER_SUB):
    """Get hot posts from a subreddit using RSS first, then JSON API."""
    if isinstance(subreddit_names, str):
        subreddit_names = [subreddit_names]

    for subreddit in subreddit_names:
        # Try RSS first (less likely to be blocked)
        posts = fetch_rss(subreddit, limit)
        if posts:
            # Enrich with JSON API for scores/comments (best effort)
            print(f"  Enriching with JSON API data...")
            json_url = f"https://www.reddit.com/r/{subreddit}/hot.json?limit={limit}"
            json_data = fetch_json(json_url, retries=2, delay=3)
            if json_data and "data" in json_data and "children" in json_data["data"]:
                for i, child in enumerate(json_data["data"]["children"][:len(posts)]):
                    d = child["data"]
                    if i < len(posts):
                        posts[i]["score"] = d.get("score", posts[i]["score"])
                        posts[i]["num_comments"] = d.get("num_comments", posts[i]["num_comments"])
                        posts[i]["upvote_ratio"] = d.get("upvote_ratio", 0)
                        posts[i]["link_flair_text"] = d.get("link_flair_text", posts[i]["link_flair_text"])
                        # Get permalink for comments
                        posts[i]["url"] = f"https://www.reddit.com{d.get('permalink', '')}"
                time.sleep(1)
            return posts, subreddit
        
        # Fallback: try JSON API directly
        print(f"  RSS failed for r/{subreddit}, trying JSON API...")
        json_url = f"https://www.reddit.com/r/{subreddit}/hot.json?limit={limit}"
        print(f"Fetching JSON: {json_url}")
        data = fetch_json(json_url, retries=3, delay=5)
        if data and "data" in data and "children" in data["data"]:
            posts = []
            for child in data["data"]["children"]:
                d = child["data"]
                posts.append({
                    "title": d.get("title", ""),
                    "author": d.get("author", ""),
                    "score": d.get("score", 0),
                    "num_comments": d.get("num_comments", 0),
                    "url": f"https://www.reddit.com{d.get('permalink', '')}",
                    "created_utc": d.get("created_utc", 0),
                    "selftext": d.get("selftext", "")[:500],
                    "link_flair_text": d.get("link_flair_text", ""),
                    "upvote_ratio": d.get("upvote_ratio", 0),
                    "subreddit": d.get("subreddit", subreddit),
                })
            if posts:
                return posts, subreddit
        
        print(f"  r/{subreddit} completely failed, trying next fallback...")
        time.sleep(3)

    return [], subreddit_names[0] if subreddit_names else "unknown"


def get_top_comments(permalink, limit=COMMENTS_PER_POST):
    """Get top comments from a post via JSON API."""
    url = f"https://www.reddit.com{permalink}.json?limit={limit}&sort=top"
    data = fetch_json(url, retries=2, delay=3)
    if not data or len(data) < 2:
        # Try old.reddit.com
        url2 = f"https://old.reddit.com{permalink}.json?limit={limit}&sort=top"
        data = fetch_json(url2, retries=2, delay=3)
        if not data or len(data) < 2:
            return []

    comments = []
    try:
        for child in data[1]["data"]["children"]:
            c = child["data"]
            if c.get("body") in ("[deleted]", "[removed]", None):
                continue
            comments.append({
                "author": c.get("author", ""),
                "body": c.get("body", "")[:300],
                "score": c.get("score", 0),
            })
    except (KeyError, IndexError):
        pass
    return comments


def summarize_discussion(posts_with_comments):
    """Generate a text summary of main discussion topics."""
    if not posts_with_comments:
        return "No data available."

    topics = {}
    for p in posts_with_comments:
        flair = p.get("link_flair_text") or "General"
        if flair not in topics:
            topics[flair] = []
        topics[flair].append(p["title"])

    lines = []
    for flair, titles in topics.items():
        lines.append(f"[{flair}] ({len(titles)} posts)")
        for t in titles[:3]:
            lines.append(f"  - {t}")
        if len(titles) > 3:
            lines.append(f"  ... and {len(titles)-3} more")
        lines.append("")

    return "\n".join(lines)


def create_excel_report(all_data, output_dir="."):
    """Create Excel report matching the template format."""
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M")
    filename = f"reddit_monitor_{date_str}_{time_str}.xlsx"
    filepath = os.path.join(output_dir, filename)

    wb = openpyxl.Workbook()

    # Styling
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    title_font = Font(bold=True, size=14)
    section_font = Font(bold=True, size=12, color="4472C4")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    wrap_align = Alignment(wrap_text=True, vertical="top")

    # === Sheet 1: 汇总日报 ===
    ws_summary = wb.active
    ws_summary.title = "汇总日报"

    ws_summary["A1"] = f"Reddit社区巡查日报 {now.strftime('%Y年%m月%d日 %H:%M')}"
    ws_summary["A1"].font = title_font
    ws_summary.merge_cells("A1:G1")

    ws_summary["A3"] = "一、巡查概况"
    ws_summary["A3"].font = section_font

    ws_summary["A4"] = "社区"
    ws_summary["B4"] = "巡查帖子数"
    ws_summary["C4"] = "总评论数"
    ws_summary["D4"] = "主要讨论话题"
    for col in ["A4", "B4", "C4", "D4"]:
        ws_summary[col].font = header_font
        ws_summary[col].fill = header_fill

    row = 5
    total_posts = 0
    total_comments = 0
    for game, data in all_data.items():
        posts = data.get("posts", [])
        total_comments_count = sum(p.get("num_comments", 0) for p in posts)
        total_posts += len(posts)
        total_comments += total_comments_count

        top_titles = [p["title"] for p in posts[:5]]
        summary = " | ".join(top_titles[:3])

        ws_summary[f"A{row}"] = game
        ws_summary[f"B{row}"] = len(posts)
        ws_summary[f"C{row}"] = total_comments_count
        ws_summary[f"D{row}"] = summary
        ws_summary[f"D{row}"].alignment = wrap_align
        row += 1

    ws_summary[f"A{row}"] = "合计"
    ws_summary[f"B{row}"] = total_posts
    ws_summary[f"C{row}"] = total_comments
    ws_summary[f"A{row}"].font = Font(bold=True)
    row += 2

    ws_summary[f"A{row}"] = "二、各社区讨论热点"
    ws_summary[f"A{row}"].font = section_font
    row += 1

    for game, data in all_data.items():
        posts = data.get("posts", [])
        ws_summary[f"A{row}"] = f"【{game}】r/{data.get('subreddit', game)}"
        ws_summary[f"A{row}"].font = Font(bold=True)
        row += 1
        for p in posts[:5]:
            ws_summary[f"A{row}"] = f"  - {p['title']} (score:{p['score']} comments:{p['num_comments']})"
            ws_summary[f"A{row}"].alignment = wrap_align
            row += 1
        row += 1

    ws_summary.column_dimensions["A"].width = 40
    ws_summary.column_dimensions["B"].width = 15
    ws_summary.column_dimensions["C"].width = 15
    ws_summary.column_dimensions["D"].width = 60

    # === Sheet per game ===
    for game, data in all_data.items():
        ws = wb.create_sheet(title=game[:31])
        sub_name = data.get("subreddit", game)
        ws["A1"] = f"{now.strftime('%m月%d日 %H:%M')} r/{sub_name} 数据"
        ws["A1"].font = title_font
        ws.merge_cells("A1:J1")

        ws["A3"] = f"巡查帖子总数: {len(data.get('posts', []))}"
        ws.merge_cells("A3:J3")

        headers = [
            "序号", "帖子标题", "Flair", "作者", "点赞数",
            "评论数", "Upvote Ratio", "帖子链接", "评论摘要(Top5)", "讨论内容总结"
        ]
        header_row = 5
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        posts = data.get("posts", [])
        for i, post in enumerate(posts, 1):
            r = header_row + i
            comments = post.get("comments", [])
            comment_summary = "\n---\n".join(
                f"[{c['author']}] (score:{c['score']}): {c['body'][:150]}"
                for c in comments[:5]
            )

            ws.cell(row=r, column=1, value=i).border = thin_border
            ws.cell(row=r, column=2, value=post["title"]).alignment = wrap_align
            ws.cell(row=r, column=3, value=post.get("link_flair_text", ""))
            ws.cell(row=r, column=4, value=post["author"])
            ws.cell(row=r, column=5, value=post["score"])
            ws.cell(row=r, column=6, value=post["num_comments"])
            ws.cell(row=r, column=7, value=post.get("upvote_ratio", 0))
            ws.cell(row=r, column=8, value=post["url"])
            ws.cell(row=r, column=9, value=comment_summary).alignment = wrap_align
            ws.cell(row=r, column=10, value=post.get("selftext", "")[:200]).alignment = wrap_align

            for col in range(1, 11):
                ws.cell(row=r, column=col).border = thin_border

        col_widths = {
            "A": 6, "B": 50, "C": 15, "D": 15, "E": 10,
            "F": 10, "G": 12, "H": 40, "I": 60, "J": 40
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        ws.freeze_panes = "A6"

    wb.save(filepath)
    print(f"\nReport saved: {filepath}")
    return filepath


def main():
    print(f"=== Reddit Community Monitor ===")
    now_bjt = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    print(f"Time: {now_bjt.strftime('%Y-%m-%d %H:%M:%S')} BJT")
    print()

    all_data = {}

    for game, subreddit_names in SUBREDDITS.items():
        print(f"\n--- Fetching {game} (subreddits: {subreddit_names}) ---")
        posts, actual_sub = get_hot_posts(subreddit_names, POSTS_PER_SUB)

        # Get top comments for each post
        for i, post in enumerate(posts):
            print(f"  [{i+1}/{len(posts)}] {post['title'][:60]}...")
            permalink = post["url"].replace("https://www.reddit.com", "").replace("https://old.reddit.com", "")
            if permalink:
                comments = get_top_comments(permalink, COMMENTS_PER_POST)
                post["comments"] = comments
                time.sleep(1)
            else:
                post["comments"] = []

        all_data[game] = {
            "subreddit": actual_sub,
            "posts": posts,
            "summary": summarize_discussion(posts),
        }

        print(f"  Got {len(posts)} posts from r/{actual_sub}")
        time.sleep(2)

    # Generate Excel
    output_dir = os.environ.get("OUTPUT_DIR", ".")
    os.makedirs(output_dir, exist_ok=True)
    filepath = create_excel_report(all_data, output_dir)

    # Also save raw JSON
    json_path = os.path.join(output_dir, f"reddit_raw_{now_bjt.strftime('%Y%m%d_%H%M')}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"Raw data saved: {json_path}")

    print("\n=== Done ===")


if __name__ == "__main__":
    main()
