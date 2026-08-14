#!/usr/bin/env python3
"""
Reddit Community Monitor - Daily Patrol Report Generator
Fetches hot posts and top comments from game subreddits and generates Excel reports.
"""

import json
import os
import sys
import time
import urllib.request
import urllib.error
import datetime
import ssl

# For Excel generation
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

# Game subreddits to monitor
SUBREDDITS = {
    "Palworld": "Palworld",
    "CS2": "GlobalOffensive",
    "Valorant": "Valorant",
    "LOL": "leagueoflegends",
    "DeltaForce": "DeltaForce",
    "TFT": "TeamfightTactics",
}

# How many hot posts to fetch per subreddit
POSTS_PER_SUB = 20
# How many top comments to fetch per post
COMMENTS_PER_POST = 10

# SSL context that doesn't verify (for environments with cert issues)
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; RedditMonitor/1.0; +https://github.com/reddit-monitor)"
}


def fetch_json(url, retries=3, delay=2):
    """Fetch JSON from URL with retries."""
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            print(f"  Attempt {attempt+1}/{retries} failed: {e}", file=sys.stderr)
            if attempt < retries - 1:
                time.sleep(delay * (attempt + 1))
    return None


def get_hot_posts(subreddit, limit=POSTS_PER_SUB):
    """Get hot posts from a subreddit."""
    url = f"https://www.reddit.com/r/{subreddit}/hot.json?limit={limit}"
    print(f"Fetching: {url}")
    data = fetch_json(url)
    if not data or "data" not in data or "children" not in data["data"]:
        print(f"  Failed to fetch posts from r/{subreddit}")
        return []

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
    return posts


def get_top_comments(permalink, limit=COMMENTS_PER_POST):
    """Get top comments from a post."""
    url = f"https://www.reddit.com{permalink}.json?limit={limit}&sort=top"
    data = fetch_json(url)
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

    # Group by flair if available
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

        # Summarize top topics
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
        ws_summary[f"A{row}"] = f"【{game}】r/{SUBREDDITS.get(game, game)}"
        ws_summary[f"A{row}"].font = Font(bold=True)
        row += 1
        for p in posts[:5]:
            ws_summary[f"A{row}"] = f"  - {p['title']} (↑{p['score']} 💬{p['num_comments']})"
            ws_summary[f"A{row}"].alignment = wrap_align
            row += 1
        row += 1

    # Adjust column widths
    ws_summary.column_dimensions["A"].width = 40
    ws_summary.column_dimensions["B"].width = 15
    ws_summary.column_dimensions["C"].width = 15
    ws_summary.column_dimensions["D"].width = 60

    # === Sheet per game ===
    for game, data in all_data.items():
        ws = wb.create_sheet(title=game[:31])

        # Title
        sub_name = SUBREDDITS.get(game, game)
        ws["A1"] = f"{now.strftime('%m月%d日 %H:%M')} r/{sub_name} 数据"
        ws["A1"].font = title_font
        ws.merge_cells("A1:J1")

        ws["A3"] = f"巡查帖子总数: {len(data.get('posts', []))}"
        ws.merge_cells("A3:J3")

        # Header row
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

        # Data rows
        posts = data.get("posts", [])
        for i, post in enumerate(posts, 1):
            r = header_row + i

            # Fetch comments
            comments = post.get("comments", [])
            comment_summary = "\n---\n".join(
                f"[{c['author']}] (↑{c['score']}): {c['body'][:150]}"
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

        # Column widths
        col_widths = {
            "A": 6, "B": 50, "C": 15, "D": 15, "E": 10,
            "F": 10, "G": 12, "H": 40, "I": 60, "J": 40
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        # Freeze header
        ws.freeze_panes = "A6"

    wb.save(filepath)
    print(f"\nReport saved: {filepath}")
    return filepath


def main():
    print(f"=== Reddit Community Monitor ===")
    print(f"Time: {datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S')} BJT")
    print()

    all_data = {}

    for game, subreddit in SUBREDDITS.items():
        print(f"\n--- Fetching r/{subreddit} ({game}) ---")
        posts = get_hot_posts(subreddit, POSTS_PER_SUB)

        # Get top comments for each post
        for i, post in enumerate(posts):
            print(f"  [{i+1}/{len(posts)}] {post['title'][:60]}...")
            permalink = post["url"].replace("https://www.reddit.com", "")
            comments = get_top_comments(permalink, COMMENTS_PER_POST)
            post["comments"] = comments
            time.sleep(0.5)  # Rate limit

        all_data[game] = {
            "subreddit": subreddit,
            "posts": posts,
            "summary": summarize_discussion(posts),
        }

        print(f"  Got {len(posts)} posts")

    # Generate Excel
    output_dir = os.environ.get("OUTPUT_DIR", ".")
    filepath = create_excel_report(all_data, output_dir)

    # Also save raw JSON
    json_path = os.path.join(output_dir, f"reddit_raw_{datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y%m%d_%H%M')}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"Raw data saved: {json_path}")

    print("\n=== Done ===")


if __name__ == "__main__":
    main()
