"""
Divar Apartment Sales Scraper (Tehran)
Scrapes apartment-for-sale listings from Divar.ir using the public API.
Saves results to an Excel file.
"""

import requests
import time
import pandas as pd
from typing import List, Dict, Any, Optional


# ---------------------------- Configuration ----------------------------
SEARCH_URL = "https://api.divar.ir/v8/postlist/w/search"
POST_URL_TEMPLATE = "https://api.divar.ir/v8/posts-v2/web/{token}"
CITY_ID = "1"  # Tehran
CATEGORY = "apartment-sell"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
    "Accept": "application/json",
}
REQUEST_DELAY = 0.7  # seconds between detail requests (avoid 429)
OUTPUT_FILE = "Divar_houses_data.xlsx"


def get_initial_pagination(session: requests.Session) -> tuple[Optional[str], Optional[str]]:
    """Fetch first page to obtain last_post_date and search_uid."""
    payload = {
        "city_ids": [CITY_ID],
        "source_view": "CATEGORY",
        "search_data": {
            "form_data": {
                "data": {
                    "category": {"str": {"value": CATEGORY}}
                }
            }
        },
    }
    try:
        resp = session.post(SEARCH_URL, headers=HEADERS, json=payload, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        pagination = data.get("pagination", {}).get("data", {})
        return pagination.get("last_post_date"), pagination.get("search_uid")
    except Exception as e:
        print(f"[ERROR] Could not get initial pagination: {e}")
        print("Check your internet connection / VPN (Divar is often blocked outside Iran).")
        return None, None


def collect_tokens(
    session: requests.Session,
    number: int,
    last_post_date: str,
    search_uid: str,
) -> List[str]:
    """Paginate through search results and collect post tokens."""
    tokens: List[str] = []
    page = 1

    while len(tokens) < number:
        payload = {
            "city_ids": [CITY_ID],
            "pagination_data": {
                "@type": "type.googleapis.com/post_list.PaginationData",
                "last_post_date": last_post_date,
                "page": page,
                "layer_page": page,
                "search_uid": search_uid,
            },
            "search_data": {
                "form_data": {
                    "data": {
                        "category": {"str": {"value": CATEGORY}},
                        "sort": {"str": {"value": "sort_date"}},
                    }
                }
            },
        }

        try:
            resp = session.post(SEARCH_URL, headers=HEADERS, json=payload, timeout=15)
            if resp.status_code != 200:
                print(f"[WARN] Search page {page} returned status {resp.status_code}")
                break

            data = resp.json()
            widgets = data.get("list_widgets", [])
            if not widgets:
                print("[INFO] No more listings found.")
                break

            # Update pagination for next page
            pag = data.get("pagination", {}).get("data", {})
            last_post_date = pag.get("last_post_date", last_post_date)
            search_uid = pag.get("search_uid", search_uid)

            for post in widgets:
                try:
                    token = post["data"]["action"]["payload"]["token"]
                    tokens.append(token)
                    if len(tokens) >= number:
                        break
                except (KeyError, TypeError):
                    continue

            print(f"[INFO] Collected {len(tokens)} / {number} tokens (page {page})")
            page += 1
            time.sleep(0.3)

        except Exception as e:
            print(f"[ERROR] Failed on page {page}: {e}")
            break

    return tokens[:number]


def safe_get(d: dict, *keys, default="ندارد"):
    """Safely navigate nested dicts."""
    for key in keys:
        if isinstance(d, dict):
            d = d.get(key, default)
        else:
            return default
    return d if d is not None else default


def extract_ad_details(data: Dict[str, Any]) -> Dict[str, Any]:
    """Parse a single post detail response into a flat dictionary."""
    ad = {
        "محله": "ندارد",
        "طبقه": "ندارد",
        "سال ساخت": "ندارد",
        "تعداد اتاق": "ندارد",
        "متراژ": "ندارد",
        "آسانسور": "ندارد",
        "پارکینگ": "ندارد",
        "انباری": "ندارد",
        "قیمت کل": "ندارد",
        "قیمت هر متر": "ندارد",
        "نقطه عرضی": "ندارد",
        "نقطه طولی": "ندارد",
    }

    # District & price from webengage (more reliable)
    webengage = data.get("webengage", {})
    ad["محله"] = webengage.get("district", "ندارد")
    ad["قیمت کل"] = webengage.get("price", "ندارد")

    # Walk through sections → widgets
    for section in data.get("sections", []):
        for widget in section.get("widgets", []):
            wtype = widget.get("widget_type")
            wdata = widget.get("data", {})

            if wtype == "MAP_ROW":
                point = safe_get(wdata, "location", "exact_data", "point", default={})
                if isinstance(point, dict):
                    ad["نقطه عرضی"] = point.get("latitude", "ندارد")
                    ad["نقطه طولی"] = point.get("longitude", "ندارد")

            elif wtype == "GROUP_INFO_ROW":
                # Contains items like متراژ, سال ساخت, تعداد اتاق, etc.
                for item in wdata.get("items", []):
                    title = item.get("title", "")
                    value = item.get("value", "ندارد")
                    if "متراژ" in title:
                        ad["متراژ"] = value
                    elif "ساخت" in title:
                        ad["سال ساخت"] = value
                    elif "اتاق" in title:
                        ad["تعداد اتاق"] = value

            elif wtype == "UNEXPANDABLE_ROW":
                title = wdata.get("title", "")
                value = wdata.get("value", "ندارد")
                if title == "طبقه":
                    ad["طبقه"] = value
                elif "متر" in title and "قیمت" in title:
                    ad["قیمت هر متر"] = value

            elif wtype == "FEATURE_ROW":
                # Features: آسانسور، پارکینگ، انباری
                for feature in wdata.get("features", []):
                    title = feature.get("title", "")
                    # presence is usually indicated by the feature existing
                    if "آسانسور" in title:
                        ad["آسانسور"] = True
                    elif "پارکینگ" in title:
                        ad["پارکینگ"] = True
                    elif "انباری" in title:
                        ad["انباری"] = True

    # Fallbacks for boolean features (if not found → ندارد)
    for key in ("آسانسور", "پارکینگ", "انباری"):
        if ad[key] is True:
            continue
        if ad[key] == "ندارد":
            # keep ندارد
            pass

    return ad


def fetch_ad_details(session: requests.Session, tokens: List[str]) -> List[Dict[str, Any]]:
    """Fetch full details for each token."""
    ads = []
    for idx, token in enumerate(tokens, 1):
        url = POST_URL_TEMPLATE.format(token=token)
        try:
            resp = session.get(url, headers=HEADERS, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                ad = extract_ad_details(data)
                ads.append(ad)
                print(f"[{idx}/{len(tokens)}] {ad['محله']} | {ad['متراژ']}m | {ad['قیمت کل']}")
            else:
                print(f"[{idx}/{len(tokens)}] Failed – status {resp.status_code}")
        except Exception as e:
            print(f"[{idx}/{len(tokens)}] Error: {e}")

        time.sleep(REQUEST_DELAY)

    return ads


def main():
    print("=" * 60)
    print("  Divar Apartment-Sell Scraper (Tehran)")
    print("=" * 60)

    try:
        number = int(input("How many advertisements do you want to retrieve? "))
        if number <= 0:
            print("Please enter a positive number.")
            return
    except ValueError:
        print("Invalid input. Please enter an integer.")
        return

    session = requests.Session()

    print("\n[1/3] Getting initial search pagination...")
    last_post_date, search_uid = get_initial_pagination(session)
    if not last_post_date or not search_uid:
        return

    print("[2/3] Collecting advertisement tokens...")
    tokens = collect_tokens(session, number, last_post_date, search_uid)
    if not tokens:
        print("No tokens collected. Exiting.")
        return
    print(f"→ Got {len(tokens)} tokens")

    print("\n[3/3] Fetching full advertisement details...")
    ads = fetch_ad_details(session, tokens)

    if ads:
        df = pd.DataFrame(ads)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n✅ Done! {len(ads)} ads saved to → {OUTPUT_FILE}")
    else:
        print("\nNo data collected. Excel file was not created.")


if __name__ == "__main__":
    main()
