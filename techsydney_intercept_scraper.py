"""
TechSydney Network-Intercept Scraper
======================================
Smarter version: instead of parsing HTML, we intercept the actual API
calls the browser makes when visiting each startup page. This gives us
the exact same JSON the React app uses — including contact details.

How it works:
  1. Navigate to /s/{slug}
  2. Intercept all XHR/fetch to api.ramenlife.co
  3. Capture: startup detail, team profiles, any chat/contact data
  4. Click "Contact" button → intercept the session creation request
  5. Navigate to each /u/{slug} → intercept profile API response

This is the most reliable approach since we capture real API responses.

Install:
    pip install playwright pandas openpyxl
    playwright install chromium

Run:
    python techsydney_intercept_scraper.py
"""

import asyncio
import csv
import json
import logging
import os
import re
import time
from dataclasses import dataclass, asdict, field
from typing import Optional

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)

# ── Config ──────────────────────────────────────────────────────────────────
ACCESS_TOKEN = "aa4ca1801627b067895c9bd5f7ab42b819f23b37"
COMMUNITY_ID = "27"
SITE_URL     = "https://www.techsydney.com.au"
API_HOST     = "api.ramenlife.co"
PAGE_LIMIT   = 50

NAV_TIMEOUT    = 25_000   # ms
IDLE_TIMEOUT   = 8_000    # ms to wait for network idle
CONTACT_WAIT   = 3_000    # ms after clicking Contact
BETWEEN_PAGES  = 1.2      # seconds between page navigations

OUTPUT_STARTUPS = "techsydney_startups.csv"
OUTPUT_FOUNDERS = "techsydney_founders.csv"
OUTPUT_XLSX     = "techsydney_all.xlsx"


# ── Data classes ────────────────────────────────────────────────────────────

@dataclass
class FounderRow:
    startup_id:        str = ""
    startup_name:      str = ""
    startup_slug:      str = ""
    startup_website:   str = ""
    profile_id:        str = ""
    user_id:           str = ""
    profile_slug:      str = ""
    name:              str = ""
    tagline:           str = ""
    role:              str = ""
    bio:               str = ""
    email_md5:         str = ""
    linkedin_raw:      str = ""
    linkedin_full:     str = ""
    twitter:           str = ""
    github:            str = ""
    instagram:         str = ""
    website:           str = ""
    city:              str = ""
    country:           str = ""
    skills:            str = ""
    looking_to:        str = ""
    is_current_founder:str = ""
    is_former_founder: str = ""
    points:            str = ""
    joined_date:       str = ""
    last_active_date:  str = ""
    profile_url:       str = ""
    # Enriched by contact-btn interception
    contact_btn_name:  str = ""
    contact_btn_role:  str = ""
    page_emails_found: str = ""


@dataclass
class StartupRow:
    id:                 str = ""
    name:               str = ""
    slug:               str = ""
    tagline:            str = ""
    description:        str = ""
    status:             str = ""
    stage:              str = ""
    city:               str = ""
    region:             str = ""
    country:            str = ""
    postcode:           str = ""
    website:            str = ""
    linkedin_raw:       str = ""
    linkedin_full:      str = ""
    twitter:            str = ""
    facebook:           str = ""
    instagram:          str = ""
    github:             str = ""
    staff_count:        str = ""
    founded_year:       str = ""
    funding:            str = ""
    funding_stage:      str = ""
    funding_types:      str = ""
    currently_raising:  str = ""
    customer_focuses:   str = ""
    business_types:     str = ""
    industries:         str = ""
    looking_fors:       str = ""
    skills_needed:      str = ""
    currently_hiring:   str = ""
    hub_resident:       str = ""
    hub_space:          str = ""
    female_founded:     str = ""
    indigenous_founded: str = ""
    social_enterprise:  str = ""
    selling_intl:       str = ""
    outside_australia:  str = ""
    points:             str = ""
    created_date:       str = ""
    updated_date:       str = ""
    founders_names:     str = ""
    founders_roles:     str = ""
    founders_linkedin:  str = ""
    founders_profiles:  str = ""
    # From page/contact-btn
    contact_person_name:str = ""
    contact_person_role:str = ""
    contact_profile_url:str = ""
    page_emails:        str = ""
    page_phones:        str = ""
    all_external_links: str = ""


# ── Helpers ─────────────────────────────────────────────────────────────────

def flat(val) -> str:
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, list):
        return " | ".join(str(x) for x in val if x)
    if isinstance(val, dict):
        return json.dumps(val)
    return str(val)


def resolve_linkedin(raw: str, kind: str = "company") -> str:
    if not raw:
        return ""
    raw = raw.strip().lstrip("/")
    if raw.startswith("http"):
        return raw
    if "linkedin.com" in raw:
        return "https://" + raw
    if kind == "company":
        return f"https://www.linkedin.com/company/{raw}"
    return f"https://www.linkedin.com/in/{raw}"


def extract_desc(data: dict) -> str:
    blocks = data.get("description_djs", {}).get("blocks", [])
    return " ".join(b.get("text", "") for b in blocks).strip()


def extract_emails(text: str) -> list:
    return list(set(re.findall(
        r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text
    )))


def extract_phones(text: str) -> list:
    return list(set(re.findall(
        r"(?:\+?61|0)[2-9]\d{8}|(?:\+?1)?[\s\-.]?\(?\d{3}\)?[\s\-.]?\d{3}[\s\-.]?\d{4}",
        text
    )))[:5]


def build_startup(raw: dict) -> StartupRow:
    s = StartupRow()
    s.id                = str(raw.get("id", ""))
    s.name              = raw.get("name", "")
    s.slug              = raw.get("slug", "")
    s.tagline           = raw.get("tagline", "")
    s.description       = extract_desc(raw)
    s.status            = raw.get("status", "")
    s.stage             = raw.get("stage", "")
    s.city              = raw.get("city", "")
    s.region            = raw.get("region", "")
    s.country           = raw.get("country", "")
    s.postcode          = str(raw.get("postcode", ""))
    s.website           = raw.get("websiteUrl", "")
    s.linkedin_raw      = raw.get("linkedInUrl", "")
    s.linkedin_full     = resolve_linkedin(raw.get("linkedInUrl", ""), "company")
    s.twitter           = raw.get("twitterUrl", "")
    s.facebook          = raw.get("facebookUrl", "")
    s.instagram         = raw.get("instagramUrl", "")
    s.github            = raw.get("githubUrl", "")
    s.staff_count       = raw.get("staffCount", "")
    s.founded_year      = str(raw.get("foundedYear", ""))
    s.funding           = raw.get("funding", "")
    s.funding_stage     = raw.get("fundingStage", "")
    s.funding_types     = flat(raw.get("fundingTypes"))
    s.currently_raising = str(raw.get("currentlyRaising", ""))
    s.customer_focuses  = flat(raw.get("customerFocuses"))
    s.business_types    = flat(raw.get("businessTypes"))
    s.industries        = flat(raw.get("industries"))
    s.looking_fors      = flat(raw.get("lookingFors"))
    s.skills_needed     = flat(raw.get("skillAreasNeeded"))
    s.currently_hiring  = str(raw.get("currentlyHiring", ""))
    s.hub_resident      = str(raw.get("sydneyStartupHubResident", ""))
    s.hub_space         = raw.get("sydneyStartupHubSpace", "")
    s.female_founded    = str(raw.get("femaleFounded", ""))
    s.indigenous_founded= str(raw.get("indigenousFounded", ""))
    s.social_enterprise = str(raw.get("socialEnterprise", ""))
    s.selling_intl      = raw.get("sellingInternationally", "")
    s.outside_australia = str(raw.get("locatedOutsideAustralia", ""))
    s.points            = str(raw.get("points", ""))
    s.created_date      = raw.get("createdDate", "")
    s.updated_date      = raw.get("updatedDate", "")

    team = raw.get("teamProfiles", [])
    s.founders_names    = " | ".join(p.get("nickname","") for p in team if p.get("nickname"))
    s.founders_roles    = " | ".join(p.get("role","") for p in team if p.get("role"))
    s.founders_profiles = " | ".join(
        f"{SITE_URL}/u/{p['slug']}" for p in team if p.get("slug")
    )
    return s


def build_founder(profile: dict, startup: StartupRow, role: str) -> FounderRow:
    f = FounderRow()
    f.startup_id         = startup.id
    f.startup_name       = startup.name
    f.startup_slug       = startup.slug
    f.startup_website    = startup.website
    f.profile_id         = str(profile.get("id", ""))
    f.user_id            = str(profile.get("user_id", ""))
    f.profile_slug       = profile.get("slug", "")
    f.name               = profile.get("nickname", "")
    f.tagline            = profile.get("tagline", "")
    f.role               = role
    f.bio                = extract_desc(profile)
    f.email_md5          = profile.get("emailMd5", "")
    f.linkedin_raw       = profile.get("linkedInUrl", "")
    f.linkedin_full      = resolve_linkedin(profile.get("linkedInUrl",""), "person")
    f.twitter            = profile.get("twitterUrl", "")
    f.github             = profile.get("githubUrl", "")
    f.instagram          = profile.get("instagramUrl", "")
    f.website            = profile.get("websiteUrl", "")
    f.city               = profile.get("city", "")
    f.country            = profile.get("country", "")
    f.is_current_founder = str(profile.get("currentFounder", ""))
    f.is_former_founder  = str(profile.get("formerFounder", ""))
    f.points             = str(profile.get("points", ""))
    f.joined_date        = profile.get("createdDate", "")
    f.last_active_date   = profile.get("lastActiveDate", "")
    f.profile_url        = SITE_URL + profile.get("link", f"/u/{f.profile_slug}")

    # Skills
    skill_map = profile.get("tagsBySkillArea", {})
    skills = []
    for area_tags in skill_map.values():
        skills.extend(t.get("name","") for t in area_tags if t.get("name"))
    f.skills = " | ".join(skills)
    f.looking_to = flat(profile.get("lookingTos"))
    return f


# ── Playwright core ──────────────────────────────────────────────────────────

class NetworkCollector:
    """Collects API responses intercepted from the browser."""

    def __init__(self):
        self.responses: dict[str, dict] = {}   # url_pattern -> parsed JSON

    async def on_response(self, response):
        url = response.url
        if API_HOST not in url:
            return
        try:
            if response.status == 200:
                body = await response.json()
                self.responses[url] = body
        except Exception:
            pass

    def get(self, pattern: str) -> Optional[dict]:
        for url, body in self.responses.items():
            if pattern in url:
                return body
        return None

    def clear(self):
        self.responses.clear()


async def inject_auth(page):
    """Inject auth token so the site treats us as logged in."""
    await page.evaluate(f"""() => {{
        localStorage.setItem('accessToken', '{ACCESS_TOKEN}');
        localStorage.setItem('authToken', '{ACCESS_TOKEN}');
        // Also set in sessionStorage
        sessionStorage.setItem('accessToken', '{ACCESS_TOKEN}');
    }}""")


async def scrape_startup_page(page, collector: NetworkCollector, slug: str) -> dict:
    """
    Visit /s/{slug}, collect intercepted API data + click Contact.
    Returns dict with enrichment data.
    """
    result = {
        "startup_detail": {},
        "team_profiles":  [],
        "contact_person": "",
        "contact_role":   "",
        "contact_profile":"",
        "page_emails":    "",
        "page_phones":    "",
        "external_links": "",
    }

    collector.clear()
    url = f"{SITE_URL}/s/{slug}"

    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await page.wait_for_timeout(2000)
    except Exception as e:
        log.warning(f"  Page load failed {slug}: {e}")
        return result

    # ── Grab intercepted API responses ──
    startup_data = (
        collector.get(f"/startups/{slug}") or
        collector.get(f"/startups?slug={slug}") or
        {}
    )
    if isinstance(startup_data, dict):
        result["startup_detail"] = startup_data

    # ── Page text: emails, phones ──
    try:
        body_text = await page.inner_text("body")
        result["page_emails"] = " | ".join(extract_emails(body_text))
        result["page_phones"] = " | ".join(extract_phones(body_text))
    except Exception:
        pass

    # ── External links ──
    try:
        links = await page.eval_on_selector_all(
            "a[href^='http']",
            f"els => els.map(e=>e.href).filter(h=>!h.includes('{SITE_URL.split('//')[1]}'))"
        )
        clean = [l for l in set(links) if "cloudfront" not in l and "ramenlife" not in l]
        result["external_links"] = " | ".join(clean[:30])
    except Exception:
        pass

    # ── Click Contact button ──
    contact_selectors = [
        "button:has-text('Contact')",
        "a:has-text('Contact')",
        "[class*='contact']",
        "button.btn-primary",
        ".open-cta button",
    ]

    for sel in contact_selectors:
        try:
            btn = page.locator(sel).first
            if await btn.is_visible(timeout=2000):
                # Set up response listener before clicking
                collector.clear()
                await btn.click()
                await page.wait_for_timeout(CONTACT_WAIT)

                # Try to read whatever modal/panel appeared
                modal_text = ""
                for msel in [
                    "[role='dialog']", ".modal-body", ".modal",
                    ".chat-panel", ".sidebar", ".contact-card",
                    "[class*='modal']", "[class*='dialog']",
                    "[class*='chat']", "[class*='message']",
                ]:
                    try:
                        el = page.locator(msel).first
                        if await el.is_visible(timeout=1000):
                            modal_text = await el.inner_text()
                            break
                    except Exception:
                        continue

                if modal_text:
                    lines = [l.strip() for l in modal_text.splitlines() if l.strip()]
                    if lines:
                        result["contact_person"] = lines[0]
                    if len(lines) > 1:
                        result["contact_role"] = lines[1]

                # Try to grab profile link from modal
                try:
                    plinks = await page.eval_on_selector_all(
                        "a[href*='/u/']",
                        "els => els.map(e=>e.href)"
                    )
                    if plinks:
                        result["contact_profile"] = plinks[0]
                except Exception:
                    pass

                # Check if a chat-session API call was intercepted
                chat = collector.get("/chat-sessions")
                if chat and isinstance(chat, dict):
                    profiles_in_chat = chat.get("profiles", [])
                    if profiles_in_chat:
                        p = profiles_in_chat[0]
                        result["contact_person"] = result["contact_person"] or p.get("nickname","")
                        result["contact_role"]   = result["contact_role"] or p.get("tagline","")
                        result["contact_profile"]= (
                            result["contact_profile"] or
                            SITE_URL + p.get("link", "")
                        )

                break
        except Exception:
            continue

    return result


async def scrape_profile_page(page, collector: NetworkCollector, slug: str) -> dict:
    """Visit /u/{slug} and collect intercepted profile API response + page data."""
    result = {"profile": {}, "emails": "", "social_links": ""}

    if not slug:
        return result

    collector.clear()
    try:
        await page.goto(
            f"{SITE_URL}/u/{slug}",
            wait_until="domcontentloaded",
            timeout=NAV_TIMEOUT,
        )
        await page.wait_for_timeout(1500)
    except Exception:
        return result

    # Intercepted profile data
    profile_data = (
        collector.get(f"/profiles/{slug}") or
        collector.get(f"/profiles?slug={slug}") or
        {}
    )
    result["profile"] = profile_data if isinstance(profile_data, dict) else {}

    # Page text
    try:
        body = await page.inner_text("body")
        result["emails"] = " | ".join(extract_emails(body))
    except Exception:
        pass

    # Social links visible on page
    try:
        slinks = await page.eval_on_selector_all(
            "a[href*='linkedin'], a[href*='twitter'], a[href*='github'],"
            "a[href*='instagram'], a[href*='facebook']",
            "els => els.map(e=>e.href)"
        )
        result["social_links"] = " | ".join(set(slinks))
    except Exception:
        pass

    return result


# ── Main orchestration ───────────────────────────────────────────────────────

async def main():
    from playwright.async_api import async_playwright

    startups_out: list[StartupRow] = []
    founders_out:  list[FounderRow] = []
    seen_profiles: set = set()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            extra_http_headers={
                "Authorization": f"Bearer {ACCESS_TOKEN}",
            },
        )

        collector = NetworkCollector()
        page = await context.new_page()
        page.on("response", collector.on_response)

        # ── Warm up: inject auth + get startup list ──────────────────────
        log.info("Warming up browser & fetching startup list...")
        await page.goto(SITE_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await inject_auth(page)
        await page.wait_for_timeout(1000)

        # Navigate to /startups to trigger the API call for the full list
        all_startup_slugs = []
        all_startup_raws  = []

        page_num = 1
        while True:
            collector.clear()
            filter_qs = json.dumps({}) if not {} else json.dumps({})
            nav_url = f"{SITE_URL}/startups?page={page_num}"
            await page.goto(nav_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                # Wait for React to render the startup list
            await page.wait_for_function(
            "() => document.querySelectorAll('a[href*=\"/s/\"]').length > 0",
                timeout=10000
                )

            # Find intercepted list response
            list_data = collector.get("/startups")
            items = []
            
            if not list_data:
                # Try reading from __INITIAL_STATE__ embedded in page
                try:
                    state_str = await page.evaluate(
                        "() => JSON.stringify(window.__INITIAL_STATE__?.startups?.models || {})"
                    )
                    models = json.loads(state_str)
                    if models:
                        items = list(models.values())
                        for item in items:
                            all_startup_raws.append(item)
                            all_startup_slugs.append(item.get("slug",""))
                        log.info(f"  List page {page_num}: got {len(items)} from __INITIAL_STATE__, total so far {len(all_startup_raws)}")
                    else:
                        break
                except Exception as e:
                    log.error(f"  Failed to get page {page_num}: {e}")
                    break
            else:
                # Got data from API interception
                if isinstance(list_data, list):
                    items = list_data
                elif isinstance(list_data, dict):
                    items = list_data.get("startups") or list_data.get("data") or []
                
                for item in items:
                    all_startup_raws.append(item)
                    all_startup_slugs.append(item.get("slug",""))
                
                log.info(f"  List page {page_num}: got {len(items)} from API, total so far {len(all_startup_raws)}")

            # Check if we got fewer items than limit (indicates last page)
            if len(items) < PAGE_LIMIT:
                break
            
            page_num += 1
            await asyncio.sleep(0.5)  # Small delay between pages

        log.info(f"Got {len(all_startup_raws)} startups from list. Now visiting each...")

        # Build initial rows from list data
        slug_to_row: dict[str, StartupRow] = {}
        for raw in all_startup_raws:
            row = build_startup(raw)
            slug_to_row[row.slug] = row
            startups_out.append(row)

        # ── Visit each startup page ──────────────────────────────────────
        total = len(startups_out)
        for i, startup_row in enumerate(startups_out, 1):
            if not startup_row.slug:
                continue

            log.info(f"  [{i}/{total}] /s/{startup_row.slug}")
            enrich = await scrape_startup_page(page, collector, startup_row.slug)

            # Merge any richer detail from intercepted response
            if enrich["startup_detail"]:
                richer = build_startup(enrich["startup_detail"])
                # Overwrite fields that are richer
                for fname in ["description", "team_profiles"]:
                    pass
                if richer.description:
                    startup_row.description = richer.description

            startup_row.contact_person_name = enrich["contact_person"]
            startup_row.contact_person_role = enrich["contact_role"]
            startup_row.contact_profile_url = enrich["contact_profile"]
            startup_row.page_emails         = enrich["page_emails"]
            startup_row.page_phones         = enrich["page_phones"]
            startup_row.all_external_links  = enrich["external_links"]

            # ── Collect team members from intercepted data ──
            raw_team_profiles = (
                enrich["startup_detail"].get("teamProfiles", []) if enrich["startup_detail"]
                else []
            )
            # Fallback: use list-level team data
            if not raw_team_profiles:
                matching_raw = next(
                    (r for r in all_startup_raws if str(r.get("id","")) == startup_row.id), {}
                )
                raw_team_profiles = matching_raw.get("teamProfiles", [])

            for member in raw_team_profiles:
                profile_slug = member.get("slug", "")
                profile_id   = str(member.get("id", ""))
                role         = member.get("role", "member")

                key = profile_id or profile_slug
                if key in seen_profiles:
                    continue
                seen_profiles.add(key)

                # Visit profile page to get intercepted full profile
                log.info(f"    Profile: /u/{profile_slug}")
                p_enrich = await scrape_profile_page(page, collector, profile_slug)

                profile_data = p_enrich["profile"] or {
                    "id":       profile_id,
                    "user_id":  str(member.get("user_id","")),
                    "slug":     profile_slug,
                    "nickname": member.get("nickname",""),
                    "tagline":  member.get("tagline",""),
                    "emailMd5": member.get("emailMd5",""),
                    "link":     f"/u/{profile_slug}",
                }

                f_row = build_founder(profile_data, startup_row, role)

                # Enrich with page-scraped data
                if p_enrich["emails"]:
                    f_row.page_emails_found = p_enrich["emails"]
                if p_enrich["social_links"] and not f_row.linkedin_full:
                    for lnk in p_enrich["social_links"].split(" | "):
                        if "linkedin" in lnk:
                            f_row.linkedin_full = lnk
                            break

                # Update startup's founders_linkedin summary
                if f_row.linkedin_full:
                    if startup_row.founders_linkedin:
                        startup_row.founders_linkedin += " | " + f_row.linkedin_full
                    else:
                        startup_row.founders_linkedin = f_row.linkedin_full

                founders_out.append(f_row)

            await asyncio.sleep(BETWEEN_PAGES)

        await browser.close()

    # ── Save ─────────────────────────────────────────────────────────────
    def save_csv(rows, path, cls):
        if not rows:
            return
        fieldnames = list(asdict(cls()).keys())
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                w.writerow(asdict(r))
        log.info(f"Saved {len(rows)} rows → {path}")

    save_csv(startups_out, OUTPUT_STARTUPS, StartupRow)
    save_csv(founders_out, OUTPUT_FOUNDERS, FounderRow)

    try:
        import pandas as pd
        with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
            pd.DataFrame([asdict(s) for s in startups_out]).to_excel(
                writer, sheet_name="Startups", index=False)
            pd.DataFrame([asdict(f) for f in founders_out]).to_excel(
                writer, sheet_name="Founders", index=False)
        log.info(f"Saved Excel → {OUTPUT_XLSX}")
    except ImportError:
        log.warning("pandas/openpyxl not installed – Excel skipped")

    log.info("=" * 60)
    log.info(f"DONE | {len(startups_out)} startups | {len(founders_out)} founders")


if __name__ == "__main__":
    asyncio.run(main())