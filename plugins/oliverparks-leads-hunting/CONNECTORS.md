# Connectors

## How tool references work

Plugin files use `~~category` as a placeholder for whatever tool the user connects in that category. For example, `~~browser` might mean Claude in Chrome or any other browser automation MCP server.

Plugins are **tool-agnostic** — they describe workflows in terms of categories (browser automation, display, etc.) rather than specific products.

## Connectors for this plugin

| Category | Placeholder | Required server | Notes |
|----------|-------------|----------------|-------|
| Browser automation | `~~browser` | Claude in Chrome | Real browser session required — headless browsers are blocked |
| Display | `~~display` | scrape-display | Custom renderers |
| Export | `~~export` | Built-in (csv, json, xlsx) | Google Sheets, Airtable |

## Browser Automation: Claude in Chrome Only

This plugin requires **Claude in Chrome** (or equivalent real-browser MCP server) for all web scraping. Headless browsers like Playwright are **not supported** — Indeed and Stepstone actively detect and block headless/automated browsers.

### Why not Playwright?

- Indeed and Stepstone deploy bot detection that blocks headless browsers outright
- These sites require a visible browser session with normal fingerprints
- Playwright connections will fail with CAPTCHAs or blank pages

### Claude in Chrome

- **Used for**: All scraping — Indeed, Stepstone, and LinkedIn
- **Why**: Real browser session, normal fingerprints, bypasses bot detection
- **LinkedIn bonus**: Uses your existing logged-in session for authenticated searches
- **MCP server**: `Claude in Chrome`

## Export Options

Results can be exported to:
- **CSV**: Flat file, CRM-compatible
- **JSON**: Full structured data with nested leads
- **XLSX**: Multi-sheet workbook (Indeed Jobs, Stepstone Jobs, Hiring Managers, HR & Recruiting)

## Authentication Notes

### Indeed / Stepstone
- No authentication required
- Public job listings scraped via real browser session

### LinkedIn
- Requires active LinkedIn session in browser
- Claude in Chrome uses your logged-in session
- Rate limits apply (~100 searches/day on free tier)
- Recruiter/Sales Navigator accounts have higher limits
