# Connectors

## How tool references work

Plugin files use `~~category` as a placeholder for whatever tool the user connects in that category. For example, `~~browser` might mean Playwright, Claude in Chrome, or any other browser automation with an MCP server.

Plugins are **tool-agnostic** — they describe workflows in terms of categories (browser automation, display, etc.) rather than specific products. The `.mcp.json` pre-configures specific MCP servers, but any MCP server in that category works.

## Connectors for this plugin

| Category | Placeholder | Included servers | Other options |
|----------|-------------|-----------------|---------------|
| Browser automation | `~~browser` | Playwright (plugin-playwright) | Claude in Chrome, Puppeteer |
| Display | `~~display` | scrape-display | Custom renderers |
| Export | `~~export` | Built-in (csv, json, xlsx) | Google Sheets, Airtable |

## Browser Automation Details

This plugin requires browser automation for web scraping. Two primary options:

### Playwright (Recommended)

- **Best for**: Job board scraping (Indeed, Stepstone)
- **Why**: Headless operation, robust selectors, fast execution
- **MCP server**: `plugin-playwright`

### Claude in Chrome

- **Best for**: LinkedIn search (requires authentication)
- **Why**: Uses your existing LinkedIn session, handles dynamic content
- **MCP server**: `Claude in Chrome`

## Display Integration

The `display_scraped_data` MCP tool renders job listings in an interactive format:
- Card view with company, title, salary, location
- Action buttons for each job (view, track, export)
- Filtering and sorting controls

## Export Options

Results can be exported to:
- **CSV**: Flat file, CRM-compatible
- **JSON**: Full structured data with nested leads
- **XLSX**: Multi-sheet workbook (Jobs, Leads, Summary)

## Authentication Notes

### Indeed / Stepstone
- No authentication required
- Public job listings are scraped directly

### LinkedIn
- Requires active LinkedIn session in browser
- Claude in Chrome uses your logged-in session
- Rate limits apply (~100 searches/day on free tier)
- Recruiter/Sales Navigator accounts have higher limits
