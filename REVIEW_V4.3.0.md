# V4.3.0 Script Review

## Critical bugs / blockers
1. **Notion configuration dialog cannot read or save settings.** The HTML calls `google.script.run.saveNotionConfig(...)` and `readNotionConfig()`, but no matching server-side functions exist in `Code_V4.3.0.gs`, so the dialog will throw runtime errors and the Notion token/DB map can never be persisted.【F:NotionConfig_V4.3.0.html†L44-L78】【F:Code_V4.3.0.gs†L542-L546】
2. **Document styling regression.** `styleParagraph` only honours the option `family`, yet all new rendering code passes `fontFamily`, so paragraph fonts never change from the default. This breaks student/answer key formatting in the generated Docs.【F:Code_V4.3.0.gs†L112-L139】【F:Code_V4.3.0.gs†L404-L430】【F:Code_V4.3.0.gs†L1076-L1084】
3. **Notion filter type mismatch.** `buildNotionApiFilter` always uses a `rich_text` filter for `Topic/Chapter`. If the Notion property is configured as `select` or `multi_select` (a supported option in previous releases), the query will 400 with “property does not support rich_text filter,” halting Sync Down. Need to detect property type or fall back to client-side filtering.【F:Code_V4.3.0.gs†L784-L788】

## Reliability risks
1. **High chance of Notion API rate-limit errors.** Sync Down and indexers pause only 80–120 ms between requests, and Sync Up has no throttling at all. Notion limits to ~3 requests/sec; these loops can easily exceed that and trigger 429 errors without retry logic.【F:Code_V4.3.0.gs†L645-L756】【F:Code_V4.3.0.gs†L692-L704】【F:Code_V4.3.0.gs†L1027-L1040】 Add ≥350 ms delays, request batching, and exponential backoff on 429.
2. **`lastEdited` accepts invalid dates.** Invalid strings are sent directly to Notion, which returns 400 errors. Validate format before building the API filter.【F:Code_V4.3.0.gs†L796-L799】
3. **Course-to-database heuristic can misroute.** `courseToDbId` falls back to the first prefix match. If two courses share the same prefix (e.g., “BIO 1130” and “BIOC 2001”), rows may be synced to the wrong DB. Consider storing a lower-case map and requiring exact matches before prefix fallbacks.【F:Code_V4.3.0.gs†L824-L833】

## Performance opportunities
1. **Batch Import updates.** `runSyncDownWithFilters` updates rows one at a time (`setValues`/`appendRow`). Accumulate updates and push them in blocks to drastically cut Apps Script calls.【F:Code_V4.3.0.gs†L632-L756】
2. **Avoid repeated allocations during filtering.** Precompute lowercase lists for `diffList`, `typeList`, `corrList`, and reuse instead of calling `.map(toLower)` for every row.【F:Code_V4.3.0.gs†L241-L315】
3. **Reuse Notion auth config.** `notionSyncUp` calls `requireNotionConfig()` for every row, causing redundant property lookups. Fetch once per sync.【F:Code_V4.3.0.gs†L685-L704】

## Suggested improvements / enhancements
- **Implement the missing Notion config service.** Add `readNotionConfig()`/`saveNotionConfig()` that persist `NOTION_TOKEN` and the DB map via `PropertiesService`, plus basic validation and success/error messaging.【F:NotionConfig_V4.3.0.html†L44-L78】
- **Harden Notion filter builder.** Determine property types (e.g., via `/databases/:id` metadata) so tag/topic filters use the correct predicate, or gracefully fall back to client-side filtering when unsupported.【F:Code_V4.3.0.gs†L784-L794】
- **Add retry/backoff wrapper around `notionRequest`.** Catch 429/5xx responses, sleep, and retry a limited number of times to keep long syncs resilient.【F:Code_V4.3.0.gs†L725-L738】
- **Expose filter presets.** Allow saving/loading common Sync Down filter sets (per course, per tag) for faster workflows—leveraging the planned config service.
- **Optional: progress reporting for Sync Down.** As multiple databases are queried, surface incremental progress to the user via `toast` or UI callbacks to avoid uncertainty during long pulls.

Addressing the critical bugs above is required before shipping V4.3.0.
