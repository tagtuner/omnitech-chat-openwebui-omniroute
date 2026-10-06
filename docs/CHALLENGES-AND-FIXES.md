# Challenges & fixes

Battle log from bringing **Open WebUI + OmniRoute** to a stable on-prem production feel (OmniTech Chat).

## 1) Front door product choice

**Challenge:** First attempt used LibreChat. Model catalog fetch from a large OmniRoute `/v1/models` list timed out; Agents marked models “not available”; empty replies when UI follow-up auto-generation fought tool flows.

**Fix:**
- Switched non-tech front door to **Open WebUI**.
- Keep **Follow-Up Auto-Generation (chips) OFF**.
- Use a dedicated **File Job Follow-up** filter (max one clarification round) instead of chip spam.
- LibreChat stack purged after cutover; OmniRoute stayed shared with Continue.

## 2) Office file downloads broken

**Challenge:** Tools wrote files that surfaced as `sandbox:/cache/files/...` links users could not download cleanly.

**Fix:** Tools upload/save through Open WebUI **Files API** and return `/api/v1/files/{id}/content`. Prefer sync DB/upload path when async handlers fail from tool context.

## 3) Non-admin cannot see Generate* tools

**Challenge:** Admin saw all five tools; user `ghazala` only saw Web Search + Code Interpreter.

**Root cause:** Tools owned by admin; no `access_grant` rows → OWUI filters tools by owner + grants.

**Fix:** Insert `access_grant` with `resource_type=tool`, `permission=read` for:
- `principal_type=user`, `principal_id=*` (public read), and/or
- explicit user id  

Then hard refresh → New Chat → Tools grid.

## 4) Web search not available

**Challenge:** Feature off in DB/compose; engine empty.

**Fix (durable):**
```yaml
ENABLE_WEB_SEARCH: "true"
WEB_SEARCH_ENGINE: "duckduckgo"
WEB_SEARCH_RESULT_COUNT: "5"
USER_PERMISSIONS_FEATURES_WEB_SEARCH: "true"
ENABLE_SEARCH_QUERY_GENERATION: "true"
ENABLE_FOLLOW_UP_GENERATION: "false"   # keep OFF
```
DuckDuckGo needs **no API key**. Smoke from container with `ddgs`.

## 5) Host instability (dirty reboots)

**Challenge:** Dell Precision T5610 — MCE / CMCI storms; overnight hangs; stack looked “randomly broken” after reboot.

**Fix:**
- Remove **mixed 4GB + 16GB** DIMMs (mismatch was the killer).
- Prefer matched size/speed/(vendor); balance across CPUs when scaling RAM.
- `rasdaemon` + `mce_watch` cron for early signal.
- **softdog** + systemd `RuntimeWatchdogSec` for soft hang recovery (no IPMI).
- Document: soft heal recovers apps after reboot; wait 1–2 minutes before blaming config.

## 6) Boot / recreate amnesia

**Challenge:** After host reboot or `docker compose up --force-recreate`, env/network links drifted (especially older LibreChat+proxy era).

**Fix pattern:**
- Pin critical flags in **compose `environment:`** (not only UI toggles).
- Optional oneshot **heal** unit after Docker starts.
- Re-`pip install` tool deps inside OWUI after image recreate.
- Do not steal OmniConnect Redis **6379** for other compose stacks.

## 7) OmniRoute vs UI timeout confusion

**Challenge:** Long spinner blamed on “WebUI hang”; actually upstream model/provider wait (e.g. slow coding model).

**Fix:** Teach users to pick `auto/best-fast` for research / chat; reserve heavier profiles for coding. Check OmniRoute request logs (HTTP 200 vs error counters).

## 8) Signup / key hygiene (deferred harden)

**Challenge:** Feel-test needed open signup briefly; keys must not land in git.

**Fix:**
- Secrets only in host `.env` (mode 600).
- Disable signup after named users exist.
- Separate restricted key for UI client vs full admin/Continue key.

## 9) Incomplete Excel from RAG chunks (FULL_FILE_READ)

**Challenge:** Attached survey spreadsheet → Subject-wise Excel showed only **2 subjects**. Model used RAG “1 source” snippets and invented rows.

**Fix:**
- Tools read the **complete `.xlsx` from disk** (`lib_source_workbook.py`) — never chat RAG invent.
- Generate Excel / Word / PDF / Slides: `job=subject_group_report` (+ `source_file_id` / `__files__`).
- Generate Forecast Pack: `job=forecast_from_source` (prefer sheets **AS GB + A2 GB**).
- File Job Follow-up filter forces the job= path when a sheet is attached.
- Deploy helpers: `install_full_source_tools.py`, `install_forecast_pack_full_source.py`.
- Acceptance: survey VERIFY **105 rows / 11 subjects**; Forecast Pack smoke with AS+A2 GB.

## 10) Same-chat “cooling down” / Maximum combo

**Challenge:** After a successful tool reply, the **next** message in the same chat failed with all credentials cooling down / combo retry — users thought they must open a new chat.

**Root cause:** Virtual `auto/best-fast` hunted dozens of dead upstream models (403/400/429 / no-balance), burned lockouts, then failed the request.

**Fix:**
- Keep only proven providers **ON** (e.g. gemini, alibaba, opencode, cloudflare-ai, deepseek, grok-cli).
- Force **OFF** burners (zai/github/huggingface/groq/mistral/web CLIs that 403, etc.).
- Soften `modelLockout` (`baseCooldownMs` / `maxCooldownMs`) so one 429 does not brick the thread for minutes.
- Script: `deploy/harden_bestfast_combo.py`.
- Teach: **same chat regenerate/retry** — do not require a new chat.

---

## Verification gates (definition of “flawless”)

- [ ] WebUI `:3080` healthy; OmniRoute `:20128` `/v1/models` OK  
- [ ] Non-admin sees Generate* tools after ACL  
- [ ] Web Search returns cited answers  
- [ ] Forecast / office tool returns downloadable Files API link  
- [ ] Attached survey → FULL_FILE_READ VERIFY (complete subjects/rows, not RAG invent)  
- [ ] Same-chat follow-up after tools works (combo harden; no forced new chat)  
- [ ] Chip follow-ups remain OFF  
- [ ] `mce_watch` quiet after RAM hygiene  
- [ ] Survive reboot + short heal wait without manual rewire  

---

*Captured 2026-08 … updated 2026-10 — OmniTech production notes. Sanitize hostnames/IPs/user ids before publishing customer-specific forks.*
