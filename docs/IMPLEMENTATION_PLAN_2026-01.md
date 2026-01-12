# Office-MCP Server: High-ROI Improvements Implementation Plan

> **STATUS: COMPLETED** - All phases 1-5 implemented and verified on 2026-01-12

## Overview

This plan addresses critical security vulnerabilities, API correctness bugs, and codebase cleanup based on thorough analysis of the office-mcp server codebase and PR review.

**Repository:** hvkshetry/office-365-mcp-server
**Completed:** 2026-01-12
**Verified by:** Codex (gpt-5.2-codex)

---

## Phase 1: Critical Security Fixes (HIGHEST PRIORITY)

### 1.1 Remove Token/Credential Logging

**File:** `auth/token-manager.js`

**Problem:** Lines 36, 40-44 log actual token content to stderr:
```javascript
console.error('[DEBUG] Token file first 200 characters:', tokenData.slice(0, 200));
console.error('[DEBUG] Parsed tokens keys:', Object.keys(tokens));
```

**Fix:** Replace with safe logging that only confirms token presence:
- Remove line 36 (first 200 chars)
- Replace lines 40-44 with: `console.error('[DEBUG] Token has access_token:', !!tokens.access_token);`
- Remove debug loop logging token keys/types

**Files affected:** `auth/token-manager.js` (lines 18-44)

---

### 1.2 Secure Token File Permissions

**Files:**
- `auth/token-manager.js:89`
- `office-auth-server.js:305`

**Problem:** Token files written with default permissions (potentially world-readable)

**Fix:** Add mode `0o600` to writeFileSync calls:
```javascript
fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2), { mode: 0o600 });
```

---

### 1.3 Remove Email Content Logging

**File:** `email/index.js`

**Problem:** Multiple console.error calls log full args including email bodies, recipients:
- Lines 153-156, 169-170, 184-186
- Lines 652-656, 743

**Fix:** Create safe logging helper that redacts `body`, `to`, `cc`, `bcc` fields, or remove verbose debug logging entirely.

---

### 1.4 Remove Sensitive URL Logging

**File:** `utils/graph-api.js`

**Problem:** Lines 53, 90, 94 log full API URLs with potential sensitive query params

**Fix:** Gate verbose URL logging behind `DEBUG_VERBOSE=true` env var, or log only path without query string.

---

## Phase 2: API Correctness Fixes (HIGH PRIORITY)

### 2.1 Fix Path Encoding Breaking Graph API

**File:** `utils/graph-api.js:56-58`

**Problem:** Current code encodes ALL path segments:
```javascript
const encodedPath = path.split('/').map(segment => encodeURIComponent(segment)).join('/');
```

This breaks:
- OData operators: `$value`, `$ref`, `$count` become `%24value`
- Drive path syntax: `root:/Documents/file.docx:/content` becomes broken

**Fix:** Selective encoding that preserves Graph API special syntax:
```javascript
function encodeGraphPath(path) {
  if (!path || path.includes('%')) return path;
  return path.split('/').map(segment => {
    // Preserve OData operators and drive path colons
    if (/^\$[a-z]+$/i.test(segment) || segment.startsWith(':')) return segment;
    if (/^[a-zA-Z0-9_-]+$/.test(segment)) return segment;
    return encodeURIComponent(segment).replace(/%3A/g, ':');
  }).join('/');
}
```

---

### 2.2 Fix Binary Upload Corruption

**File:** `utils/graph-api.js:182-184`

**Problem:** Line 183 always JSON.stringifies data, corrupting binary uploads:
```javascript
req.write(JSON.stringify(data));
```

**Fix:** Check Content-Type before serializing:
```javascript
if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
  const contentType = customHeaders['Content-Type'] || 'application/json';
  if (contentType.includes('application/octet-stream') || Buffer.isBuffer(data)) {
    req.write(Buffer.isBuffer(data) ? data : Buffer.from(data));
  } else {
    req.write(JSON.stringify(data));
  }
}
```

---

### 2.3 Fix Search API Entity Type Limits

**File:** `search/index.js`

**Problem:**
- Line 55: Default entity types `['driveItem', 'message', 'event', 'listItem']` are incompatible
- Line 86: Allows `limit: 500` but email search max is 25

**Fix:**
- Change default to compatible types: `entityTypes = ['driveItem', 'listItem']`
- Add limit validation based on entity type:
```javascript
let effectiveLimit = validEntityTypes.some(t => ['message', 'event'].includes(t))
  ? Math.min(limit, 25)  // Email/event limit
  : Math.min(limit, 500);
```

---

## Phase 3: OAuth Configuration Consolidation (MEDIUM)

### 3.1 Unify Scopes

**Files:** `config.js`, `office-auth-server.js`

**Problem:** Scope mismatch - `config.js` missing `offline_access` and other scopes that `office-auth-server.js` requests

**Fix:** Add to `config.js` scopes array:
```javascript
'offline_access',  // CRITICAL for refresh tokens
'User.ReadBasic.All',
'Files.ReadWrite.All',
'Group.Read.All',
'Directory.Read.All',
'Presence.Read',
'Presence.ReadWrite'
```

Then update `office-auth-server.js` to import and use `config.AUTH_CONFIG.scopes`.

---

## Phase 4: Process Management Fix (MEDIUM)

### 4.1 Fix SIGTERM Handling

**File:** `index.js:157-159`

**Problem:** SIGTERM handler keeps process alive, hostile to supervisors

**Fix:** Implement graceful shutdown:
```javascript
process.on('SIGTERM', () => {
  console.error('[SHUTDOWN] SIGTERM received, shutting down gracefully');
  server.close?.();
  setTimeout(() => process.exit(0), 1000);
});
```

---

## Phase 5: Codebase Cleanup (LOW)

### 5.1 Fix Broken npm Script

**File:** `package.json:8`

**Problem:** `start:http` references non-existent `server-http.js`

**Fix:** Remove the broken script line.

---

### 5.2 Remove Development Documentation

**Files to delete or move to `docs/`:**
- `ENHANCED_SEARCH_IMPLEMENTATION_PLAN.md` (35KB) - Development roadmap, Phase 2/3 pending
- `planner/ASSIGNMENT_FIX_SUMMARY.md` (3KB) - Dev notes
- `CONTACTS_API.md` (6KB) - API docs (could consolidate into README)

**Optional:** Add to `.gitignore`:
```
*_IMPLEMENTATION_PLAN.md
*_FIX_SUMMARY.md
```

---

## Critical Files Summary

| File | Changes |
|------|---------|
| `auth/token-manager.js` | Remove token logging (lines 18-44), add file mode 0o600 (line 89) |
| `utils/graph-api.js` | Fix path encoding (lines 56-58), fix binary uploads (lines 182-184), reduce URL logging |
| `email/index.js` | Remove/redact sensitive logging (lines 153-156, 652-656, 743) |
| `config.js` | Add missing scopes including `offline_access` |
| `office-auth-server.js` | Add file mode 0o600 (line 305), use config scopes |
| `search/index.js` | Fix default entity types (line 55), add limit validation |
| `index.js` | Fix SIGTERM handling (lines 157-159) |
| `package.json` | Remove broken `start:http` script |

---

## Verification Steps

### After Phase 1 (Security):
```bash
# Verify no token logging remains
grep -r "first 200\|tokenData.slice\|Parsed tokens keys" auth/
# Check token file permissions after auth
ls -la ~/.office-mcp-tokens.json  # Should show -rw-------
```

### After Phase 2 (API):
```bash
# Test file download (should not error on /$value endpoint)
# Test binary file upload (PDF/image should not be corrupted)
# Test email search with limit > 25 (should cap appropriately)
```

### After Phase 3 (OAuth):
```bash
# Complete fresh authentication
# Wait for token expiry or manually expire
# Verify auto-refresh succeeds with correct scopes
```

### After Phase 4 (Process):
```bash
# Start server, send SIGTERM, verify exit code 0
node index.js &
PID=$!
sleep 2
kill -TERM $PID
echo "Exit code: $?"
```

---

## Implementation Priority

| # | Task | Effort | Impact |
|---|------|--------|--------|
| 1 | Token logging fix | 15 min | Critical security |
| 2 | File permissions | 10 min | Security |
| 3 | Path encoding fix | 30 min | API breaking bug |
| 4 | Binary upload fix | 20 min | API breaking bug |
| 5 | Email log redaction | 20 min | Security |
| 6 | URL log reduction | 10 min | Security |
| 7 | OAuth scope unification | 30 min | Reliability |
| 8 | SIGTERM handling | 10 min | Operations |
| 9 | Search limits | 15 min | API correctness |
| 10 | Package.json fix | 5 min | Cleanup |
| 11 | Dev doc cleanup | 15 min | Maintenance |

---

## Open PRs Note

- **PR #8** (Open): Email contact extraction - Review independently
- **PR #9** (Open): CLAUDE.md + newsletter detector - Review independently
- **PR #3** (Merged): Attempted path encoding fix but issue persists - This plan supersedes

---

## Codex Review Feedback

### Plan Critique (from Codex)
1. **Security logging scope wider than noted** - Token preview/keys logged at `auth/token-manager.js:18,34,43`; raw email args at `email/index.js:153,184`. Fix should be "default off + redacted" across ALL modules.
2. **Token file permissions must cover ALL writers** - `auth/auto-refresh.js:82` also needs fix. Consider atomic writes (write temp → rename) to avoid partial/corrupt files.
3. **URL/query param redaction** - `utils/graph-api.js:90,94` log full query strings with search terms/emails. Should redact even when DEBUG enabled.
4. **Path encoding fragile with regex** - SDK-style parser preferred over regex to handle `root:/path:/content`, `$value`, and already-encoded segments.
5. **Binary handling incomplete** - Responses accumulated as strings at `utils/graph-api.js:106` corrupt downloads. Fix must handle binary responses AND request bodies.
6. **Search should NOT silently drop types** - `search/index.js:300` filters incompatible types. Prefer multiple requests with per-entity limits then merge.

### Additional Codex Recommendations
- Add centralized redaction helper with log levels (DEBUG/INFO) across ALL modules
- Respect `Retry-After` headers for 429/503 and add jitter to avoid synchronized retries
- Clone `queryParams` in `utils/graph-api.js:64` to avoid mutating caller params
- Add tests for path encoding, binary upload/download, and search per-entity limits

---

## Phase 6: High-ROI Graph API Feature Additions (Future)

Based on Codex research of Microsoft Graph API, these features would add significant value:

### 6.1 Delta Sync (HIGH VALUE)
- **Endpoints**: `/me/messages/delta`, `/me/events/delta`, `/me/drive/root/delta`
- **Value**: Enable incremental updates, reduce polling overhead
- **Effort**: Medium

### 6.2 Large File Upload Sessions (HIGH VALUE)
- **Endpoint**: `/createUploadSession`
- **Value**: Support files >4MB, resumable uploads
- **Effort**: Medium

### 6.3 Calendar Scheduling
- **Endpoints**: `getSchedule`, `findMeetingTimes`, Places/rooms API
- **Value**: Resource booking, availability checking
- **Effort**: Low-Medium

### 6.4 Microsoft To Do Integration
- **Endpoints**: `/me/todo/lists`, `/me/todo/lists/{id}/tasks`
- **Value**: Personal task management alongside Planner
- **Effort**: Low

### 6.5 OneNote Integration
- **Endpoints**: `/me/onenote/notebooks`, `/sections`, `/pages`
- **Value**: Meeting notes, knowledge capture
- **Effort**: Medium

### 6.6 SharePoint Lists CRUD
- **Endpoints**: `/sites/{id}/lists`, `/lists/{id}/items`
- **Value**: Structured data beyond files
- **Effort**: Low

### 6.7 Sharing Links & Permissions
- **Endpoints**: `/drive/items/{id}/createLink`, `/permissions`
- **Value**: Collaboration workflow automation
- **Effort**: Low

### 6.8 Presence & People Insights
- **Endpoints**: `/me/presence`, `/me/people`
- **Value**: Chat/meeting context, quick contact resolution
- **Effort**: Low

### 6.9 Advanced Change Notifications
- **Features**: Resource data in notifications, lifecycle events
- **Value**: Reduce polling, handle renewals securely
- **Effort**: Medium

---

## Future Considerations (Out of Scope for This Plan)

1. **Unified Search Multi-Request**: Run parallel queries for message vs files then merge (recommended by Codex)
2. **Streaming uploads**: Support large file uploads via Graph upload sessions (>4MB)
3. **Winston logging**: Structured logging with automatic redaction (adds dependency)
