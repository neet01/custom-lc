# Master Slack ⟷ Cortex Plan

Last updated: 2026-06-16
Supersedes: `GOVSLACK_EXECUTION_PLAN.md` (2026-05-21), which predates the Slack archive
scaffold now present in the repo and treats Phase 1 as greenfield.

---

## 0. What changed since the previous plan

The previous plan assumed Slack work started from zero. It does not. A Slack archive
**scaffold already exists in the working tree** (currently uncommitted):

| Area | Status | Files |
|---|---|---|
| Slack OAuth v2 (bot + user tokens, enterprise_id) | Built (skeleton) | `api/server/services/SlackArchiveOAuthService.js` |
| Archive models/schemas/methods (conversation, message, syncJob, syncLease) | Built | `packages/data-schemas/src/{models,schema,methods}/slackArchive*` |
| Archive service | Early skeleton (~420 lines vs Teams' ~5,961) | `api/server/services/SlackArchiveService.js` |
| Agent tool + route + status UI + data-provider types | Built (skeleton) | `api/app/clients/tools/util/slackArchive.js`, `api/server/routes/slackArchive.js`, `client/.../SlackArchiveStatus.tsx`, `packages/data-provider/src/types/slackArchive.ts` |
| **Slack → enterprise-memory projection** | **Missing** | (Teams has `EnterpriseMemory/teamsProjection.js`; no Slack equivalent) |
| **Hardening ported from Teams** (deferred/partial run handling, `$text` retrieval) | **Missing** | — |
| **Bot runtime** (Socket Mode / Bolt) | **Not started** | no `@slack/*` deps present |

**Implication:** Phase 1 is ~40–60% scaffolded, not greenfield. The plan below reframes
the work as *complete + harden the existing scaffold and build the missing retrieval
bridge*, not *build from scratch*. It also corrects three structural problems in the
prior plan (see §2).

---

## 1. Goal

A GovSlack bot users interface with, backed by the Cortex reasoning layer (agents,
enterprise memory, Teams/Outlook/Slack archives), running fully in-boundary (GovCloud /
GCC High, Bedrock, self-managed Mongo, in-boundary vector store).

Two distinct capabilities, often conflated, kept separate here:
- **Slack as a source** — archive + index Slack history so Cortex can *retrieve* Slack
  context (read path). Mostly scaffolded.
- **Slack as a surface** — a bot users *talk to* inside GovSlack (interaction path). Not
  started. This is an identity + ingress problem, not an AI problem.

---

## 2. What I'd change about the prior plan

1. **It is out of date.** Reframe Phase 1 from "build the Slack pipeline" to "harden the
   existing scaffold + build the missing Slack→memory projection." Porting the two
   patterns just landed for Teams (`deferred_failed`/`partial` run handling, `$text`
   relevance retrieval) is explicit, scheduled work — not a vague "reuse Teams."

2. **It couples the bot to shared sessions (wrong dependency).** The prior plan makes the
   GovSlack bot (Phase 3) depend on multi-user shared sessions (Phase 2). A **single-user
   bot** — answers as the invoking user, in a DM or thread — needs *none* of the
   owner→session-centric runtime refactor. Shared/collaborative sessions are a large, risky
   refactor that should **not gate** a useful bot going live. This plan **decouples** them:
   ship a single-user bot first; treat collaborative sessions as a later, parallel track.

3. **It under-weights the identity/token crux.** The hardest problem is mapping a Slack
   user → Cortex/Entra identity and obtaining a **delegated Graph (OBO) token** so a Slack
   user can ask about their *Teams/Outlook* data. The Slack OAuth scaffold yields *Slack*
   tokens only. Account-linking to Entra is a first-class workstream here, not a footnote.

4. **It defers the connection-model decision.** For GovCloud, choose **Socket Mode
   (outbound WebSocket, no inbound ingress)** up front — it shapes the network/ATO posture.
   Don't leave it as a Phase-3 implementation detail.

5. **It omits data-leakage controls.** Slack is a shared medium; user-scoped answers must
   be **DM / ephemeral only**, never posted into a channel others can read. Compliance-
   critical; called out explicitly below.

6. **It omits Slack-specific ingestion realities** (threads, tier-based rate limits,
   cursor pagination, edited/deleted subtypes, private-channel/DM consent). These are not
   1:1 with Teams and are where Phase 1 quality is won or lost.

7. **It front-loads no de-risking spike.** The single biggest external unknown — GovSlack
   app approval, Socket Mode availability, scope approvals — is added here as a **Phase 0
   gate** before committing the rest of the sequence.

---

## 3. Architecture (target)

```
GovSlack (in GovCloud)
   │  outbound WebSocket (Socket Mode) — no public inbound endpoint
   ▼
slack_worker            ← NEW dedicated Bolt service (mirrors services/spreadsheet-worker)
   │  • verify + resolve Slack user → Cortex user (SlackIdentityLink)
   │  • ack < 3s, process async, post/update Block Kit message
   │  • enforce DM/ephemeral for user-scoped answers
   ▼
Reused Cortex internals (NOT rebuilt):
   controllers/agents (agent runtime) · tool registry (teams/slack/outlook/memory tools)
   GraphTokenService (delegated/OBO tokens) · enterprise-memory retrieval ($text/vector)
   packages/api/src/usage (credits) · audit trail · MongoDB · in-boundary vector store

Read path (Slack as a source), mostly scaffolded:
   Slack Web API → SlackArchiveService → SlackArchive{Conversation,Message}
       → slackProjection (NEW) → EnterpriseMemory{Entity,Relationship,Chunk}
       → retrieval ($text now, vector later)
```

**Connection model:** Socket Mode (outbound-only) — minimal attack surface, no ingress to
expose/sign, simpler ATO. Keep Events API (signed request URL behind nginx) as the
fallback only if GovSlack mandates a request URL or HA scale-out later requires it.

**Runtime placement:** a dedicated `slack_worker` Node service (the
`services/spreadsheet-worker` pattern), running `@slack/bolt`, importing the shared
`packages/api` agent/tool modules in-process. Keeps the long-lived WebSocket off the
single web/api process and independently restartable.

---

## 4. Identity & authorization (the crux)

A Slack user ≠ a Cortex user, and Cortex's scoped tools (Teams/Outlook/Slack-DM) require a
**delegated token tied to an Entra identity**.

- **Account linking:** first contact from an unlinked Slack user → ephemeral/DM "Connect
  your account" button → existing Entra OAuth (`api/server/routes/oauth.js`) with `state`
  carrying Slack user/team → callback writes a **`SlackIdentityLink`** (slackUserId,
  slackTeamId, enterpriseId, verified slackEmail, cortexUserId, tenantId, linkedAt, status)
  and mints/persists the delegated Graph token (token persistence already exists for Teams).
- **Steady state:** every Slack event resolves Slack→Cortex user and runs the agent **as
  that user**, with that user's `GraphTokenService` token and the same per-user visibility
  scoping the web app enforces.
- **Identity assurance:** map only on **admin-governed, verified** email (Slack profile
  email ↔ Entra UPN). Never trust self-asserted identity on a security platform.
- **Degraded mode:** no link / no token ⇒ bot answers only non-sensitive, org-general
  queries; it must refuse user-scoped retrieval rather than guess.

**Data-scoping rules (compliance):**
- User-scoped answers (Teams/Outlook/Slack-DM/personal memory) → **DM or
  `chat.postEphemeral`** only.
- In channels, answer general/non-scoped queries openly; anything touching a user's
  archive is ephemeral or redirected to DM.
- **Audit every interaction** (who asked, which tools ran, which user-scope, channel
  visibility) into the existing audit trail.

---

## 5. Slack-as-a-source: completing & hardening the archive (read path)

These are the gaps between the current scaffold and Teams-parity.

1. **Build the Slack → enterprise-memory projection** (`EnterpriseMemory/slackProjection.js`),
   mirroring `teamsProjection.js`: per-message chunks + conversation/thread-window chunks,
   person/conversation entities, `has_participant` relationships. **This is the actual
   retrieval bridge and is currently absent** — without it, archived Slack data is not
   searchable through enterprise memory.
2. **Port the Teams rate-limit hardening** into `SlackArchiveService`: durable
   `deferred_failed` conversation state + `partial` run status + per-conversation backoff +
   consecutive-failure budget, behind a flag — adapted to Slack's **tier-based** limits
   (respect `Retry-After`; `conversations.history`/`replies` are Tier 3).
3. **Port `$text` retrieval** for Slack chunks (same flagged, relevance + regex-fallback
   pattern), and make the partial-aware projection rule apply to Slack runs too.
4. **Slack-specific ingestion correctness:**
   - **Threads:** fetch replies via `conversations.replies` per parent `thread_ts`; the
     window-chunker must model thread structure (differs from Teams).
   - **Cursor pagination:** `response_metadata.next_cursor` as the resumable cursor
     (analog of Teams `nextLink`), persisted on the sync job/lease.
   - **Edited/deleted/subtype messages:** handle `message_changed`/`message_deleted` and
     system subtypes; filter bot/app/system noise (analog of Teams system-message logic).
   - **Identity resolution:** Slack user IDs → email → Entra via `users.info`, cached.
   - **Enterprise Grid / GovSlack:** respect `enterprise_id` vs `team_id` (already captured
     by the OAuth scaffold); decide org-wide vs per-workspace token scope.
5. **Consent & scope governance (legal/compliance):** archiving private channels and DMs
   is a data-governance decision, not just a technical scope. Define, with the org, exactly
   which conversation types are archived, under whose token, and with what retention —
   before the pilot.
6. **Capacity:** Slack roughly doubles indexing volume. Apply the vector-store decision
   already reached for Teams — **managed pgvector (RDS/Aurora) separated from file-RAG,
   window-chunk-only, 512-dim + halfvec** — and size from a measured pilot, not estimates.

---

## 6. Phases (reordered, partially parallel)

Dependency-driven, not strictly sequential. Bot delivery is **decoupled** from shared
sessions.

### Phase 0 — GovSlack connectivity spike (de-risk first) · ~1–2 wks
Create the GovSlack app; confirm in *your* GovSlack tenant: Socket Mode availability,
slash commands, Block Kit, account-linking, and the **app-approval process + required
scope approvals**. Achieve a "bot echoes a message" round-trip. **Gate:** do not commit
to Phases 2–4 timelines until this passes.

### Phase 1 — Slack source: complete + harden the archive · ~4–6 wks
Build the Slack projection (§5.1); port deferred/partial hardening (§5.2) and `$text`
(§5.3); fix thread/cursor/edited-message ingestion (§5.4); settle consent/scope (§5.5);
pilot on a bounded channel set and validate archive completeness + retrieval quality.
**Exit:** Slack context is retrievable through enterprise memory at parity with Teams on a
pilot scope. *(Prereq, per prior plan: Teams ingestion/retrieval is stable — now largely
addressed by the deferred/partial + `$text` work.)*

### Phase 2 — Single-user GovSlack bot · ~3–5 wks · depends on Phase 0 + account-linking
Stand up the `slack_worker` Bolt/Socket Mode service; `@Cortex` mention / DM / `/cortex`
handlers; 3-second ack → async → `chat.update`; **account-linking OAuth** + `SlackIdentityLink`;
resolve Slack→Cortex user; run the existing agent **as that user**; enforce DM/ephemeral
scoping; usage attribution + audit. **Exit:** a linked user can ask Cortex from GovSlack
and get user-scoped, in-boundary answers across Teams/Slack/Outlook/memory.
*Does **not** require shared sessions.*

### Phase 3 — Hardening & rollout · ~2–3 wks
Slack tier-based rate-limit/backoff discipline; thread conversation-state mapping
(`SlackConversationState`: thread_ts ↔ Cortex conversation); secrets in AWS Secrets
Manager (GovCloud); monitoring; admin governance of identity links; deploy `slack_worker`
as a compose service; broader pilot.

### Phase 4 — (Optional, parallel/later) Shared collaborative sessions · ~6–10 wks
The owner→session-centric runtime refactor from the prior plan's Phase 2. Enables group/
channel invocations to map to a shared Cortex session. **Deliberately last and decoupled**
— highest engineering effort, and not required for a useful single-user bot.

**Indicative total to a useful bot (Phases 0–3): ~10–16 weeks**, vs the prior plan's
15–25 weeks to the same point (because shared sessions no longer block the bot). Shared
sessions, if pursued, add on top in parallel.

---

## 7. New components to build (net of the scaffold)

- `EnterpriseMemory/slackProjection.js` (+ spec) — the missing retrieval bridge.
- `slack_worker` service (Bolt, Socket Mode) under `services/` — the bot runtime.
- `SlackIdentityLink` model + account-linking flow on the existing Entra OAuth.
- `SlackConversationState` (thread_ts ↔ Cortex conversation) for threaded context.
- Slack rate-limit/deferred-failure hardening + `$text` retrieval in `SlackArchiveService`.

## 8. Reuse (do not rebuild)

Agent runtime (`controllers/agents`), tool registry, `GraphTokenService` (OBO + token
persistence), enterprise-memory retrieval (`$text`/vector), usage/credits
(`packages/api/src/usage`), audit trail, the Slack OAuth + archive models already
scaffolded, and the `services/spreadsheet-worker` service pattern.

---

## 9. Top risks (spike before committing)

1. **GovSlack feature parity & app approval** — Socket Mode, slash commands, scope
   approvals, org app-approval gate. (#1 unknown → Phase 0.)
2. **Delegated Graph token for Slack-initiated requests** — account-linking is mandatory;
   without an Entra login there is no OBO token, so the bot cannot do user-scoped
   Teams/Outlook. Decide: require one-time web/OAuth onboarding (recommended) vs degrade to
   non-scoped.
3. **Per-user data leakage in channels** — enforce DM/ephemeral for scoped data.
4. **Identity assurance** — verified, admin-governed email mapping only.
5. **Consent/governance for private channels & DMs** — settle with the org before pilot.
6. **Vector capacity** — Slack doubles volume; managed pgvector + window-chunk-only +
   512-dim/halfvec, sized from a measured pilot.

---

## 10. Summary

Build order: **(0) spike GovSlack → (1) finish/harden the Slack source → (2) single-user
bot via account-linking → (3) harden/rollout**, with **(4) shared sessions as a later,
parallel, optional track**. The biggest corrections over the prior plan: it's no longer
greenfield (a scaffold exists), the bot is decoupled from the shared-session refactor, and
identity-linking + data-scoping are promoted to first-class workstreams.
