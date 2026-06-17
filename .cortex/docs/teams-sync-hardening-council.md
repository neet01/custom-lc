# Teams Archive Sync Hardening — Council Transcript

> Multi-model Octopus council, captured 2026-06-16. Goal: **advice** · depth: **standard** · domain: **architecture**.
> Members: **Claude Opus 4.6** (strategy-analyst, chair) vs **Codex / GPT-5.3-codex** (backend, database, cloud architects + code-reviewer/verifier).
> Run id: `20260616-195922-00636e`. Quorum met, no veto, two critique rounds, no revisions at standard depth.
>
> This is the durable record of the deliberation that preceded the `deferred_failed` / partial-aware-projection work.
> Source artifacts: `~/.claude-octopus/councils/20260616-195922-00636e/`.

## Topology resolution (answered after the council, from the repo)

The council named "single API process vs. multiple replicas" as a **P0 prerequisite**. Resolved from the codebase:

- Backend starts as plain `node api/server/index.js` (package.json) — **no `cluster`/`fork`/`worker_threads`** in `api/server`. One Node process per container.
- Prod is Docker (self-managed Mongo). `deploy-compose.yml` defines a single `api` service with a fixed `container_name: LibreChat-API`, which *prevents* `docker compose --scale api=N>1`. The upstream Helm chart agrees: `replicaCount: 1`, `autoscaling.enabled: false`.

**Conclusion: single API process today.** Therefore the per-process throttles (300ms page delay, 60s backfill throttle, 10s DB gate, module-level Maps) currently compose as if global, and a cross-process token bucket is not needed. The multi-replica gap the council flagged is **latent** — it activates the moment the `api` service is scaled, at which point only the Mongo lease cap holds.

Bonus: `deploy-compose.yml` confirms `vectordb` (pgvector) + `rag_api` already run in prod — the in-boundary semantic-retrieval path needs no new AWS infra.

---

## Contents
1. [Stage 0 — Shared research context](#stage-0--shared-research-context)
2. [Stage 1 — Opening positions](#stage-1--opening-positions)
3. [Stage 2 — Critiques](#stage-2--critiques)
4. [Stage 3 — Revisions](#stage-3--revisions)
5. [Stage 4 — Synthesis](#stage-4--synthesis)

---

## Stage 0 — Shared research context

# Council Research Context

## Task

Evaluate the Teams archive SYNC hardening in a LibreChat fork ("Cortex"), an enterprise memory platform in AWS GovCloud (GCC High, Microsoft Graph US endpoint, self-managed MongoDB in a Docker container). Main file: api/server/services/TeamsArchiveService.js.

Grounded current state (verified against the working tree, do not re-derive or contradict):
- Uncommitted rate-limit fix already applied: graphRequest() now retries statuses 429/500/502/503/504 with Retry-After header honoring + exponential backoff + jitter; a 300ms delay was added between message pages; the backfill-state snapshot refresh is throttled to 60s; ensureSyncJobActive DB check is time-gated to once per 10s; DEFAULT_MAX_CONCURRENT_SYNCS lowered 3->1; DEFAULT_DISCOVERY_CONCURRENCY lowered 4->2.
- isRecoverableChatMessageError() only treats 403/404 as skippable, NOT 500/502. So when 500/502 retries exhaust, the error propagates and ABORTS the entire sync run (not just the one conversation).
- The message sync loop is fully serial: one conversation at a time, one 50-message page at a time. So the 300ms page delay is request smoothing, not a fix for parallel fanout.
- maxConcurrentSyncs is enforced via MongoDB slot leases (cross-process/replica-safe). It bounds shared Graph pressure, lease contention, Mongo write load, AND inline enterprise-memory projection load (projection is queued at sync completion).
- Deployment: unknown whether single API process or multiple replicas. Module-level throttle Maps are per-process; the lease cap is cross-process.
- A prior single-model (Codex-only) council recommended: treat exhausted 500/502 as a conversation-level "deferred_failed" state that does NOT advance the checkpoint and marks the run "partial", plus a failure budget (abort/alert if exhausted 5xx exceeds N conversations or X%); keep cap=1 for rollout; telemetry-first before raising concurrency; treat per-chat Graph delta as an unvalidated hypothesis (Teams message delta surface may be user-scoped not per-chat); defer a global token bucket. Stress-test and either confirm, refine, or challenge this prior position.

Provide ADVICE (a structured, prioritized recommendation) on:
1. The 500/502 abort risk and the proposed "deferred_failed" + failure-budget design: is it correct, and what are the failure modes (silent archive gaps, retry storms, alert fatigue)?
2. Whether dropping the global sync cap to 1 is appropriate vs an unnecessary throughput cliff; what concrete signal should gate raising it.
3. Whether Microsoft Graph per-CHAT message delta queries are feasible on the GCC High surface and worth adopting; how to validate cheaply.
4. Whether a global token-bucket rate limiter is worth introducing vs current concurrency+page-delay knobs, given the serial loop.
5. The single highest-leverage next step for rollout stability, and the minimal telemetry to add first.
Be conservative and operationally grounded; favor rollout-hardening over architectural elegance. Disagree with the prior council where warranted.

## Local Corpus Evidence

No local corpus workspace was detected for this run.

## Current Source Handling

The shell runner does not fetch external sources directly. Web-capable council members should validate current external sources during fanout when provider tooling allows it.

---

## Stage 1 — Opening positions

### Strategy Analyst — Chair · Claude (Opus 4.6)

## Recommendation: Cortex Teams Archive SYNC Hardening

**Bottom line:** The prior council's direction is right on 4 of 5 points. The one place it needs real correction is the *mechanics* of "deferred_failed" — specifically the checkpoint semantics and the budget trigger. Fix the 500/502 abort first; it's the only failure mode that can wedge the entire archive. Everything else is throughput tuning that can wait behind telemetry.

---

### 1. 500/502 abort + "deferred_failed" — direction correct, mechanics need refinement

**Confirm:** Converting an exhausted-5xx from run-fatal to conversation-level isolation is correct and standard (poison-message / dead-letter handling). At cap=1 with a fully serial loop, the current behavior is the most acute risk in the system: **one persistently-500ing conversation aborts every run and can stall the entire archive indefinitely (head-of-line blocking).** That is worse than a data gap — it's a total stoppage masquerading as a transient error.

**Refine — the most important correction:** "Does NOT advance the checkpoint" is under-specified and, taken literally, recreates the wedge.
- If the checkpoint is a single monotonic run cursor, you cannot both skip a conversation *and* not advance — you either reprocess everything before it forever, or you stall.
- The correct model is **two-level**: (a) **run-level forward progress** must continue past the poison conversation; (b) the conversation's own delta cursor / retry state is recorded in a **separate deferred set** that is *not* advanced and carries its own retry schedule. Decouple "the batch moves on" from "this conversation gets retried later."

**Refine — the failure budget:** A raw "exhausted 5xx > N or > X%" trigger is an alert-fatigue and false-abort trap:
- **Percentage on small batches** (1 of 3 = 33%) fires spuriously. Require an **absolute floor AND a rate** (e.g., N≥5 *and* ≥X%).
- A single-run budget **conflates a transient Graph outage** (correlated burst, self-heals, do not page, do not mark spurious partial) **with a structural poison conversation** (chronic, needs a human). These want different responses. The durable signal is **deferred conversations not clearing after K cycles** — trigger escalation on *persistence across runs*, not raw count in one run.
- The budget-abort itself must use the **clean/resumable shutdown path** (release leases, persist the deferred set), not a thrown exception. Otherwise you trade a data-gap bug for a stuck-lease bug.

**Failure modes to instrument against:**
- *Silent archive gaps* — deferred conversations that never retry, or retry-and-fail forever. In a GovCloud compliance archive, completeness *is* the product. Mitigate with `deferred_count` and `oldest_deferred_age` metrics + a max-attempts cap that escalates to "needs manual intervention."
- *Retry storms* — re-burning the retry budget on the same poison conversation every cycle amplifies Graph pressure. Mitigate with per-conversation exponential backoff on the deferred set.
- *Alert fatigue* — covered above: persistence-based, floored, outage-vs-poison distinction.

---

### 2. Global cap = 1 — correct as a temporary floor, but it needs a paired drain SLO

**Confirm** for the firefight: cap=1 is the right blunt instrument to stop the 429 bleeding during rollout. It bounds all four pressures (Graph, lease contention, Mongo writes, inline projection) at once.

**Challenge (prior council understated this):** cap=1 + serial loop + 300ms page delays is a steep throughput cliff. For a multi-user/multi-team enterprise archive, the real risk is **backlog growing faster than it drains — the archive falls permanently behind**, which is *also* a silent completeness failure, just a slower one. "Keep cap=1" is safe for Graph but potentially *unsafe for archive freshness*. So cap=1 must be paired with a **backlog-drain SLO**, not just 429-watching.

**Signals that gate raising it (1→2 only, one step, then hold):**
1. Graph 429 rate and Retry-After magnitudes at cap=1 → sustained near-zero with headroom.
2. Retry-exhaustion rate → ~0.
3. **Backlog depth + drain rate** → backlog draining, not growing. (If growing at cap=1, you *must* raise it — but only behind a global limiter; see §4.)
4. Mongo write latency / lease contention and projection lag → headroom confirmed.

Soak each step across a window covering peak load before the next.

**Cheap prerequisite:** Resolve the single/multi-replica unknown first — it's one question to ops and it changes the per-process throttle math. (Module-level throttle Maps are per-replica; at cap=1 only one replica holds the lease so only one is actively syncing, so the impact is currently small — but you need to know before raising the cap.)

---

### 3. Per-CHAT Graph delta on GCC High — prior council is right to flag it; validate, don't architect around it

**Confirm the skepticism.** Both surfaces nominally exist in Graph (`/chats/{id}/messages/delta` and the user-scoped, *metered/licensed* `getAllMessages` export API), but GCC High routinely lags commercial in endpoint availability, delta parity, and billing model. I cannot assert from knowledge that per-chat delta returns a working deltaLink, with acceptable throttling, un-double-metered, on the GovCloud surface. That is precisely the uncertainty.

**Two caveats that change its value even if it works:**
- Delta only helps **steady-state incremental** sync. The **initial backfill still pages everything**. If your 429 storm originates in backfill (likely, for a new archive platform), delta won't touch that phase. **Confirm where the 429s actually originate before betting on delta.**
- Fewer requests beats more backoff — if delta works, it's the *real* fix for steady-state throttling, more than any concurrency knob.

**Cheap validation (an afternoon, a few dozen calls — keep off the critical path):** spike against a handful of real GovCloud chat IDs: (a) does `…/messages/delta` return 200 + deltaLink; (b) does the deltaLink actually return only changes on re-query; (c) throttle headers/limits; (d) required scopes; (e) is it billed under the metered export model. Until that confirms, full-paging stays the baseline. High value, medium uncertainty, **not a rollout blocker.**

---

### 4. Global token bucket — defer now, but it's the explicit prerequisite for cap > 1

**Confirm: don't build it now.** At cap=1 single-active there is no parallel fanout to coordinate; concurrency=1 already serializes and the page delay smooths the single stream. A token bucket would add complexity for ~zero current benefit.

**Refine the framing:** It is not "never" — it is **"the thing you build before raising the cap."** Graph throttles on *aggregate per-tenant-app* rate, not per-process. The 300ms page delay is a per-stream property; with N concurrent syncs (or multiple replicas each syncing) the aggregate QPS is uncoordinated and you recreate the 429 storm. So a **cross-process** limiter (Mongo- or Redis-backed leaky/token bucket) is the natural gate for cap>1. Retry-After honoring (already in) is *reactive*; the bucket is *proactive*. Sequence: stabilize at cap=1 → telemetry → to gain throughput, choose *reduce calls (delta)* and/or *global bucket then raise cap* — never raise cap without one of those.

---

### 5. Highest-leverage next step + minimal telemetry

**Single highest-leverage step: fix the 500/502 abort → two-level conversation deferral with a clean, resumable run.** This is #1 on your list and it is correct: at cap=1 serial, the abort is the one failure that can take the whole archive offline on a single flaky conversation. Throughput and delta are optimizations; this is a wedge.

**Minimal telemetry to ship *with* that fix (before any concurrency change) — three questions:**
1. **"Are we still being throttled?"** — Graph outcomes by status (200/429/5xx), retry counts, **retry-exhaustion count**, observed Retry-After distribution.
2. **"Are we silently dropping archive data?"** — `deferred_count`, `oldest_deferred_age`, attempts-per-conversation, run status (complete/partial).
3. **"Is cap=1 keeping up?"** — backlog depth + drain rate.

Add lease acquisition-wait/contention next, but those three are the decision-gating minimum.

---

### Assumptions
- Checkpoint is currently a run-level cursor (the §1 refinement matters most if so; if it's already per-conversation, the deferred set still needs to exist but the wedge risk is lower).
- The 429 storm that prompted this is at least partly backfill-driven (affects how much delta helps).
- Projection-at-completion load scales with sync throughput (so the cap genuinely bounds it).

### Risks of *following* this advice
- Two-level checkpoint + deferred set is more state to get right; a buggy implementation could double-archive or skip. Mitigate with idempotent writes keyed on message ID.
- Holding cap=1 too long risks permanent backlog; the drain SLO is the guardrail — don't let "conservative" become "never catches up."

### Confidence
- **High** — 500/502 abort is the top risk; deferred-set/checkpoint decoupling; budget-on-persistence-not-single-run; token-bucket-as-prerequisite; telemetry set.
- **Medium** — cap-raise sequencing details (depend on the replica topology and where 429s originate).
- **Low/uncertain** — GCC High per-chat delta behavior; *this is exactly why the recommendation is "cheap spike, off critical path," not "adopt" or "reject."*

**Net vs prior council:** Confirm the strategy; correct the checkpoint mechanics and the budget trigger; add a backlog-drain guardrail to cap=1; keep delta and token-bucket explicitly deferred but with named gates.


### Backend Architect — Advisor · Codex (GPT-5.3-codex)

**Recommendation**

I mostly agree with the prior council, with one tightening: do not treat exhausted `500/502` as globally recoverable, and do not merely skip them. Treat them as conversation-level `deferred_failed`, make the whole run `partial`, and keep the conversation checkpoint unchanged. Add a failure budget so systemic Graph trouble still stops the run instead of grinding through every conversation.

**1. 500/502 Abort Risk**

The current behavior is too brittle: one exhausted `500/502` aborts the entire sync, even though the message loop is serial and the failure may be localized to one chat.

Use `deferred_failed` only if it is durable and visible:

- Store `conversationId`, status, last error class/status, attempt count, last failure time, next retry time.
- Do not advance that conversation’s checkpoint.
- Mark the sync run `partial`, not `success`.
- Surface partial counts in admin telemetry and alerts.

Failure modes to guard against:

- **Silent archive gaps:** avoided only if partial runs and deferred conversations are queryable/alerted. A log line is not enough.
- **Retry storms:** likely if every scheduled run immediately retries the same bad conversations. Add per-conversation backoff and `nextRetryAt`.
- **Alert fatigue:** alert on budget breach or stale deferred backlog, not every single failed chat.
- **Systemic outage masking:** if many conversations hit exhausted `5xx` or `429`, stop the run and alert. Suggested initial budget: abort if `deferred_failed >= 5` conversations or `>= 10%` of attempted conversations, whichever comes first.

**2. Global Cap = 1**

Keeping `DEFAULT_MAX_CONCURRENT_SYNCS=1` is appropriate for rollout. It is a throughput cliff, but it is a controlled one. Because the lease is cross-process, it protects Graph, Mongo, lease contention, and the inline memory projection path across replicas. That matters while deployment topology is unknown.

Raise it only after telemetry shows the system is healthy under real load:

- exhausted `429/5xx` near zero over 24-72 hours
- retry rate low, for example `< 1-2%` of Graph requests
- p95/p99 Graph page latency stable
- Mongo write latency stable
- projection queue does not grow after sync completion
- sync backlog/RPO is unacceptable at cap `1`

Then raise by feature flag to `2`, not back to `3`, and observe again.

**3. Per-Chat Delta**

Do not assume per-chat message delta exists. Microsoft’s current Graph docs show chat message delta at the user scope: `GET /users/{id}/chats/getAllMessages/delta`, available in US Government L4/L5, with `$top` max `50` and an eight-month delta window. The documented per-chat API is `GET /chats/{chat-id}/messages`; it supports filtering/order by timestamps, but the v1.0 doc does not show `/chats/{chat-id}/messages/delta`.

Cheap validation in GCC High:

- Call `https://graph.microsoft.us/v1.0/users/{id}/chats/getAllMessages/delta?$top=1`.
- Call `https://graph.microsoft.us/v1.0/chats/{chatId}/messages/delta?$top=1`.
- Record status, error code, headers, and whether `@odata.deltaLink` appears.
- If per-chat delta returns unsupported/404, stop pursuing it.
- Consider timestamp-filtered per-chat polling as a safer incremental improvement before redesigning around user-scoped delta.

Sources: Microsoft Graph delta docs for `chats/getAllMessages/delta`, per-chat list messages docs, and national cloud endpoint docs.  
https://learn.microsoft.com/en-us/graph/api/chatmessage-delta  
https://learn.microsoft.com/en-us/graph/api/chat-list-messages  
https://learn.microsoft.com/en-us/graph/deployments

**4. Global Token Bucket**

Defer it. With one global sync lease and a serial message loop, the current knobs already bound request pressure: retry/backoff, page delay, discovery concurrency, and cross-process sync leases.

A token bucket becomes worth it when either:

- cap is raised above `1`,
- discovery and message sync overlap enough to create burst pressure,
- multiple Graph-heavy services share the same app registration,
- telemetry shows `429` despite cap `1`.

Until then, a distributed token bucket adds state, failure modes, and tuning burden without clear payoff.

**5. Highest-Leverage Next Step**

Implement durable partial-sync semantics before increasing throughput.

Minimal telemetry first:

- Graph request count by endpoint family and status
- retry count and retry delay total
- exhausted retry count by status
- conversations attempted/succeeded/deferred/skipped
- sync run final state: `success`, `partial`, `failed`, `aborted_budget`
- oldest deferred conversation age
- sync duration and pages/messages processed
- projection queue enqueue count and lag

**Confidence:** High on the rollout recommendation and cap strategy. Medium on delta feasibility because Microsoft documents user-scoped chat message delta for Gov clouds, but GCC High tenant behavior should still be verified with live calls against `graph.microsoft.us`.


### Database Architect — Advisor · Codex (GPT-5.3-codex)

**Recommendation**

The prior council is mostly right. The current 500/502 behavior is the biggest rollout risk: after retries exhaust, one bad Teams conversation can abort the entire archive run. That is operationally fragile and will create large archive gaps from small transient failures.

1. **Implement conversation-level `deferred_failed` for exhausted 500/502, with a run-level failure budget.**
   - Correct design: do not advance that conversation’s checkpoint; mark the overall run `partial`; continue with other conversations.
   - Add `nextAttemptAt`, `attemptCount`, `lastErrorStatus`, and `lastErrorRequestId`/Graph request id if available, without logging message content.
   - Abort the run only when the budget is exceeded, for example `>N conversations` or `>X%` deferred, because broad 5xx usually means service pressure or outage.
   - Treat exhausted `429` differently: that is usually global pressure, not a bad conversation. Prefer run cooldown/abort rather than marking many conversations deferred.

2. **Keep global sync cap at `1` for rollout, but make it explicitly temporary.**
   - This is conservative but justified because the Mongo lease cap protects Graph, Mongo writes, and projection load across replicas.
   - Raise to `2` only after telemetry shows: near-zero exhausted 5xx, low `429` rate, stable `Retry-After`, no growing deferred backlog, acceptable Mongo write latency, and projection queue drain time staying flat.
   - Raise by feature flag/canary, not as a default flip.

3. **Do not assume per-chat delta exists. Validate, but do not redesign around it yet.**
   - Microsoft documents chat message delta as user-scoped: `GET /users/{id}/chats/getAllMessages/delta`, not `/chats/{chat-id}/messages/delta`.
   - The delta API is documented as available in US Gov L4/GCC High, and GCC High should use `https://graph.microsoft.us`.
   - Cheap validation: with one GCC High test user and app token, call `/v1.0/users/{id}/chats/getAllMessages/delta?$top=1`, verify `200`, `@odata.nextLink`/`@odata.deltaLink`, returned `chatId`, permissions, and compare a few known messages.
   - Caveat: delta only returns messages within the last eight months, so it cannot replace historical backfill.

4. **Defer a global token bucket.**
   - With cap `1` and a serial message loop, a distributed bucket is probably complexity without much benefit.
   - Current controls already address the immediate pressure: cross-process lease cap, page delay, Retry-After honoring, and exponential backoff.
   - If throttling persists after raising concurrency, first add a Mongo-backed tenant/job cooldown from observed `Retry-After`; only then consider a true distributed token bucket.

**Highest-Leverage Next Step**

Build the durable `deferred_failed` path plus minimal telemetry before changing concurrency again. That gives you failure isolation, visible archive gaps, and the evidence needed to raise throughput safely.

Minimal telemetry:
- run status: `success`, `partial`, `failed`
- conversations: processed, skipped 403/404, deferred 5xx, aborted budget
- Graph requests by endpoint/status, retry count, final exhausted status, Retry-After delay
- pages fetched, messages upserted, checkpoint advanced count
- deferred backlog count and oldest deferred age
- Mongo write latency/errors and projection queue enqueue/drain timing
- active sync leases and lease wait/contend count

**Key Risks**

Silent gaps if `partial` is treated like success. Retry storms if deferred conversations retry every run without per-conversation backoff. Alert fatigue if every conversation emits an alert instead of one aggregate run alert with sampled examples.

Confidence: high on 500/502 handling and cap guidance; medium on delta adoption because GCC High behavior still needs tenant-level validation despite Microsoft’s docs. Sources: Microsoft Graph chat delta docs and national cloud endpoint docs.


### Cloud Architect — Advisor · Codex (GPT-5.3-codex)

**Recommendation**

I mostly agree with the prior council, with one refinement: make `deferred_failed` impossible to confuse with success.

1. **500/502 Abort Risk**
   The current behavior is too brittle for rollout. One exhausted transient Graph `500/502` aborting the whole serial sync can cause repeated large reruns, more Graph pressure, and delayed archives for unrelated conversations.

   Use conversation-level `deferred_failed` for exhausted `500/502/503/504`, but only if:
   - The conversation checkpoint/watermark is **not advanced**.
   - The overall job is marked `partial`, never `success`.
   - The failed conversation stores `chatId`, status, Graph request id if available, attempt count, first/last failure time, and `nextRetryAt`.
   - Projection/memory indexing does not treat the archive as complete without seeing the partial state.

   Add a failure budget to distinguish isolated Graph instability from systemic outage. Example: abort the run if transient failures exceed `5 conversations` or `10-20%` after a minimum sample size. Page once per tenant/run, not once per chat.

   Main failure modes:
   - **Silent archive gaps** if partial status is hidden or downstream projection treats it as complete.
   - **Retry storms** if failed chats are retried immediately every run without exponential cooldown.
   - **Alert fatigue** if every deferred chat creates an alert instead of a rolled-up run-level signal.

2. **Global Sync Cap = 1**
   Keep `maxConcurrentSyncs=1` for initial GCC High rollout. It is a throughput cliff, but it is also the only cross-process control currently bounding Graph calls, Mongo writes, lease contention, and post-sync projection load.

   Raise it only behind a feature flag after telemetry shows:
   - Low Graph retry rate, especially `429` and exhausted `5xx`.
   - Low accumulated retry/backoff time per run.
   - No Mongo write latency or lease contention regression.
   - Projection backlog drains within the expected window.
   - Sync lag remains operationally acceptable.

   Concrete gate: run a representative tenant at cap `1`; if `429` rate is near zero, exhausted transient failures are rare, Mongo/projection latency is stable, and backlog still misses SLA, test cap `2` for that tenant only.

3. **Per-Chat Delta**
   Do not assume per-chat message delta exists. Microsoft documents chat message delta as user-scoped: `GET /users/{id}/chats/getAllMessages/delta`, available in US Gov L4/L5, with `$top` max `50` and an eight-month delta window. The ordinary per-chat messages API supports `GET /chats/{chat-id}/messages`, but the documented v1.0 page does not expose `/chats/{chat-id}/messages/delta`.

   Cheap validation:
   - In a GCC High test tenant, call `https://graph.microsoft.us/v1.0/chats/{chat-id}/messages/delta?$top=1`.
   - Record exact status/body/headers.
   - Also test documented `users/{id}/chats/getAllMessages/delta`.
   - Adopt only documented/supported behavior unless Microsoft confirms the per-chat surface.

4. **Global Token Bucket**
   Defer it. With cap `1`, discovery concurrency `2`, serial message sync, retry handling, and 300ms page delay, a distributed token bucket adds operational complexity before proving the current knobs are insufficient.

   Reconsider a Mongo-backed or Redis-backed tenant/app token bucket only if telemetry shows throttling across replicas despite the lease cap, or if raising cap above `1` causes correlated `429/Retry-After` spikes.

5. **Highest-Leverage Next Step**
   Implement `partial` sync semantics with persisted conversation-level transient failures and minimal telemetry before raising concurrency.

   Minimal telemetry first:
   - Per run: conversations attempted/succeeded/deferred/failed, pages fetched, messages written, duration.
   - Graph: request count by endpoint/status, retry count, retry sleep seconds, exhausted retry count, `Retry-After` distribution.
   - Control-plane: lease wait/acquire/release, cap in effect, process id.
   - Downstream: projection queue depth/time-to-drain.
   - Archive health: oldest deferred conversation age and count.

**Confidence:** High on keeping cap `1`, adding partial/deferred semantics, and deferring token bucket. Medium on Graph delta specifics because GCC High behavior should still be validated directly against `graph.microsoft.us`.

Sources: Microsoft Graph `chatMessage: delta`, `list messages in a chat`, national cloud endpoints, and throttling guidance.


### Code Reviewer — Verifier · Codex (GPT-5.3-codex-spark)

**Recommendation**

The prior position is mostly correct. Keep the rollout conservative, but make the 500/502 behavior explicit and observable before changing throughput.

1. **500/502 abort risk: fix first**
   Treat exhausted `500/502/503/504/429` as conversation-level `deferred_failed`, not run-fatal, unless a failure budget is exceeded. The key invariant: do **not** advance that conversation’s checkpoint or mark it synced. The run should complete as `partial`.

   Main failure modes to guard:
   - **Silent gaps:** avoided only if partial runs, failed conversations, and stale checkpoints are visible in status/telemetry.
   - **Retry storms:** add per-conversation `nextAttemptAt` with backoff after exhausted retries.
   - **Alert fatigue:** alert on aggregate conditions, not each failed chat. Example: failure budget exceeded, same chat failing for N runs, or archive freshness breaching RPO.

2. **Global cap = 1 is appropriate for rollout**
   It is a throughput cliff, but a defensible one because the lease cap is cross-process and protects Graph, Mongo, slot contention, and projection load. Do not treat it as the target steady state.

   Raise it only when telemetry shows:
   - low or zero sustained Graph `429`;
   - low exhausted retry rate, especially `5xx`;
   - bounded total `Retry-After` sleep time per run;
   - Mongo write latency and projection queue lag stable;
   - sync duration/backlog is missing the product RPO at cap 1.

   Then raise behind a feature flag from `1 -> 2`, not directly back to `3+`.

3. **Per-chat delta: do not bet rollout on it**
   Microsoft’s documented Teams chat delta endpoint is user-scoped: `GET /users/{id}/chats/getAllMessages/delta`, returns messages from all chats for that user, and is available in US Government L4/L5 clouds. It is not the same as a per-chat `/chats/{chat-id}/messages/delta` surface. The normal per-chat list endpoint documents `/chats/{chat-id}/messages` with paging/filtering, not per-chat delta. Sources: Microsoft Graph chat delta docs and national cloud endpoint docs:  
   https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0  
   https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0  
   https://learn.microsoft.com/en-us/graph/deployments

   Cheap validation: in a non-production GCC High tenant, try exactly the hypothesized per-chat delta URL against `https://graph.microsoft.us`. If it returns unsupported/404, close the idea. If it works, still verify token semantics, permissions, message coverage, deletes/edits, retention window, and whether it changes throttling behavior.

4. **Global token bucket: defer**
   With a serial message loop, cap leases, retries honoring `Retry-After`, and a page delay, a distributed token bucket is likely complexity before evidence. Reconsider only if telemetry shows throttling across multiple job types/processes that the lease cap cannot control, or if raising cap creates bursty shared Graph pressure.

5. **Highest-leverage next step**
   Implement `deferred_failed + partial run + failure budget + telemetry` before increasing throughput.

**Minimal Telemetry**

Add counters/timers for:
- Graph requests by endpoint class/status/retry count;
- total `Retry-After` sleep and exponential-backoff sleep;
- exhausted retries by conversation and status;
- conversations completed/skipped/deferred_failed per run;
- checkpoint age / archive freshness per conversation;
- run status: `success`, `partial`, `failed`;
- sync duration, lease wait time, Mongo write latency, projection queue lag.

**Confidence**

High on the 500/502 handling and cap=1 rollout posture. Medium on delta adoption details: docs support user-scoped chat delta in GCC High, but not per-chat delta, so validate live before designing around it.


---

## Stage 2 — Critiques

### Strategy Analyst — Chair · Claude (Opus 4.6)

## Strategy-Analyst Cross-Critique: Teams Archive SYNC Hardening

The four peer responses are nearly unanimous, and on the load-bearing calls (deferred_failed semantics, cap=1, defer token bucket, validate-don't-assume delta) they are correct. My job here is not to re-ratify that consensus but to stress-test where it's too comfortable and reorder priorities for a compliance-sensitive GovCloud memory platform. I confirm the prior council's spine, **refine three points materially, and challenge two.**

---

### Recommendation (prioritized)

**P0 — Resolve deployment topology *first*. It's free and it gates everything else.**
Every peer wrote "unknown whether single or multiple replicas" and then reasoned past it. That's a mistake. Single- vs multi-replica changes the meaning of nearly every other knob: the 300ms page delay, the 60s backfill throttle, and the 10s DB gate are **per-process** and do **not** compose across replicas; only the Mongo lease cap is cross-process. You cannot correctly size the cap, judge whether per-process smoothing is real protection, or decide on a token bucket without this fact. It is a one-line ops lookup. Do it before anything else.

**P0 — The real correctness invariant is checkpoint-non-advancement *plus a partial-aware projection*.** The peers all named "don't advance the checkpoint" but largely glossed the downstream consumer. The projection is queued at sync completion; for a **memory platform**, the catastrophic failure is not a delayed archive — it's an incomplete archive that the projection indexes and presents as authoritative. In GovCloud/compliance terms, "memory that silently lost the 5xx-failed conversations" is worse than a failed run, because nobody knows to distrust it. So the invariant is two-sided:
- Deferred conversation → checkpoint not advanced, run marked `partial`.
- Projection must be **partial-aware**: either skip indexing deferred conversations and mark their memory slice stale, or refuse to treat a `partial` run as a complete snapshot. This is the single highest-stakes item and it is *not* primarily a telemetry problem — telemetry tells you gaps exist; this invariant prevents gaps from being mistaken for completeness.

**P1 — Adopt deferred_failed, but change the failure-budget shape.** Here I disagree with the prior council and all four peers, who parroted "abort if ≥5 conversations or ≥10%." A fixed cumulative count is dimensionally wrong: for a 10-conversation tenant, 5 = 50%; for a 10,000-conversation tenant, 5 aborts on the normal transient-failure floor. More importantly, **a cumulative budget conflates two different signatures**:
- *Scattered* exhausted-5xx across unrelated chats over time → genuinely bad individual conversations → defer and continue is correct.
- *Clustered/consecutive* exhausted-5xx → service outage or throttling wall → continuing just generates a retry storm against a sick endpoint.

Better budget: **abort on N consecutive exhausted 5xx (outage signature), with a percentage ceiling that only activates above a minimum-sample floor.** Consecutive-failure detection is a cheaper, sharper outage detector than a cumulative tally and avoids both false aborts on large tenants and grind-through on small ones.

**P1 — Keep cap=1 for rate safety, but instrument backlog-drain-rate from day one as a co-equal gate.** The peers reflexively endorsed cap=1 and listed only *rate-health* signals (429 near zero, stable Retry-After) as the gate to raise it. That's half the picture. If the deployment is multi-replica, cap=1 serializes the **entire fleet** to one sync globally — that may be unable to drain the backlog within RPO *even when perfectly healthy*. In that case cap=1 is not "conservative," it's a guaranteed SLO miss wearing a safety costume, and waiting on rate telemetry just delays discovering you need per-tenant caps or parallelism. So the gate to raise cap has **two independent axes**:
- Rate axis (raise is *safe*): exhausted 5xx/429 ≈ 0, bounded Retry-After sleep, stable Mongo write + projection drain.
- Throughput axis (raise is *needed*): measured backlog-drain-rate vs. ingest-rate at cap=1 against the product RPO.

Raise (flag, 1→2, canary one tenant) only when rate says *safe* and throughput says *needed*. If throughput says needed but rate says unsafe, you have a real architecture problem (per-tenant fairness / cross-process rate control), not a knob to turn.

**P2 — Token bucket: defer, with a *precise* trigger.** Agree with consensus, but sharpen it. The trigger is not vague "429 despite cap=1." It is specifically: **raising cap above 1 in a confirmed multi-replica deployment.** At that moment the per-process smoothing knobs stop composing and the lease cap becomes the *only* cross-process control — too coarse to shape burst pressure. That is exactly when a Mongo/Redis-backed cross-process limiter earns its complexity. Until then it's premature.

**P2 — Per-chat delta: validate cheaply, but deprioritize it harder than the peers did.** Three peers correctly cite that Microsoft documents *user-scoped* `GET /users/{id}/chats/getAllMessages/delta` (available in US Gov L4/L5), not `/chats/{chat-id}/messages/delta`. I agree — and I'll add the strategic point they soft-pedaled: **delta is a throughput/cost optimization, not a stability fix, and the 8-month delta window means it cannot replace historical backfill regardless.** The task explicitly favors rollout-hardening over elegance. So even a *successful* validation doesn't move the rollout needle. Run the cheap probe (one app-token call each to the per-chat and user-scoped URLs against `https://graph.microsoft.us`, record status/headers/deltaLink), file the result, and move on. Do not let it consume rollout bandwidth.

---

### Minimal telemetry to add first (tightened to the decisions above)
Only what gates a P0/P1 decision:
1. **Run status enum** (`success` | `partial` | `failed` | `aborted_budget`) — without this, `partial` masquerades as `success`.
2. **Deferred-conversation registry**: `chatId`, status, attemptCount, firstFailedAt, `nextRetryAt`, and **oldest-deferred age** (this is your archive-freshness/RPO signal — alert on *this*, not per-chat).
3. **Consecutive exhausted-5xx counter** (drives the refined budget) + exhausted-retry count by status.
4. **Backlog-drain-rate vs ingest-rate** (the throughput-axis cap gate the peers omitted).
5. Graph requests by endpoint/status, total Retry-After + backoff sleep seconds (rate-axis cap gate).
6. Projection enqueue/drain + a flag for whether projection saw a `partial` run (validates the P0 invariant).

Alerting: SLO alert on **freshness/RPO breach** (user-facing contract) and on **budget abort**; everything else is a diagnostic dashboard, not a page. This avoids the alert-fatigue failure mode the peers flagged.

---

### Assumptions
- "Enterprise-memory projection" is a real downstream consumer that can act on archive state; if it already tolerates incremental/partial data, the P0 projection-awareness item shrinks to a verification rather than a build.
- RPO/freshness is an actual product commitment; if there's no freshness SLO, define one before raising the cap, because "is cap=1 enough" is otherwise unanswerable.
- The Mongo lease is genuinely cross-process-correct (atomic acquire/expiry). If lease expiry races exist, that's a separate latent bug that cap=1 partially masks.

### Risks
- **Silent archive gaps (highest):** mitigated only by the partial-aware projection invariant, not by telemetry alone.
- **Retry storms:** mitigated by `nextRetryAt` per-conversation backoff *and* the consecutive-failure abort (so a tenant-wide outage doesn't re-storm every scheduled run).
- **False safety from cap=1 in multi-replica:** mitigated by resolving topology (P0) and the throughput-axis instrumentation.
- **Alert fatigue:** mitigated by paging on freshness/budget aggregates only.
- **Delta rabbit hole:** mitigated by treating it as post-rollout, time-boxed validation.

### Implementation notes
- Treat exhausted **429** differently from 5xx (the database-architect peer got this right and it's underweighted elsewhere): 429 is global pressure → prefer run cooldown/abort over deferring many conversations, since deferring-and-continuing just re-hits the throttle wall.
- Persist the deferred registry in Mongo keyed by `chatId`, not in the per-process throttle Maps — those evaporate on restart and won't survive replica churn.
- Keep `isRecoverableChatMessageError()` returning 403/404 as *skip* (genuinely inaccessible), and route exhausted 500/502/503/504 to the new *defer* path — distinct states, distinct telemetry, distinct retry policy.

### Confidence
**High** on: topology-first, the projection partial-awareness invariant as the true P0, deferring the token bucket, and deprioritizing delta. **High** on the consecutive-failure-vs-cumulative budget refinement being a genuine improvement over the prior council. **Medium** on the cap-throughput-gate framing — it's directionally right but depends on backlog math I can't compute without tenant-size and per-message-latency figures. **Medium** on delta feasibility specifics (docs support user-scoped, not per-chat; GCC High needs live confirmation), consistent with the peers.


### Backend Architect — Advisor · Codex (GPT-5.3-codex)

**Recommendation**

1. **Fix the 500/502 abort first.** The prior council is directionally correct: exhausted transient Graph failures should not abort the whole sync. Implement conversation-level `deferred_failed` for exhausted `500/502/503/504`, keep that conversation’s checkpoint unchanged, continue the run, and mark the run `partial`, never `success`.

2. **Refine the checkpoint model.** “Do not advance checkpoint” must mean the failed conversation checkpoint, not the whole run cursor. If there is only a run-level cursor today, split it. Otherwise one bad conversation still causes head-of-line blocking.

3. **Use a failure budget, but avoid noisy single-run thresholds.** Prefer `absolute floor + percentage + persistence`, e.g. escalate when at least `N` conversations and `X%` fail, or when the same deferred conversations survive `K` runs. Treat exhausted `429` as tenant/app pressure, not a bad chat: clean cooldown/abort is better than marking many chats deferred.

4. **Keep global sync cap at `1` for rollout, but make it explicitly temporary.** It is a throughput cliff, but it is currently the only cross-process control bounding Graph pressure, Mongo writes, lease contention, and projection load. Raise `1 -> 2` only after telemetry shows low `429`, low exhausted retries, stable Retry-After durations, stable Mongo/projection latency, and backlog not clearing fast enough at cap `1`.

5. **Do not bet on per-chat delta.** Microsoft documents chat message delta as user-scoped: `GET /users/{id}/chats/getAllMessages/delta`, available for US Government L4/L5, with `$top` max `50` and an eight-month window. The documented per-chat list API is `GET /chats/{chat-id}/messages`; it does not document `/chats/{chat-id}/messages/delta`. Validate cheaply against `https://graph.microsoft.us`, but do not redesign rollout around it yet. Sources: [chatMessage delta](https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0), [list chat messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0), [national cloud endpoints](https://learn.microsoft.com/en-us/graph/deployments).

6. **Defer a global token bucket.** With cap `1` and a fully serial message loop, concurrency leases plus Retry-After handling and the 300ms page delay are enough for now. Revisit a cross-process Mongo/Redis limiter only before raising concurrency materially or if telemetry shows correlated throttling across job types.

**Assumptions**

- Message writes are idempotent by Graph message ID.
- Projection can consume a `partial` archive state without treating it as complete.
- Deployment topology may include multiple API replicas, so process-local throttles cannot be trusted as global controls.

**Risks**

- Silent archive gaps if `partial` is hidden or downstream memory projection treats it as success.
- Retry storms if deferred chats retry every run without `nextAttemptAt` and exponential backoff.
- Alert fatigue if each chat failure pages instead of one aggregate run-level signal.
- Freshness/RPO breach if cap `1` remains in place after backlog starts growing.

**Implementation Notes**

Persist per deferred conversation: `chatId`, status, attempt count, first/last failure time, `nextAttemptAt`, last Graph request id/correlation id, and sanitized error class. Do not store message content or sensitive payloads in logs.

Minimal telemetry first:

- Graph requests by endpoint/status, retry count, exhausted retry count.
- Retry-After/backoff sleep seconds.
- Conversations attempted/succeeded/skipped/deferred per run.
- Run status: `success`, `partial`, `failed`.
- Deferred backlog count and oldest deferred age.
- Backlog depth, drain rate, sync duration.
- Mongo write latency, lease wait/contention, projection queue lag.

**Highest-Leverage Next Step**

Ship `deferred_failed + partial run semantics + minimal telemetry` before any further concurrency change. That removes the current whole-run wedge and gives the signals needed to safely tune throughput.

**Confidence**

High on the 500/502 handling, cap `1` rollout posture, and deferring token bucket. Medium on delta adoption details: the documented GovCloud surface is user-scoped delta, but live GCC High tenant validation is still required.


### Database Architect — Advisor · Codex (GPT-5.3-codex)

## Recommendation

As database architect, I agree with the prior council’s direction, with one important refinement: `deferred_failed` must be modeled as durable per-conversation state, not as an in-memory skip or a vague run outcome. The database state machine is the control plane here.

Priority order:

1. Fix exhausted `500/502/503/504/429` handling before changing throughput.
2. Keep global sync cap at `1` for rollout, but measure backlog drain so safety does not become permanent staleness.
3. Validate Graph delta cheaply, but do not design rollout around undocumented per-chat delta.
4. Defer a global token bucket until telemetry proves current controls are insufficient or cap is raised.
5. Add minimal archive-health telemetry with the deferral change.

## 1. 500/502 Abort Risk

Current behavior is too brittle. A single exhausted transient Graph failure aborting the whole serial sync creates head-of-line blocking and can repeatedly prevent unrelated conversations from archiving.

The proposed `deferred_failed` design is correct if implemented this way:

- The failed conversation’s checkpoint is not advanced.
- The run-level cursor or scheduler can continue past that conversation.
- The sync run is marked `partial`, never `success`.
- The deferred conversation is persisted with:
  - `conversationId`
  - last status/error class
  - attempt count
  - first/last failure time
  - `nextRetryAt`
  - last Graph request id if available
- Retries use per-conversation backoff.
- Downstream memory projection must see the run as partial and must not imply archive completeness.

The key database concern: do not couple global run progress to a single poisoned conversation. You want two levels of state: run progress and per-conversation sync/checkpoint state.

Failure budget should avoid both false positives and silent degradation. Prefer:

- Abort or mark `aborted_budget` only after an absolute floor and percentage threshold, for example `>= 5 conversations` and `>= 10%` after enough samples.
- Alert on persistent deferred backlog, not every individual transient failure.
- Escalate if `oldest_deferred_age` exceeds RPO or the same conversation fails across multiple cycles.

Main failure modes:

- Silent archive gaps if partial state is not surfaced.
- Retry storms if failed chats retry every scheduled run.
- Alert fatigue if every failed conversation pages someone.
- Stuck leases or inconsistent state if budget aborts throw instead of taking the normal cleanup path.

## 2. Global Cap = 1

Keeping `DEFAULT_MAX_CONCURRENT_SYNCS=1` is appropriate for rollout. It is conservative, but it is the only currently verified cross-process control that bounds Graph pressure, Mongo writes, lease contention, and projection load.

It is also a throughput cliff. Do not leave it at `1` indefinitely without a freshness signal.

Raise from `1` to `2` only when these are true under representative load:

- Sustained low `429` rate.
- Rare exhausted `5xx/429` retries.
- Total retry/backoff sleep is low and stable.
- Mongo write latency is stable.
- Lease contention is low.
- Projection queue drains within target.
- Archive backlog or RPO is unacceptable at cap `1`.

If cap increases, do it by feature flag and tenant cohort, not globally.

## 3. Graph Per-Chat Delta

Do not assume per-chat message delta is available or supported in GCC High. The better-known documented surface is user-scoped chat message delta, not necessarily `/chats/{chatId}/messages/delta`.

Cheap validation:

- In a GCC High test tenant, call:
  - `https://graph.microsoft.us/v1.0/chats/{chatId}/messages/delta?$top=1`
  - `https://graph.microsoft.us/v1.0/users/{userId}/chats/getAllMessages/delta?$top=1`
- Record status, error code, headers, `@odata.deltaLink`, permissions required, and throttling behavior.
- Confirm edits/deletes, retention window, and message coverage.
- If per-chat delta is unsupported, close that path.

Delta is worth pursuing only for steady-state incremental sync. It will not solve initial backfill pressure.

## 4. Global Token Bucket

Defer it for now. With one cross-process sync lease and a serial message loop, a distributed rate limiter adds operational state without much immediate benefit.

Reconsider it when:

- cap is raised above `1`;
- multiple Graph-heavy workers share the same app registration;
- discovery and sync overlap enough to create aggregate bursts;
- telemetry shows `429` despite cap `1`.

If introduced later, it should be cross-process, likely Mongo-backed unless another self-hosted rate-control service already exists. Do not use external SaaS.

## Implementation Notes

Persist deferred state in Mongo as first-class sync state, not logs. Make updates idempotent using stable conversation/message identifiers.

Suggested run states:

- `success`
- `partial`
- `failed`
- `aborted_budget`

Suggested conversation states:

- `synced`
- `skipped_permission`
- `deferred_failed`
- `retry_scheduled`

Minimal telemetry to add first:

- Graph requests by endpoint family and status.
- Retry count, exhausted retry count, and retry sleep seconds.
- Conversations attempted, succeeded, skipped, deferred.
- `oldest_deferred_age`.
- Sync backlog depth and drain rate.
- Run duration and final state.
- Lease wait/acquisition time.
- Projection queue depth or lag.

## Assumptions

- Message sync is serial as stated.
- Mongo slot leases are the authoritative cross-process sync limiter.
- Archive completeness matters more than short-term throughput.
- Deployment topology may include multiple API replicas, so module-level Maps are not reliable global controls.

## Confidence

High on fixing `500/502` as conversation-level durable deferral first, keeping cap `1` for rollout, and deferring token bucket.

Medium on delta feasibility because GCC High behavior must be validated live against `graph.microsoft.us`.

The highest-leverage next step is: implement durable `deferred_failed` plus `partial` run semantics and the minimal telemetry above before raising concurrency.


### Cloud Architect — Advisor · Codex (GPT-5.3-codex)

**Recommendation**

1. **Fix the exhausted `500/502` abort first.** The prior council is directionally correct: one bad or transiently failing conversation should not abort the entire serial sync. Treat exhausted `500/502/503/504` as conversation-level `deferred_failed`, mark the run `partial`, and continue.

   Key refinement: separate run progress from conversation progress. The run can move past the conversation, but that conversation’s checkpoint/delta cursor must not advance. Persist `conversationId`, last status/error class, Graph request id if available, attempt count, `lastFailedAt`, and `nextRetryAt`.

   Do **not** silently skip it. The main risks are:
   - **Silent archive gaps:** if `partial` is treated like success or deferred conversations are not visible.
   - **Retry storms:** if every run retries the same failing chat immediately. Use per-conversation backoff.
   - **Alert fatigue:** alert on aggregate/stale conditions, not every failed chat.
   - **Outage masking:** if many conversations fail with exhausted 5xx, stop cleanly and alert.

   I would not classify exhausted `429` the same way. `429` usually means tenant/app pressure, so prefer a run-level cooldown or clean abort rather than creating many per-chat deferred records.

2. **Keep global sync cap at `1` for rollout, but treat it as a temporary guardrail.** Because the lease is Mongo-backed and cross-process, cap `1` is the only currently reliable control across unknown replica topology. It protects Graph, Mongo writes, slot contention, and projection load.

   It is also a throughput cliff. Raise only when telemetry shows:
   - sustained low `429` and low Retry-After time;
   - near-zero exhausted 5xx;
   - deferred backlog is clearing, not aging;
   - Mongo write latency and lease contention are stable;
   - projection queue lag is bounded;
   - archive backlog/RPO cannot be met at cap `1`.

   Raise by feature flag from `1 -> 2`, then soak through peak load. Do not jump back to `3`.

3. **Do not bet on per-chat delta yet.** Microsoft documents chat message delta as user-scoped: `GET /users/{id}/chats/getAllMessages/delta`, available in US Gov L4/L5, with an eight-month delta window and `$top` max `50`. The standard per-chat list API is `GET /chats/{chat-id}/messages`; I would treat `/chats/{chat-id}/messages/delta` as unvalidated until tested in the GCC High tenant. Sources: [chatMessage delta](https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0), [list chat messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0), [national cloud endpoints](https://learn.microsoft.com/en-us/graph/deployments).

   Cheap validation: against `https://graph.microsoft.us`, call both:
   - `/v1.0/users/{id}/chats/getAllMessages/delta?$top=1`
   - `/v1.0/chats/{chatId}/messages/delta?$top=1`

   Record status, error code, permissions, presence of `@odata.nextLink`/`@odata.deltaLink`, returned `chatId`, and throttling headers. Even if user-scoped delta works, it helps steady-state incremental sync more than initial historical backfill.

4. **Defer a global token bucket for now.** With one active sync lease and a serial message loop, the current controls are enough for rollout: lease cap, page delay, Retry-After honoring, exponential backoff, and reduced discovery concurrency.

   Revisit a distributed limiter before raising concurrency beyond `1` if telemetry shows aggregate Graph pressure, multiple Graph-heavy jobs sharing the same app registration, or 429s persisting despite cap `1`.

**Highest-Leverage Next Step**

Implement durable `deferred_failed` partial-run semantics with telemetry. This removes the current head-of-line failure where one exhausted 5xx can wedge the whole archive.

**Minimal Telemetry First**

Track:
- Graph requests by endpoint family, status, retry count, exhausted status, Retry-After delay;
- conversations attempted, completed, skipped `403/404`, deferred 5xx, and budget-aborted;
- run status: `success`, `partial`, `failed`, `cooldown`;
- deferred backlog count, oldest deferred age, attempts per deferred conversation;
- pages fetched, messages upserted, checkpoints advanced;
- sync duration, backlog depth/drain rate, Mongo write latency, projection queue lag, lease wait/contends.

**Confidence**

High on fixing 5xx aborts and keeping cap `1` for initial rollout. Medium on delta adoption because official docs support user-scoped Gov delta, not clearly per-chat delta; validate live before designing around it.


### Code Reviewer — Verifier · Codex (GPT-5.3-codex-spark)

## Recommendation

I mostly agree with the prior council, with two corrections: `deferred_failed` needs two-level checkpoint semantics, and cap `1` must be treated as a temporary rollout guardrail with backlog telemetry.

1. **Fix exhausted transient 5xx as conversation-level isolation, not run-fatal.**  
   The current behavior is too brittle: one exhausted `500/502` can abort a fully serial run and block unrelated conversations. Treat exhausted `500/502/503/504` as conversation-level `deferred_failed`, mark the run `partial`, and continue. Do **not** treat exhausted `429` the same way; that usually means tenant/app pressure and should trigger run cooldown or clean abort.

2. **Checkpoint rule:** advance run-level progress past the conversation, but do **not** advance that conversation’s message checkpoint/delta cursor.  
   If there is only a single monotonic checkpoint today, the design needs a deferred set/queue. Otherwise “do not advance checkpoint” can recreate head-of-line blocking.

3. **Failure budget is correct, but make it persistence-aware.**  
   Use a budget to distinguish isolated bad chats from systemic Graph failure, but avoid noisy single-run percentage alerts. Prefer: abort/cooldown only after an absolute floor plus ratio, and alert on deferred backlog that persists across runs. Example starter policy: clean abort if `deferred_transient_5xx >= 5` **and** `>= 10-20%` of attempted conversations, or if the same deferred conversations remain uncleared for `K` runs / `N` hours.

4. **Keep global cap `1` for rollout, but watch freshness.**  
   It is appropriate while topology is unknown because the Mongo lease is the only cross-process control over Graph pressure, Mongo writes, lease contention, and projection load. But it can become a throughput cliff. Raise to `2` only after a peak-load soak shows low throttling and stable downstream load, or if backlog/RPO is clearly failing at cap `1`.

5. **Do not build around per-chat delta yet.**  
   Microsoft documents chat message delta as user-scoped: `GET /users/{id}/chats/getAllMessages/delta`, available in US Gov L4/L5, with `$top` max `50` and an eight-month delta window. The documented per-chat list API is `/chats/{chat-id}/messages`; I do not see a documented v1.0 per-chat `/messages/delta` surface. Validate cheaply against `https://graph.microsoft.us`, but do not make it a rollout dependency. Sources: [chatMessage delta](https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0), [list chat messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0), [national cloud endpoints](https://learn.microsoft.com/en-us/graph/deployments).

6. **Defer global token bucket for now.**  
   With one cross-process sync lease and a serial message loop, current controls are enough for rollout. A distributed bucket becomes worth it before raising cap materially above `1`, or if telemetry shows `429` despite cap `1`.

## Implementation Notes

Persist deferred records with `conversationId`, tenant/user/job id, last status, Graph request id/correlation id if available, attempt count, first/last failure time, `nextRetryAt`, and last non-sensitive error class. Use per-conversation exponential backoff to prevent retry storms.

Minimal telemetry first:
- Graph requests by endpoint/status, retry count, retry sleep, exhausted retry count, `Retry-After` distribution.
- Conversations attempted/succeeded/skipped/deferred, run status `success|partial|failed|aborted_budget`.
- Deferred backlog count, oldest deferred age, attempts per deferred conversation.
- Sync backlog/RPO, pages fetched, messages upserted, duration.
- Mongo write latency/errors, lease wait/contention, projection enqueue/drain lag.

## Risks

Silent archive gaps if `partial` is treated as success or deferred backlog is invisible. Retry storms if deferred conversations retry every scheduled run without `nextRetryAt`. Alert fatigue if each chat failure pages instead of one aggregate tenant/run alert. Throughput starvation if cap `1` stays in place after evidence shows Graph and Mongo have headroom.

## Confidence

High on fixing transient 5xx isolation first, keeping cap `1` initially, and deferring token bucket. Medium on delta adoption details because GCC High tenant behavior still needs live validation even though user-scoped delta is documented.


---

## Stage 3 — Revisions

_No revisions: at standard depth no member changed position after the critique round._

---

## Stage 4 — Synthesis

# Teams Archive SYNC Hardening — Council Synthesis

## Council Recommendation
Ship a **durable, conversation-level `deferred_failed` path before touching throughput.** Exhausted transient 5xx (500/502/503/504) must isolate to the failing conversation — its checkpoint is *not* advanced — while the run continues and finishes as `partial` (never `success`). Treat exhausted **429 differently**: it signals tenant/app-wide pressure, so prefer clean run cooldown/abort over minting many per-chat deferrals. Keep `DEFAULT_MAX_CONCURRENT_SYNCS=1` as a **temporary** rollout guardrail, not a target steady state. Defer the global token bucket. Validate per-chat Graph delta cheaply and off the critical path, but do not design rollout around it. Add minimal archive-health telemetry alongside the deferral fix so the cap-raise decision is evidence-driven.

## Why This Council Was Selected
The task is a rollout-hardening decision spanning failure semantics, throughput control, an external API (Graph/GCC High) feasibility question, and operational telemetry — so it drew **strategy, backend, database, cloud, and code-review** lenses. That mix is what surfaced the load-bearing refinements: the database lens insisted deferral be first-class persisted state (not an in-memory skip), the cloud/backend lenses sourced the actual Graph delta documentation, and the strategy lens forced the topology-first and backlog-drain gaps the others reasoned past.

## Agreement
Unanimous across all members and both critique rounds:
- **500/502 abort is the top risk** and the highest-leverage fix. At cap=1 serial, one poison conversation causes head-of-line blocking that can wedge the entire archive — worse than a data gap.
- **`deferred_failed` is correct** if durable: persist `conversationId`, last error status/class, attempt count, first/last failure time, `nextRetryAt`, and Graph request-id; per-conversation exponential backoff; run marked `partial`.
- **Keep cap=1 for rollout**, raise only behind a feature flag, `1→2` (not back to 3), canary/soak through peak load.
- **Do not assume per-chat delta exists.** Microsoft documents *user-scoped* `GET /users/{id}/chats/getAllMessages/delta` (US Gov L4/L5, `$top`≤50, 8-month window), **not** `/chats/{chatId}/messages/delta`. Validate live against `https://graph.microsoft.us`; delta only helps steady-state incremental, never initial backfill.
- **Defer the token bucket** — no parallel fanout to coordinate at cap=1.
- **Three named failure modes**: silent archive gaps, retry storms, alert fatigue — all mitigated by visible `partial` state, `nextRetryAt` backoff, and aggregate (not per-chat) alerting.

## Disagreement
Material, unresolved tensions worth preserving:

1. **Failure-budget shape.** The prior council and most members endorsed the starter "abort if ≥5 conversations *or* ≥10%." The **strategy lens challenges this as dimensionally wrong**: a cumulative count fires spuriously on small tenants (5 of 10 = 50%) and never trips on large ones, and it conflates a *scattered* poison-chat pattern (defer-and-continue is right) with a *clustered/consecutive* outage signature (continuing just storms a sick endpoint). Its counter-proposal: **abort on N consecutive exhausted-5xx, with a percentage ceiling gated behind a minimum-sample floor, and escalate on persistence across K runs** rather than raw single-run count. The backend and code-review lenses partially converged here (floor + ratio + persistence). Treat the fixed "≥5 or ≥10%" as a *placeholder*, not a settled threshold.

2. **What gates raising the cap.** Most members listed only *rate-health* signals (429≈0, stable Retry-After, exhausted-5xx≈0). The **strategy lens adds a second, co-equal axis**: backlog-drain-rate vs. ingest-rate against the product RPO. Cap=1 can be a *guaranteed SLO miss in disguise* in a multi-replica deployment even when perfectly healthy. Raise only when rate says *safe* **and** throughput says *needed*; if throughput says needed but rate says unsafe, that's an architecture problem (per-tenant fairness / cross-process rate control), not a knob.

3. **Topology priority.** Every member noted "single vs. multi-replica unknown" then reasoned past it. The **strategy lens elevates it to a P0 prerequisite**: the 300ms page delay, 60s backfill throttle, and 10s DB gate are per-process and do not compose across replicas — only the lease cap is cross-process. The cap math and the token-bucket trigger are undecidable without this one-line ops answer.

## Minority Positions
- **(Strategy) Partial-aware projection is the true P0, not just telemetry.** For a *memory platform*, the catastrophic failure is the projection indexing an incomplete archive and presenting it as authoritative. The invariant is two-sided: deferred conversation → checkpoint frozen + run `partial`, **and** the projection must refuse to treat a `partial` run as complete (skip/mark-stale the deferred slices). Telemetry detects gaps; this invariant prevents gaps from masquerading as completeness. Held most strongly by the strategy lens; the backend lens flagged it as an assumption to verify.
- **(Strategy) Deprioritize delta harder than peers did.** Even a *successful* validation doesn't move the rollout needle, because delta is a throughput/cost optimization and the 8-month window can't replace backfill. Run a time-boxed probe, file the result, move on.
- **(Cloud/Code-review) Confirm where the 429s actually originate (backfill vs. steady-state) before betting any effort on delta** — if the storm is backfill-driven, delta is irrelevant to it.

## Risks And Unknowns
- **Silent archive gaps (highest)** — mitigated only by the partial-aware projection invariant plus visible deferred backlog, not by telemetry alone.
- **Two-level checkpoint bugs** — splitting run-progress from per-conversation cursor is more state; a buggy implementation could double-archive or skip. Mitigate with idempotent writes keyed on Graph message ID.
- **Budget-abort via thrown exception** — must use the clean/resumable shutdown path (release leases, persist deferred set), or you trade a data-gap bug for a stuck-lease bug.
- **Permanent backlog under cap=1** — "conservative" must not become "never catches up"; the drain SLO is the guardrail.
- **GCC High delta behavior** — documented surface is user-scoped; per-chat delta is unvalidated. *Medium confidence, requires live tenant test.*
- **Replica topology** — unresolved; changes per-process throttle math and the token-bucket trigger.
- **Lease correctness** — if atomic acquire/expiry has races, cap=1 partially masks a latent bug.

## Implementation Path
1. **P0 (free, do first):** Resolve single- vs. multi-replica with ops. One question; gates everything downstream.
2. **P0 (the fix):** Route exhausted 500/502/503/504 to a durable `deferred_failed` path (persisted in Mongo keyed by `chatId`, not per-process Maps). Keep `isRecoverableChatMessageError()` returning 403/404 as *skip*. Decouple run-level progress from the frozen per-conversation cursor. Mark run `partial`. Treat exhausted **429** as run-level cooldown/abort, not mass deferral. Ensure budget-abort uses the clean shutdown path.
3. **P0 (the invariant):** Make the enterprise-memory projection partial-aware — do not index a `partial` run as a complete snapshot.
4. **P1:** Implement the failure budget as **consecutive-failure + floored-percentage + cross-run persistence** (not raw cumulative count). Per-conversation exponential backoff via `nextRetryAt`; max-attempts cap escalating to "needs manual intervention."
5. **P1:** Ship minimal telemetry (below) *with* the fix.
6. **P2:** Time-boxed delta probe — one app-token call each to `/chats/{chatId}/messages/delta` and `/users/{id}/chats/getAllMessages/delta` against `https://graph.microsoft.us`; record status, headers, `@odata.deltaLink`, scopes, throttling, metering. File and move on.
7. **P2 (gated):** Only when rate-axis says *safe* and throughput-axis says *needed*, raise cap `1→2` by flag, one tenant, soak. A cross-process limiter is the explicit prerequisite for cap>1 in a confirmed multi-replica deployment.

**Minimal telemetry (decision-gating only):** run-status enum (`success|partial|failed|aborted_budget`); deferred registry with `oldest_deferred_age` + attempts-per-conversation; consecutive-exhausted-5xx counter + exhausted-retry count by status; **backlog-drain-rate vs. ingest-rate**; Graph requests by endpoint/status + total Retry-After/backoff sleep seconds; projection enqueue/drain + a flag for whether projection saw a `partial` run. **Alert/page only on RPO-freshness breach and budget abort**; everything else is a dashboard.

## Confidence
**High:** 500/502 isolation is the top risk and the right first fix; durable per-conversation deferral with frozen checkpoint + `partial` run; cap=1 as temporary guardrail; defer token bucket; validate-don't-assume delta with user-scoped (not per-chat) being the documented surface; the telemetry set. **High** on the consecutive-vs-cumulative budget refinement and topology-first being genuine improvements over the prior council. **Medium:** cap-raise sequencing and the backlog-drain math (depends on tenant size, per-message latency, and replica topology). **Low/uncertain:** GCC High per-chat delta behavior — hence "cheap spike, off critical path," not adopt-or-reject.

## Next Step
**Resolve replica topology (one ops question), then implement durable `deferred_failed` + `partial`-run semantics + partial-aware projection, shipping the minimal telemetry in the same change.** This removes the whole-run wedge — the only failure that can take the archive offline on a single flaky conversation — and produces the evidence needed to safely tune throughput. Do not raise the cap until that telemetry is in hand.
