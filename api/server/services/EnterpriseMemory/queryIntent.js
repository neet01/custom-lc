const INTENT_BOOST_WEIGHT = 3;
const TOPIC_COVERAGE_WEIGHT = 2;
const RECENCY_WEIGHT = 0.5;

const COMMITMENT_PHRASES = [
  "i'll",
  'i will',
  'i am going to',
  "i'm going to",
  'im going to',
  "i'm gonna",
  'i plan to',
  'i intend to',
  'i can take',
  'i can handle',
  'i can own',
  "i'll take",
  "i'll handle",
  "i'll own",
  "i'll get",
  "i'll do",
  "i'll send",
  "i'll set up",
  "i'll put together",
  "i'll draft",
  "i'll write",
  "i'll review",
  "i'll follow up",
  "i'll circle back",
  "i'll get back",
  "i'll ping",
  "i'll reach out",
  "i'll cover",
  'let me take',
  'let me handle',
  'will do',
  'on it',
  'i got this',
  'i got it',
  'action item',
  'by eod',
  'by end of day',
  'by tomorrow',
];

const RESPONSIBILITY_PHRASES = [
  'responsible for',
  "i'm responsible",
  'i am responsible',
  'my responsibility',
  'in charge of',
  'i own',
  "i'll own",
  'owner of',
  'owns the',
  'i lead',
  "i'm leading",
  "i'll lead",
  'accountable for',
  'my job',
  'assigned to me',
  "i've got",
  'ive got',
  "that's mine",
  'thats mine',
  'i handle',
  "i'll handle",
  'i take care of',
  "i'll take care of",
  'taking ownership',
  'point person',
  'i cover',
  "i'll cover",
  "i'm the owner",
  "i'm the lead",
  'dri',
];

const SELF_COMMITMENT_PATTERNS = [
  /what did i (tell|say|promise|commit|agree)/,
  /what have i (told|said|promised|committed|agreed)/,
  /what i (said|told)[\s\S]*(do|i'?d|i would)/,
  /things i (said|promised|committed|agreed)/,
  /did i (say|tell|promise|commit|agree)/,
  /what did i agree to/,
  /my (commitments|action items|todos|to-dos|promises)/,
  /what (am i|do i) (need to|have to|supposed to) (do|deliver|finish|send|ship)/,
];

const SELF_ACTIVITY_PATTERNS = [
  /\b(my|recent|latest|last)\b[\s\S]*\b(outgoing|sent)\b[\s\S]*\bmessages?\b/,
  /\b(outgoing|sent) messages?\b/,
  /\bmessages? i (sent|wrote|posted|put out)\b/,
  /\bmy (recent|latest|last|outgoing|sent) messages?\b/,
  /\bwhat (have|did) i (post|posted|send|sent|write|wrote)\b/,
  /\bmy (recent|latest) (slack )?(activity|posts|messages)\b/,
];

const GENERIC_COMMITMENT_PATTERNS = [
  /\bcommit(ment|ments|ted)?\b/,
  /\baction items?\b/,
  /\bwho (is|'s) going to\b/,
  /\bwhat (was|were) (promised|agreed)\b/,
];

const RESPONSIBILITY_PATTERNS = [
  /who(\s|'s| is| are)?[\s\S]*(responsible|in charge|owns?|owner|handl|lead|accountable|on the hook|point person|\bdri\b)/,
  /who (owns|handles|leads|manages)\b/,
  /who (is|'s) (the )?(owner|lead|dri|point person)\b/,
  /who said (they|he|she|we)[\s\S]*(responsible|own|handle|lead|in charge|cover|take)/,
];

const TIME_CUES = [
  { pattern: /\btoday\b/, daysBack: 1 },
  { pattern: /\bthis (morning|afternoon|evening)\b/, daysBack: 1 },
  { pattern: /\byesterday\b/, daysBack: 2 },
  { pattern: /\b(this|past|last) week\b/, daysBack: 7 },
  { pattern: /\b(this|past|last) month\b/, daysBack: 31 },
  { pattern: /\b(recently|lately)\b/, daysBack: 14 },
];

const TOPIC_STOP_WORDS = new Set([
  'what', 'who', 'whom', 'whose', 'when', 'where', 'why', 'how', 'which',
  'did', 'do', 'does', 'have', 'has', 'had', 'was', 'were', 'are', 'is', 'be', 'been',
  'i', 'me', 'my', 'we', 'our', 'you', 'your', 'they', 'them', 'their', 'he', 'she', 'it',
  'said', 'say', 'says', 'tell', 'told', 'telling', 'mention', 'mentioned', 'claim', 'claimed',
  'promise', 'promised', 'commit', 'committed', 'commitment', 'commitments', 'agree', 'agreed',
  'responsible', 'responsibility', 'own', 'owns', 'owner', 'owning', 'handle', 'handles',
  'handling', 'lead', 'leads', 'leading', 'charge', 'accountable', 'manage', 'manages',
  'would', 'will', 'going', 'gonna', 'gonna', 'should', 'could', 'can', 'plan', 'planning',
  'about', 'with', 'from', 'into', 'that', 'this', 'those', 'these', 'and', 'the', 'for', 'to',
  'people', 'everyone', 'everybody', 'team', 'teams', 'folks', 'guys', 'someone', 'anyone',
  'today', 'yesterday', 'week', 'month', 'morning', 'afternoon', 'evening', 'recently', 'lately',
  'past', 'last', 'recent', 'latest', 'most', 'message', 'messages', 'chat', 'chats',
  'discussion', 'discussed', 'outgoing', 'sent', 'send', 'post', 'posted', 'posting', 'posts',
  'wrote', 'write', 'put', 'activity',
  'show', 'find', 'search', 'look', 'get', 'give', 'around', 'regarding', 'related',
]);

function normalizeText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[‘’ʼ`]/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeRegex(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function matchesAny(value, patterns) {
  return patterns.some((pattern) => pattern.test(value));
}

function extractTopicTerms(normalizedQuery) {
  return [
    ...new Set(
      normalizedQuery
        .split(/[^a-z0-9._-]+/i)
        .map((term) => term.trim())
        .filter((term) => term.length >= 3 && !TOPIC_STOP_WORDS.has(term)),
    ),
  ].slice(0, 8);
}

function detectImpliedDaysBack(normalizedQuery) {
  let impliedDaysBack;
  for (const cue of TIME_CUES) {
    if (cue.pattern.test(normalizedQuery)) {
      impliedDaysBack =
        impliedDaysBack === undefined ? cue.daysBack : Math.min(impliedDaysBack, cue.daysBack);
    }
  }
  return impliedDaysBack;
}

function classifyArchetype(normalizedQuery) {
  const hasWho = /\bwho\b/.test(normalizedQuery);

  if (hasWho && matchesAny(normalizedQuery, RESPONSIBILITY_PATTERNS)) {
    return { archetype: 'responsibility', selfAuthored: false };
  }

  if (matchesAny(normalizedQuery, SELF_ACTIVITY_PATTERNS)) {
    return { archetype: 'self_activity', selfAuthored: true };
  }

  if (matchesAny(normalizedQuery, SELF_COMMITMENT_PATTERNS)) {
    return { archetype: 'commitment', selfAuthored: true };
  }

  if (matchesAny(normalizedQuery, RESPONSIBILITY_PATTERNS)) {
    return { archetype: 'responsibility', selfAuthored: false };
  }

  if (matchesAny(normalizedQuery, GENERIC_COMMITMENT_PATTERNS)) {
    return { archetype: 'commitment', selfAuthored: false };
  }

  return { archetype: 'general', selfAuthored: false };
}

function getIntentPhrases(archetype) {
  if (archetype === 'commitment') {
    return COMMITMENT_PHRASES;
  }
  if (archetype === 'responsibility') {
    return RESPONSIBILITY_PHRASES;
  }
  return [];
}

/**
 * Classifies a natural-language Slack search query into a retrieval intent.
 * Pure and deterministic so retrieval can infer sender scope, timeframe, and
 * paraphrase expansion without depending on the caller setting them explicitly.
 */
function analyzeSlackQuery(query) {
  const normalizedQuery = normalizeText(query);
  if (!normalizedQuery) {
    return {
      archetype: 'general',
      selfAuthored: false,
      impliedSenderScope: undefined,
      impliedDaysBack: undefined,
      intentPhrases: [],
      topicTerms: [],
    };
  }

  const { archetype, selfAuthored } = classifyArchetype(normalizedQuery);

  return {
    archetype,
    selfAuthored,
    impliedSenderScope: selfAuthored ? 'me' : undefined,
    impliedDaysBack: detectImpliedDaysBack(normalizedQuery),
    intentPhrases: getIntentPhrases(archetype),
    topicTerms: extractTopicTerms(normalizedQuery),
  };
}

function buildIntentChunkClause(intentPhrases, fields = ['text', 'summary', 'title']) {
  if (!Array.isArray(intentPhrases) || intentPhrases.length === 0) {
    return null;
  }

  const phraseClauses = intentPhrases.flatMap((phrase) => {
    const normalized = normalizeText(phrase);
    if (!normalized) {
      return [];
    }
    const regex = new RegExp(escapeRegex(normalized).replace(/ /g, '\\s+'), 'i');
    return fields.map((field) => ({ [field]: regex }));
  });

  return phraseClauses.length > 0 ? { $or: phraseClauses } : null;
}

function findMatchedIntentPhrase(normalizedChunkText, intentPhrases) {
  for (const phrase of intentPhrases) {
    const normalizedPhrase = normalizeText(phrase);
    if (normalizedPhrase && normalizedChunkText.includes(normalizedPhrase)) {
      return phrase;
    }
  }
  return null;
}

function computeTopicCoverage(normalizedChunkText, topicTerms) {
  if (!Array.isArray(topicTerms) || topicTerms.length === 0) {
    return 0;
  }
  const matched = topicTerms.filter((term) => normalizedChunkText.includes(term)).length;
  return matched / topicTerms.length;
}

function getChunkTime(chunk) {
  const value = chunk?.sourceTimestamp;
  if (!value) {
    return 0;
  }
  const time = value instanceof Date ? value.getTime() : new Date(value).getTime();
  return Number.isFinite(time) ? time : 0;
}

/**
 * Re-ranks candidate chunks by a composite of paraphrase-intent match, topic-term
 * coverage, and recency, returning the chunks annotated with their relevance score
 * and the matched intent phrase. Deterministic, with timestamp as the stable tiebreaker.
 */
function rerankSlackChunks(chunks, analysis, { limit } = {}) {
  const intentPhrases = analysis?.intentPhrases || [];
  const topicTerms = analysis?.topicTerms || [];
  const times = chunks.map(getChunkTime);
  const minTime = times.length ? Math.min(...times) : 0;
  const maxTime = times.length ? Math.max(...times) : 0;
  const timeSpan = maxTime - minTime;

  const scored = chunks.map((chunk, index) => {
    const normalizedChunkText = normalizeText(
      `${chunk?.text || ''} ${chunk?.summary || ''} ${chunk?.title || ''}`,
    );
    const matchedIntent = findMatchedIntentPhrase(normalizedChunkText, intentPhrases);
    const topicCoverage = computeTopicCoverage(normalizedChunkText, topicTerms);
    const recencyNorm = timeSpan > 0 ? (times[index] - minTime) / timeSpan : 0;
    const relevanceScore =
      (matchedIntent ? INTENT_BOOST_WEIGHT : 0) +
      TOPIC_COVERAGE_WEIGHT * topicCoverage +
      RECENCY_WEIGHT * recencyNorm;

    return { chunk, index, matchedIntent, topicCoverage, relevanceScore, time: times[index] };
  });

  scored.sort((a, b) => {
    if (b.relevanceScore !== a.relevanceScore) {
      return b.relevanceScore - a.relevanceScore;
    }
    if (b.time !== a.time) {
      return b.time - a.time;
    }
    return a.index - b.index;
  });

  const limited = typeof limit === 'number' && limit > 0 ? scored.slice(0, limit) : scored;
  return limited.map(({ chunk, matchedIntent, topicCoverage, relevanceScore }) => ({
    chunk,
    matchedIntent,
    topicCoverage,
    relevanceScore: Math.round(relevanceScore * 1000) / 1000,
  }));
}

module.exports = {
  analyzeSlackQuery,
  buildIntentChunkClause,
  rerankSlackChunks,
};
