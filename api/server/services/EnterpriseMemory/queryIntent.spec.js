const { analyzeSlackQuery, buildIntentChunkClause, rerankSlackChunks } = require('./queryIntent');

describe('queryIntent.analyzeSlackQuery', () => {
  it('classifies a self-authored commitment query and infers "me" + today', () => {
    const analysis = analyzeSlackQuery('what did I tell people I would do today');

    expect(analysis.archetype).toBe('commitment');
    expect(analysis.selfAuthored).toBe(true);
    expect(analysis.impliedSenderScope).toBe('me');
    expect(analysis.impliedDaysBack).toBe(1);
    expect(analysis.intentPhrases).toEqual(expect.arrayContaining(["i'll", 'i will']));
    expect(analysis.topicTerms).not.toContain('people');
  });

  it('classifies a responsibility/ownership query without forcing sender scope', () => {
    const analysis = analyzeSlackQuery('who said they are responsible for landing gear');

    expect(analysis.archetype).toBe('responsibility');
    expect(analysis.selfAuthored).toBe(false);
    expect(analysis.impliedSenderScope).toBeUndefined();
    expect(analysis.intentPhrases).toEqual(expect.arrayContaining(['responsible for', 'i own']));
    expect(analysis.topicTerms).toEqual(expect.arrayContaining(['landing', 'gear']));
    expect(analysis.topicTerms).not.toContain('responsible');
  });

  it('leaves an ordinary topic query as a general archetype with no expansion', () => {
    const analysis = analyzeSlackQuery('budget approval timeline');

    expect(analysis.archetype).toBe('general');
    expect(analysis.selfAuthored).toBe(false);
    expect(analysis.impliedSenderScope).toBeUndefined();
    expect(analysis.intentPhrases).toEqual([]);
    expect(analysis.topicTerms).toEqual(expect.arrayContaining(['budget', 'approval', 'timeline']));
  });

  it('detects week-scoped timeframes', () => {
    expect(analyzeSlackQuery('what did I commit to this week').impliedDaysBack).toBe(7);
  });

  it('handles empty input safely', () => {
    const analysis = analyzeSlackQuery('');
    expect(analysis.archetype).toBe('general');
    expect(analysis.topicTerms).toEqual([]);
  });
});

describe('queryIntent.buildIntentChunkClause', () => {
  it('builds a phrase-aware $or clause across text fields', () => {
    const clause = buildIntentChunkClause(['responsible for']);
    expect(clause).toHaveProperty('$or');
    expect(clause.$or).toHaveLength(3);
    const textClause = clause.$or.find((entry) => entry.text);
    expect(textClause.text.test('I am responsible  for landing gear')).toBe(true);
  });

  it('returns null when there are no intent phrases', () => {
    expect(buildIntentChunkClause([])).toBeNull();
  });
});

describe('queryIntent.rerankSlackChunks', () => {
  it('ranks an intent+topic match above topic-only and intent-only matches', () => {
    const analysis = analyzeSlackQuery('who said they are responsible for landing gear');
    const chunks = [
      { _id: 'topic-only', text: 'the landing gear inspection is overdue', sourceTimestamp: new Date('2026-06-01') },
      { _id: 'intent-and-topic', text: "I'll own landing gear from here", sourceTimestamp: new Date('2026-06-02') },
      { _id: 'intent-only', text: "I'll own the deployment pipeline", sourceTimestamp: new Date('2026-06-03') },
    ];

    const ranked = rerankSlackChunks(chunks, analysis, { limit: 3 });

    expect(ranked[0].chunk._id).toBe('intent-and-topic');
    expect(ranked[0].matchedIntent).toBeTruthy();
    expect(ranked[0].relevanceScore).toBeGreaterThan(ranked[1].relevanceScore);
  });

  it('surfaces recent commitments first when there is no topic signal', () => {
    const analysis = analyzeSlackQuery('what did I tell people I would do today');
    const chunks = [
      { _id: 'greeting', text: 'good morning everyone', sourceTimestamp: new Date('2026-06-02T09:00:00Z') },
      { _id: 'commitment', text: "I'll send the status report by EOD", sourceTimestamp: new Date('2026-06-02T08:00:00Z') },
    ];

    const ranked = rerankSlackChunks(chunks, analysis, { limit: 2 });

    expect(ranked[0].chunk._id).toBe('commitment');
  });
});
