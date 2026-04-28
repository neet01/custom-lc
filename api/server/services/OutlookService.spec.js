jest.mock('@librechat/api', () => ({
  isEnabled: (value) => value === true || value === 'true' || value === '1',
}));

jest.mock('~/server/services/GraphTokenService', () => ({
  getGraphApiToken: jest.fn(),
}));

jest.mock('~/server/services/OutlookAIService', () => ({
  isModelBackedAIEnabled: jest.fn(() => false),
  generateAnalysis: jest.fn(),
  generateSelectionBrief: jest.fn(),
  generateDailyBrief: jest.fn(),
  generateReplyDraft: jest.fn(),
  generateMeetingInviteNote: jest.fn(),
  logModelFailure: jest.fn(),
}));

jest.mock('@librechat/data-schemas', () => ({
  logger: {
    warn: jest.fn(),
    error: jest.fn(),
  },
}));

const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const OutlookAIService = require('~/server/services/OutlookAIService');
const OutlookService = require('./OutlookService');

const user = {
  id: 'user-1',
  provider: 'openid',
  openidId: 'entra-user',
  federatedTokens: {
    access_token: 'openid-access-token',
  },
};

describe('OutlookService', () => {
  const originalEnv = process.env;
  const originalFetch = global.fetch;

  beforeEach(() => {
    jest.clearAllMocks();
    OutlookAIService.isModelBackedAIEnabled.mockReturnValue(false);
    OutlookAIService.generateMeetingInviteNote.mockResolvedValue(null);
    process.env = {
      ...originalEnv,
      OUTLOOK_AI_ENABLED: 'true',
      OPENID_REUSE_TOKENS: 'true',
      OUTLOOK_GRAPH_BASE_URL: 'https://graph.microsoft.us',
      OUTLOOK_GRAPH_SCOPES: 'https://graph.microsoft.us/.default',
    };
    getGraphApiToken.mockResolvedValue({ access_token: 'graph-token' });
    global.fetch = jest.fn();
  });

  afterEach(() => {
    process.env = originalEnv;
    global.fetch = originalFetch;
  });

  it('reports delegated connection status without exchanging a token', () => {
    const status = OutlookService.getStatus(user);

    expect(status).toMatchObject({
      enabled: true,
      connected: true,
      graphBaseUrl: 'https://graph.microsoft.us/v1.0',
      scopes: 'https://graph.microsoft.us/.default',
    });
    expect(getGraphApiToken).not.toHaveBeenCalled();
  });

  it('rejects mailbox access when the feature is disabled', async () => {
    process.env.OUTLOOK_AI_ENABLED = 'false';

    await expect(OutlookService.listMessages(user)).rejects.toMatchObject({
      name: 'OutlookServiceError',
      status: 403,
    });
  });

  it('lists recent inbox messages from the GCC High Graph endpoint', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        value: [
          {
            id: 'message-1',
            conversationId: 'thread-1',
            subject: 'Budget follow-up',
            from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
            receivedDateTime: '2026-04-21T12:00:00Z',
            bodyPreview: 'Please review the runway numbers.',
            importance: 'high',
            inferenceClassification: 'focused',
            isRead: false,
            hasAttachments: true,
            webLink: 'https://outlook.example/message-1',
          },
        ],
      }),
    });

    const result = await OutlookService.listMessages(user, { limit: 10 });

    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]).toMatchObject({
      id: 'message-1',
      subject: 'Budget follow-up',
      from: { name: 'Finance', address: 'finance@example.mil' },
      hasAttachments: true,
      inferenceClassification: 'focused',
    });
    expect(global.fetch).toHaveBeenCalledWith(
      expect.objectContaining({
        origin: 'https://graph.microsoft.us',
        pathname: '/v1.0/me/mailFolders/inbox/messages',
      }),
      expect.objectContaining({
        headers: expect.objectContaining({
          Authorization: 'Bearer graph-token',
        }),
      }),
    );
    const requestedUrl = global.fetch.mock.calls[0][0];
    expect(requestedUrl.searchParams.get('$filter')).toBeNull();
  });

  it('filters other inbox messages in app to avoid complex Graph filter/order queries', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        value: [
          {
            id: 'focused-message',
            subject: 'Internal note',
            from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
            bodyPreview: 'Focused item',
            inferenceClassification: 'focused',
          },
          {
            id: 'other-message',
            subject: 'Vendor note',
            from: { emailAddress: { name: 'Vendor', address: 'vendor@example.com' } },
            bodyPreview: 'Other item',
            inferenceClassification: 'other',
          },
        ],
      }),
    });

    const result = await OutlookService.listMessages(user, { inboxView: 'other', limit: 10 });

    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]).toMatchObject({
      id: 'other-message',
      inferenceClassification: 'other',
    });
    const requestedUrl = global.fetch.mock.calls[0][0];
    expect(requestedUrl.searchParams.get('$filter')).toBeNull();
  });

  it('searches inbox messages through Microsoft Graph search', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        value: [
          {
            id: 'older-match',
            subject: 'Vendor renewal',
            from: { emailAddress: { name: 'Vendor', address: 'rep@vendor.com' } },
            receivedDateTime: '2026-04-19T12:00:00Z',
            bodyPreview: 'Renewal terms attached.',
            inferenceClassification: 'other',
          },
          {
            id: 'newer-match',
            subject: 'Vendor renewal follow-up',
            from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
            receivedDateTime: '2026-04-21T12:00:00Z',
            bodyPreview: 'Please review before Friday.',
            inferenceClassification: 'focused',
          },
        ],
      }),
    });

    const result = await OutlookService.listMessages(user, {
      inboxView: 'focused',
      limit: 10,
      search: 'vendor renewal',
    });

    expect(result.search).toBe('vendor renewal');
    expect(result.messages.map((message) => message.id)).toEqual(['newer-match', 'older-match']);
    const requestedUrl = global.fetch.mock.calls[0][0];
    expect(requestedUrl.searchParams.get('$search')).toBe('"vendor renewal"');
    expect(requestedUrl.searchParams.get('$orderby')).toBeNull();
    expect(requestedUrl.searchParams.get('$top')).toBe('10');
  });

  it('loads selected messages with conversation thread context', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'latest-message',
          conversationId: 'thread-1',
          subject: 'Budget follow-up',
          from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
          receivedDateTime: '2026-04-21T13:00:00Z',
          body: {
            contentType: 'html',
            content:
              '<div><strong>Latest</strong> thread note.<img src="https://vendor.test/logo.png"></div>',
          },
          bodyPreview: 'Latest thread note.',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'latest-message',
              conversationId: 'thread-1',
              subject: 'Budget follow-up',
              from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
              receivedDateTime: '2026-04-21T13:00:00Z',
              body: {
                contentType: 'html',
                content: '<div><strong>Latest</strong> thread note.</div>',
              },
            },
            {
              id: 'first-message',
              conversationId: 'thread-1',
              subject: 'Budget follow-up',
              from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
              receivedDateTime: '2026-04-21T12:00:00Z',
              body: {
                contentType: 'html',
                content: '<p>Original&nbsp;thread note.</p>',
              },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'draft-message',
              conversationId: 'thread-1',
              subject: 'RE: Budget follow-up',
              toRecipients: [{ emailAddress: { name: 'Finance', address: 'finance@example.mil' } }],
              ccRecipients: [{ emailAddress: { name: 'Ops', address: 'ops@example.mil' } }],
              bodyPreview: 'Draft response preview',
              isDraft: true,
              createdDateTime: '2026-04-21T13:10:00Z',
              lastModifiedDateTime: '2026-04-21T13:12:00Z',
              webLink: 'https://outlook.example/draft-message',
            },
          ],
        }),
      });

    const result = await OutlookService.getMessage(user, 'latest-message');

    expect(result.threadMessageCount).toBe(2);
    expect(result.thread.map((message) => message.id)).toEqual(['first-message', 'latest-message']);
    expect(result.draftReplyCount).toBe(1);
    expect(result.draftReplies.map((message) => message.id)).toEqual(['draft-message']);
    expect(result.draftReplies[0].toRecipients).toEqual([
      { name: 'Finance', address: 'finance@example.mil' },
    ]);
    expect(result.body).toBe('Latest thread note.');
    expect(result.bodyHtml).toContain('vendor.test/logo.png');
    expect(result.thread[0].body).toBe('Original thread note.');
    expect(global.fetch).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        pathname: '/v1.0/me/messages',
      }),
      expect.any(Object),
    );
    expect(global.fetch.mock.calls[0][1].headers.Prefer).toBe('outlook.body-content-type="html"');
    expect(global.fetch.mock.calls[1][1].headers.Prefer).toBe('outlook.body-content-type="html"');
    expect(global.fetch.mock.calls[2][1].headers.Prefer).toBe('outlook.body-content-type="html"');
    const threadUrl = global.fetch.mock.calls[1][0];
    expect(threadUrl.searchParams.get('$filter')).toBe(
      "conversationId eq 'thread-1' and isDraft eq false",
    );
    expect(threadUrl.searchParams.get('$orderby')).toBeNull();
    const draftUrl = global.fetch.mock.calls[2][0];
    expect(draftUrl.searchParams.get('$filter')).toBe(
      "conversationId eq 'thread-1' and isDraft eq true",
    );
    expect(draftUrl.searchParams.get('$orderby')).toBe('lastModifiedDateTime desc');
  });

  it('caps mailbox list requests at 100 messages', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({ value: [] }),
    });

    await OutlookService.listMessages(user, { limit: 250, inboxView: 'all' });

    const requestedUrl = global.fetch.mock.calls[0][0];
    expect(requestedUrl.searchParams.get('$top')).toBe('100');
  });

  it('deletes a message through Microsoft Graph', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      status: 204,
    });

    const result = await OutlookService.deleteMessage(user, 'message-to-delete');

    expect(result).toMatchObject({
      messageId: 'message-to-delete',
      message: 'Email moved to Deleted Items.',
    });
    expect(global.fetch).toHaveBeenCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/messages/message-to-delete',
      }),
      expect.objectContaining({
        method: 'DELETE',
      }),
    );
  });

  it('proposes meeting slots from thread attendees without creating an event', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Schedule budget review',
          from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
          toRecipients: [{ emailAddress: { name: 'Test User', address: 'test.user@example.mil' } }],
          body: { content: 'Can we find time?' },
          bodyPreview: 'Can we find time?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Schedule budget review',
              from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
              toRecipients: [
                { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
              ],
              body: { content: 'Can we find time?' },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          timeZone: 'Pacific Standard Time',
          workingHours: {
            daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
            startTime: '08:30:00.0000000',
            endTime: '16:30:00.0000000',
            timeZone: { name: 'Pacific Standard Time' },
          },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          meetingTimeSuggestions: [
            {
              confidence: 90,
              meetingTimeSlot: {
                start: { dateTime: '2026-04-23T17:00:00.0000000', timeZone: 'UTC' },
                end: { dateTime: '2026-04-23T17:30:00.0000000', timeZone: 'UTC' },
              },
              suggestionReason: 'Suggested because everyone is free.',
            },
            {
              confidence: 90,
              meetingTimeSlot: {
                start: { dateTime: '2026-04-24T02:00:00.0000000', timeZone: 'UTC' },
                end: { dateTime: '2026-04-24T02:30:00.0000000', timeZone: 'UTC' },
              },
              suggestionReason: 'Late slot should be filtered out.',
            },
          ],
        }),
      });

    const result = await OutlookService.proposeMeetingSlots(user, 'source-message', {
      durationMinutes: 30,
    });

    expect(result.attendees).toEqual([{ name: 'Finance', address: 'finance@example.mil' }]);
    expect(result.suggestions).toHaveLength(1);
    expect(result.suggestions[0]).toMatchObject({ confidence: 90 });
    expect(global.fetch).toHaveBeenLastCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/findMeetingTimes',
      }),
      expect.objectContaining({
        method: 'POST',
        body: expect.stringContaining('finance@example.mil'),
        headers: expect.objectContaining({
          Prefer: 'outlook.timezone="Pacific Standard Time"',
        }),
      }),
    );
    const requestBody = JSON.parse(global.fetch.mock.calls[5][1].body);
    expect(requestBody.timeConstraint.timeslots[0].start).toMatchObject({
      dateTime: expect.stringContaining('T08:30:00'),
      timeZone: 'Pacific Standard Time',
    });
    expect(requestBody.timeConstraint.timeslots[0].end).toMatchObject({
      dateTime: expect.stringContaining('T16:30:00'),
      timeZone: 'Pacific Standard Time',
    });
  });

  it('creates a calendar-backed Teams meeting with an AI-generated invite note', async () => {
    const slot = {
      start: { dateTime: '2026-04-23T17:00:00.0000000', timeZone: 'UTC' },
      end: { dateTime: '2026-04-23T17:30:00.0000000', timeZone: 'UTC' },
    };
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Schedule budget review',
          from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
          toRecipients: [{ emailAddress: { name: 'Test User', address: 'test.user@example.mil' } }],
          body: { content: 'Can we find time?' },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Schedule budget review',
              from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
              toRecipients: [
                { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
              ],
              body: { content: 'Can we find time?' },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({
          id: 'event-1',
          subject: 'Meeting: Schedule budget review',
          start: slot.start,
          end: slot.end,
          webLink: 'https://outlook.example/event-1',
          onlineMeeting: {
            joinUrl: 'https://teams.example/join',
          },
        }),
      });

    const result = await OutlookService.createTeamsMeeting(user, 'source-message', {
      slot,
      subject: 'Meeting: Schedule budget review',
      sendInvites: true,
    });

    expect(result.event.onlineMeeting.joinUrl).toBe('https://teams.example/join');
    expect(result.draft).toBeUndefined();
    expect(result.message).toBe('Teams meeting invite sent to attendees.');
    expect(result.meetingNotePreview).toContain('Objective: Meeting: Schedule budget review.');
    expect(result.meetingNotePreview).toContain('Context:');
    expect(result.meetingDraft).toMatchObject({
      id: 'event-1',
      subject: 'Meeting: Schedule budget review',
      webLink: 'https://outlook.example/event-1',
    });
    expect(global.fetch).toHaveBeenNthCalledWith(
      5,
      expect.objectContaining({
        pathname: '/v1.0/me/events',
      }),
      expect.objectContaining({
        method: 'POST',
        body: expect.stringContaining('"isOnlineMeeting":true'),
      }),
    );
    const eventPayload = JSON.parse(global.fetch.mock.calls[4][1].body);
    expect(eventPayload.onlineMeetingProvider).toBe('teamsForBusiness');
    expect(eventPayload.isOnlineMeeting).toBe(true);
    expect(eventPayload.body.content).toContain('<strong>Meeting brief:</strong>');
    expect(eventPayload.body.content).toContain('Objective: Meeting: Schedule budget review.');
    expect(eventPayload.attendees).toEqual([
      {
        type: 'required',
        emailAddress: {
          name: 'Finance',
          address: 'finance@example.mil',
        },
      },
    ]);
    expect(global.fetch).toHaveBeenCalledTimes(5);
  });

  it('schedules against internal attendees and annotates external availability behavior', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'License review sync',
          from: { emailAddress: { name: 'Vendor Rep', address: 'rep@vendor.com' } },
          toRecipients: [{ emailAddress: { name: 'Test User', address: 'test.user@example.mil' } }],
          ccRecipients: [
            { emailAddress: { name: 'Finance Lead', address: 'finance@example.mil' } },
          ],
          body: { content: 'I am available Tuesday at 2pm ET.' },
          bodyPreview: 'I am available Tuesday at 2pm ET.',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'License review sync',
              from: { emailAddress: { name: 'Vendor Rep', address: 'rep@vendor.com' } },
              toRecipients: [
                { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
              ],
              ccRecipients: [
                { emailAddress: { name: 'Finance Lead', address: 'finance@example.mil' } },
              ],
              body: { content: 'I am available Tuesday at 2pm ET.' },
              bodyPreview: 'I am available Tuesday at 2pm ET.',
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          timeZone: 'Pacific Standard Time',
          workingHours: {
            daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
            startTime: '08:30:00.0000000',
            endTime: '16:30:00.0000000',
            timeZone: { name: 'Pacific Standard Time' },
          },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          meetingTimeSuggestions: [
            {
              confidence: 92,
              meetingTimeSlot: {
                start: { dateTime: '2026-04-23T17:00:00.0000000', timeZone: 'UTC' },
                end: { dateTime: '2026-04-23T17:30:00.0000000', timeZone: 'UTC' },
              },
              suggestionReason: 'Suggested because internal attendees are free.',
            },
          ],
        }),
      });

    const result = await OutlookService.proposeMeetingSlots(user, 'source-message', {
      durationMinutes: 30,
    });

    expect(result.attendees).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'finance@example.mil' }),
        expect.objectContaining({ address: 'rep@vendor.com' }),
      ]),
    );
    expect(result.schedulingAttendees).toEqual([
      expect.objectContaining({ address: 'finance@example.mil' }),
    ]);
    expect(result.externalAttendeesExcluded).toEqual([]);
    expect(result.externalAttendeesWithThreadAvailability).toEqual([
      expect.objectContaining({ address: 'rep@vendor.com' }),
    ]);
    expect(result.availabilityNotes).toEqual([
      expect.stringContaining('Thread-stated availability detected for external attendees'),
    ]);

    const requestBody = JSON.parse(global.fetch.mock.calls[5][1].body);
    const requestAddresses = requestBody.attendees.map((attendee) => attendee.emailAddress.address);
    expect(requestAddresses).toEqual(['finance@example.mil']);
  });

  it('reduces meeting confidence when the suggested time has tentative conflicts', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Schedule budget review',
          from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
          toRecipients: [{ emailAddress: { name: 'Test User', address: 'test.user@example.mil' } }],
          body: { content: 'Can we find time?' },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Schedule budget review',
              from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
              toRecipients: [
                { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
              ],
              body: { content: 'Can we find time?' },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          timeZone: 'Pacific Standard Time',
          workingHours: {
            daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
            startTime: '08:30:00.0000000',
            endTime: '16:30:00.0000000',
            timeZone: { name: 'Pacific Standard Time' },
          },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          meetingTimeSuggestions: [
            {
              confidence: 100,
              organizerAvailability: 'tentative',
              attendeeAvailability: [
                {
                  attendee: {
                    emailAddress: { name: 'Finance', address: 'finance@example.mil' },
                    type: 'required',
                  },
                  availability: 'tentative',
                },
              ],
              meetingTimeSlot: {
                start: { dateTime: '2026-04-23T17:00:00.0000000', timeZone: 'UTC' },
                end: { dateTime: '2026-04-23T17:30:00.0000000', timeZone: 'UTC' },
              },
              suggestionReason: 'Suggested because everyone is free.',
            },
          ],
        }),
      });

    const result = await OutlookService.proposeMeetingSlots(user, 'source-message', {
      durationMinutes: 30,
    });

    expect(result.suggestions).toHaveLength(1);
    expect(result.suggestions[0]).toMatchObject({
      confidence: 65,
      confidenceReason: expect.stringContaining('organizer is tentatively busy'),
    });
    expect(result.suggestions[0].confidenceReason).toContain('1 attendee has tentative conflicts');
  });

  it('creates a reply draft without sending mail', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Need input',
          from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
          body: { content: 'Can you review this?' },
          bodyPreview: 'Can you review this?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Need input',
              from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
              body: { content: 'Can you review this?' },
              bodyPreview: 'Can you review this?',
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
          jobTitle: 'Program Manager',
          department: 'Ops',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Need input',
          bodyPreview: 'Thanks for reaching out.',
          webLink: 'https://outlook.example/draft-message',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Need input',
          webLink: 'https://outlook.example/draft-message',
          toRecipients: [{ emailAddress: { name: 'Ops', address: 'ops@example.mil' } }],
          ccRecipients: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Need input',
          webLink: 'https://outlook.example/draft-message',
          toRecipients: [{ emailAddress: { name: 'Ops', address: 'ops@example.mil' } }],
          ccRecipients: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Need input',
          webLink: 'https://outlook.example/draft-message',
          toRecipients: [{ emailAddress: { name: 'Ops', address: 'ops@example.mil' } }],
          ccRecipients: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({}),
      });

    const result = await OutlookService.createReplyDraft(user, 'source-message', {
      instructions: 'Ask for the due date.',
    });

    expect(result).toMatchObject({
      sourceMessageId: 'source-message',
      draftId: 'draft-message',
      message: 'Draft reply created. Review it in Outlook before sending.',
    });
    expect(global.fetch).toHaveBeenNthCalledWith(
      5,
      expect.objectContaining({
        pathname: '/v1.0/me/messages/source-message/createReply',
      }),
      expect.objectContaining({
        method: 'POST',
        body: expect.stringContaining('"comment":""'),
      }),
    );
    expect(global.fetch).toHaveBeenNthCalledWith(
      6,
      expect.objectContaining({
        pathname: '/v1.0/me/messages/draft-message',
      }),
      expect.objectContaining({
        method: 'GET',
      }),
    );
    expect(global.fetch).toHaveBeenLastCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/messages/draft-message',
      }),
      expect.objectContaining({
        method: 'PATCH',
        body: expect.stringContaining('Ask for the due date.'),
      }),
    );
  });

  it('uses model-backed analysis when Outlook AI is configured', async () => {
    OutlookAIService.isModelBackedAIEnabled.mockReturnValue(true);
    OutlookAIService.generateAnalysis.mockResolvedValue({
      insights: {
        mode: 'bedrock',
        summary: 'The sender needs a quick decision.',
        suggestedActions: ['Reply with the decision owner.'],
        riskSignals: ['Decision deadline appears soon.'],
        generatedAt: '2026-04-21T00:00:00.000Z',
      },
      usage: {
        input_tokens: 220,
        output_tokens: 54,
        total_tokens: 274,
        model: 'amazon.nova-micro-v1:0',
        provider: 'bedrock',
      },
    });
    global.fetch.mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        id: 'source-message',
        subject: 'Decision needed',
        from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
        body: { content: 'Can you decide today?' },
        bodyPreview: 'Can you decide today?',
      }),
    });

    const result = await OutlookService.analyzeMessage(user, 'source-message');

    expect(result.insights).toMatchObject({
      mode: 'bedrock',
      summary: 'The sender needs a quick decision.',
    });
    expect(result._usage).toEqual([
      expect.objectContaining({
        context: 'outlook_analyze',
        usage: expect.objectContaining({
          input_tokens: 220,
          output_tokens: 54,
          model: 'amazon.nova-micro-v1:0',
          provider: 'bedrock',
        }),
      }),
    ]);
    expect(OutlookAIService.generateAnalysis).toHaveBeenCalledWith(
      expect.objectContaining({
        message: expect.objectContaining({ subject: 'Decision needed' }),
        outlookContext: expect.objectContaining({
          signedInUser: expect.any(Object),
        }),
      }),
    );
  });

  it('summarizes selected emails into a structured brief', async () => {
    OutlookAIService.isModelBackedAIEnabled.mockReturnValue(true);
    OutlookAIService.generateSelectionBrief.mockResolvedValue({
      brief: {
        mode: 'bedrock',
        headline: 'Two selected emails need follow-up.',
        summary: 'One email needs a decision and another requests scheduling.',
        priorities: ['Reply to Ops with the decision owner.'],
        followUps: ['Confirm a meeting time with Finance.'],
        meetingHighlights: [],
        notableEmails: ['Ops: Decision needed'],
        risks: ['A deadline appears close.'],
        generatedAt: '2026-04-21T00:00:00.000Z',
      },
      usage: {
        input_tokens: 120,
        output_tokens: 45,
        total_tokens: 165,
        model: 'amazon.nova-micro-v1:0',
        provider: 'bedrock',
      },
    });
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'message-1',
          subject: 'Decision needed',
          from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
          body: { contentType: 'text', content: 'Can you confirm the owner today?' },
          bodyPreview: 'Can you confirm the owner today?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'message-2',
          subject: 'Need scheduling',
          from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
          body: { contentType: 'text', content: 'Can we schedule time tomorrow?' },
          bodyPreview: 'Can we schedule time tomorrow?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      });

    const result = await OutlookService.analyzeSelectedMessages(user, ['message-1', 'message-2']);

    expect(result).toMatchObject({
      messageCount: 2,
      brief: expect.objectContaining({
        headline: 'Two selected emails need follow-up.',
      }),
    });
    expect(result._usage).toEqual([
      expect.objectContaining({
        context: 'outlook_selection_summary',
        usage: expect.objectContaining({
          input_tokens: 120,
          output_tokens: 45,
        }),
      }),
    ]);
    expect(OutlookAIService.generateSelectionBrief).toHaveBeenCalledWith(
      expect.objectContaining({
        messages: expect.arrayContaining([
          expect.objectContaining({ subject: 'Decision needed' }),
          expect.objectContaining({ subject: 'Need scheduling' }),
        ]),
      }),
    );
  });

  it('builds a daily brief from the past 24 hours of email and meetings', async () => {
    const recentMessageOne = new Date(Date.now() - 2 * 60 * 60 * 1000).toISOString();
    const recentMessageTwo = new Date(Date.now() - 4 * 60 * 60 * 1000).toISOString();
    const recentMeetingStart = new Date(Date.now() - 3 * 60 * 60 * 1000).toISOString();
    const recentMeetingEnd = new Date(Date.now() - (3 * 60 * 60 - 30 * 60) * 1000).toISOString();

    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'recent-message-1',
              subject: 'Budget follow-up',
              from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
              receivedDateTime: recentMessageOne,
              bodyPreview: 'Please review the updated numbers.',
              importance: 'high',
              isRead: false,
              hasAttachments: false,
            },
            {
              id: 'recent-message-2',
              subject: 'Schedule check-in',
              from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
              receivedDateTime: recentMessageTwo,
              bodyPreview: 'Can we find time this afternoon?',
              importance: 'normal',
              isRead: true,
              hasAttachments: false,
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'meeting-1',
              subject: 'Morning standup',
              start: { dateTime: recentMeetingStart, timeZone: 'UTC' },
              end: { dateTime: recentMeetingEnd, timeZone: 'UTC' },
              organizer: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'recent-message-1',
          subject: 'Budget follow-up',
          from: { emailAddress: { name: 'Finance', address: 'finance@example.mil' } },
          receivedDateTime: recentMessageOne,
          body: { contentType: 'text', content: 'Please review the updated numbers.' },
          bodyPreview: 'Please review the updated numbers.',
          importance: 'high',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'recent-message-2',
          subject: 'Schedule check-in',
          from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
          receivedDateTime: recentMessageTwo,
          body: { contentType: 'text', content: 'Can we find time this afternoon?' },
          bodyPreview: 'Can we find time this afternoon?',
          importance: 'normal',
        }),
      });

    const result = await OutlookService.generateDailyBrief(user, { hours: 24 });

    expect(result.emailCount).toBe(2);
    expect(result.meetingCount).toBe(1);
    expect(result.messageIds).toEqual(['recent-message-1', 'recent-message-2']);
    expect(result.brief).toMatchObject({
      mode: 'local-extractive',
    });
    expect(result.brief.headline).toContain('Daily brief');
  });

  it('builds signed-in user and participant context for Outlook AI prompts', async () => {
    process.env.OUTLOOK_AI_INCLUDE_DIRECTORY_CONTEXT = 'true';
    process.env.OUTLOOK_AI_INCLUDE_MAILBOX_SETTINGS = 'true';
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'User Two',
          mail: 'user2@example.mil',
          userPrincipalName: 'user2@example.mil',
          jobTitle: 'Program Manager',
          department: 'Engineering',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'manager-1',
          displayName: 'Director One',
          mail: 'director@example.mil',
          jobTitle: 'Director',
          department: 'Engineering',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          timeZone: 'Pacific Standard Time',
          workingHours: {
            daysOfWeek: ['monday', 'tuesday'],
            startTime: '09:00:00.0000000',
            endTime: '17:00:00.0000000',
            timeZone: { name: 'Pacific Standard Time' },
          },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'user1',
          displayName: 'User One',
          mail: 'user1@example.mil',
          userPrincipalName: 'user1@example.mil',
          jobTitle: 'VP Finance',
          department: 'Finance',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'user2',
          displayName: 'User Two',
          mail: 'user2@example.mil',
          userPrincipalName: 'user2@example.mil',
          jobTitle: 'Program Manager',
          department: 'Engineering',
        }),
      });

    const context = await OutlookService.getOutlookAIContext(user, {
      from: { name: 'User One', address: 'user1@example.mil' },
      toRecipients: [{ name: 'User Two', address: 'user2@example.mil' }],
      ccRecipients: [],
      body: 'Can you set up time with finance?',
    });

    expect(context.signedInUser).toMatchObject({
      displayName: 'User Two',
      email: 'user2@example.mil',
      jobTitle: 'Program Manager',
    });
    expect(context.manager).toMatchObject({ displayName: 'Director One' });
    expect(context.mailboxSettings).toMatchObject({ timeZone: 'Pacific Standard Time' });
    expect(context.participants).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          address: 'user2@example.mil',
          relationshipToSignedInUser: 'signed_in_user',
          profile: expect.objectContaining({ jobTitle: 'Program Manager' }),
        }),
        expect.objectContaining({
          address: 'user1@example.mil',
          relationshipToSignedInUser: 'internal_user',
          profile: expect.objectContaining({ jobTitle: 'VP Finance' }),
        }),
      ]),
    );
  });

  it('uses model-backed draft text and patches the Outlook draft body', async () => {
    OutlookAIService.isModelBackedAIEnabled.mockReturnValue(true);
    OutlookAIService.generateReplyDraft.mockResolvedValue({
      draft: 'Thanks for the note. I can review this today and follow up with next steps.',
      usage: {
        input_tokens: 180,
        output_tokens: 40,
        total_tokens: 220,
        model: 'amazon.nova-micro-v1:0',
        provider: 'bedrock',
      },
    });
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Need input',
          from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
          body: { content: 'Can you review this?' },
          bodyPreview: 'Can you review this?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Need input',
              from: { emailAddress: { name: 'Ops', address: 'ops@example.mil' } },
              body: { content: 'Can you review this?' },
              bodyPreview: 'Can you review this?',
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
          jobTitle: 'Program Manager',
          department: 'Ops',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Need input',
          webLink: 'https://outlook.example/draft-message',
          toRecipients: [{ emailAddress: { name: 'Ops', address: 'ops@example.mil' } }],
          ccRecipients: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({}),
      });

    const result = await OutlookService.createReplyDraft(user, 'source-message', {
      instructions: 'Be helpful.',
    });

    expect(result.bodyPreview).toContain('I can review this today');
    expect(result.replyMode).toBe('reply');
    expect(result._usage).toEqual([
      expect.objectContaining({
        context: 'outlook_draft',
        usage: expect.objectContaining({
          input_tokens: 180,
          output_tokens: 40,
        }),
      }),
    ]);
    expect(OutlookAIService.generateReplyDraft).toHaveBeenCalledWith(
      expect.objectContaining({
        outlookContext: expect.objectContaining({
          signedInUser: expect.objectContaining({
            displayName: 'Test User',
          }),
        }),
      }),
    );
    expect(global.fetch).toHaveBeenLastCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/messages/draft-message',
      }),
      expect.objectContaining({
        method: 'PATCH',
        body: expect.stringContaining('I can review this today'),
      }),
    );
    expect(OutlookAIService.generateReplyDraft).toHaveBeenCalledWith(
      expect.objectContaining({
        draftRecipients: expect.objectContaining({
          toRecipients: expect.any(Array),
          ccRecipients: expect.any(Array),
        }),
        replyMode: 'reply',
      }),
    );
  });

  it('uses reply-all in smart mode and aligns salutation with resolved recipients', async () => {
    OutlookAIService.isModelBackedAIEnabled.mockReturnValue(true);
    OutlookAIService.generateReplyDraft.mockResolvedValue({
      draft: 'Hi John, Jim and Jenny,\n\nI can take this next step.',
      usage: {
        input_tokens: 160,
        output_tokens: 38,
        total_tokens: 198,
      },
    });

    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Coordination request',
          from: { emailAddress: { name: 'John', address: 'john@example.mil' } },
          toRecipients: [
            { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
            { emailAddress: { name: 'Jim', address: 'jim@example.mil' } },
          ],
          ccRecipients: [{ emailAddress: { name: 'Jenny', address: 'jenny@example.mil' } }],
          body: { content: 'Can everyone align?' },
          bodyPreview: 'Can everyone align?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Coordination request',
              from: { emailAddress: { name: 'John', address: 'john@example.mil' } },
              toRecipients: [
                { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
                { emailAddress: { name: 'Jim', address: 'jim@example.mil' } },
              ],
              ccRecipients: [{ emailAddress: { name: 'Jenny', address: 'jenny@example.mil' } }],
              body: { content: 'Can everyone align?' },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Coordination request',
          webLink: 'https://outlook.example/draft-message',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Coordination request',
          webLink: 'https://outlook.example/draft-message',
          toRecipients: [
            { emailAddress: { name: 'John', address: 'john@example.mil' } },
            { emailAddress: { name: 'Jim', address: 'jim@example.mil' } },
          ],
          ccRecipients: [{ emailAddress: { name: 'Jenny', address: 'jenny@example.mil' } }],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({}),
      });

    const result = await OutlookService.createReplyDraft(user, 'source-message', {});

    expect(result.replyMode).toBe('reply_all');
    expect(global.fetch).toHaveBeenNthCalledWith(
      5,
      expect.objectContaining({
        pathname: '/v1.0/me/messages/source-message/createReplyAll',
      }),
      expect.objectContaining({
        method: 'POST',
      }),
    );
    expect(global.fetch).toHaveBeenLastCalledWith(
      expect.objectContaining({
        pathname: '/v1.0/me/messages/draft-message',
      }),
      expect.objectContaining({
        method: 'PATCH',
        body: expect.stringContaining('Hi John and Jim,'),
      }),
    );
    expect(OutlookAIService.generateReplyDraft).toHaveBeenCalledWith(
      expect.objectContaining({
        replyMode: 'reply_all',
        draftRecipients: expect.objectContaining({
          toRecipients: expect.arrayContaining([
            expect.objectContaining({ address: 'john@example.mil' }),
            expect.objectContaining({ address: 'jim@example.mil' }),
          ]),
        }),
      }),
    );
  });

  it('expands reply-all recipients using thread participants when Graph draft is incomplete', async () => {
    global.fetch
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'source-message',
          conversationId: 'thread-1',
          subject: 'Vendor alignment',
          from: { emailAddress: { name: 'John', address: 'john@example.mil' } },
          toRecipients: [{ emailAddress: { name: 'Test User', address: 'test.user@example.mil' } }],
          body: { content: 'Can we align quickly?' },
          bodyPreview: 'Can we align quickly?',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [
            {
              id: 'source-message',
              conversationId: 'thread-1',
              subject: 'Vendor alignment',
              from: { emailAddress: { name: 'John', address: 'john@example.mil' } },
              toRecipients: [
                { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
              ],
              body: { content: 'Can we align quickly?' },
            },
            {
              id: 'older-message',
              conversationId: 'thread-1',
              subject: 'Vendor alignment',
              from: { emailAddress: { name: 'Test User', address: 'test.user@example.mil' } },
              toRecipients: [
                { emailAddress: { name: 'John', address: 'john@example.mil' } },
                { emailAddress: { name: 'Jim', address: 'jim@example.mil' } },
              ],
              ccRecipients: [{ emailAddress: { name: 'Jenny', address: 'jenny@example.mil' } }],
              body: { content: 'Sharing attendees.' },
            },
          ],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          value: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'me',
          displayName: 'Test User',
          mail: 'test.user@example.mil',
          userPrincipalName: 'test.user@example.mil',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Vendor alignment',
          webLink: 'https://outlook.example/draft-message',
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({
          id: 'draft-message',
          subject: 'RE: Vendor alignment',
          webLink: 'https://outlook.example/draft-message',
          toRecipients: [{ emailAddress: { name: 'John', address: 'john@example.mil' } }],
          ccRecipients: [],
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({}),
      });

    const result = await OutlookService.createReplyDraft(user, 'source-message', {
      replyMode: 'smart',
    });

    expect(result.replyMode).toBe('reply_all');
    expect(result.toRecipients).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'john@example.mil' }),
        expect.objectContaining({ address: 'jim@example.mil' }),
      ]),
    );
    expect(result.ccRecipients).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: 'jenny@example.mil' })]),
    );

    const patchBody = JSON.parse(global.fetch.mock.calls[6][1].body);
    const toAddresses = patchBody.toRecipients.map((recipient) => recipient.emailAddress.address);
    const ccAddresses = patchBody.ccRecipients.map((recipient) => recipient.emailAddress.address);
    expect(toAddresses).toEqual(expect.arrayContaining(['john@example.mil', 'jim@example.mil']));
    expect(ccAddresses).toEqual(expect.arrayContaining(['jenny@example.mil']));
  });
});
