import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken } from './msal';
import { Timeblock, GraphCalendarEvent } from '@/types';
import { formatISO, startOfDay, endOfDay, addDays } from 'date-fns';

let graphClient: Client | null = null;

/**
 * Initialize Microsoft Graph client with access token
 */
const getGraphClient = async (): Promise<Client> => {
  if (graphClient) return graphClient;

  const token = await getAccessToken();

  graphClient = Client.init({
    authProvider: async (done) => {
      done(null, token);
    },
  });

  return graphClient;
};

/**
 * Get user's mailbox settings (timezone, etc)
 */
export const getUserMailboxSettings = async () => {
  try {
    const client = await getGraphClient();
    const settings = await client.api('/me/mailboxSettings').get();
    return {
      timeZone: settings.timeZone || 'UTC',
      language: settings.language?.locale || 'en-US',
    };
  } catch (error) {
    console.error('Failed to fetch mailbox settings:', error);
    return { timeZone: 'UTC', language: 'en-US' };
  }
};

/**
 * Fetch calendar events from Microsoft Graph
 * Uses /calendarView to automatically expand recurring events
 */
export const getCalendarEvents = async (
  startDateTime: string,
  endDateTime: string
): Promise<GraphCalendarEvent[]> => {
  try {
    const client = await getGraphClient();

    const response = await client
      .api('/me/calendar/calendarView')
      .query({
        startDateTime,
        endDateTime,
        $top: 250, // Max results per page
        $select: [
          'id',
          'subject',
          'start',
          'end',
          'isAllDay',
          'location',
          'organizer',
          'webLink',
          'bodyPreview',
        ].join(','),
        $orderby: 'start/dateTime asc',
      })
      .get();

    return response.value || [];
  } catch (error) {
    console.error('Failed to fetch calendar events:', error);
    return [];
  }
};

/**
 * Fetch all calendar events with pagination support
 */
export const getAllCalendarEvents = async (
  startDateTime: string,
  endDateTime: string
): Promise<GraphCalendarEvent[]> => {
  let allEvents: GraphCalendarEvent[] = [];
  let skipCount = 0;
  const pageSize = 250;

  try {
    const client = await getGraphClient();

    while (true) {
      const response = await client
        .api('/me/calendar/calendarView')
        .query({
          startDateTime,
          endDateTime,
          $top: pageSize,
          $skip: skipCount,
          $select: [
            'id',
            'subject',
            'start',
            'end',
            'isAllDay',
            'location',
            'organizer',
            'webLink',
          ].join(','),
          $orderby: 'start/dateTime asc',
        })
        .get();

      const events = response.value || [];
      allEvents = [...allEvents, ...events];

      if (events.length < pageSize) break;
      skipCount += pageSize;
    }

    return allEvents;
  } catch (error) {
    console.error('Failed to fetch all calendar events:', error);
    return allEvents;
  }
};

/**
 * Transform Microsoft Graph event to Timeblock format
 */
export const transformGraphEventToTimeblock = (
  event: GraphCalendarEvent,
  userId: string
): Timeblock => {
  const startTime = event.start.dateTime;
  const endTime = event.end.dateTime;
  const date = startTime.split('T')[0]; // Extract date part

  return {
    id: crypto.randomUUID(),
    user_id: userId,
    outlook_event_id: event.id,
    subject: event.subject,
    start_time: startTime,
    end_time: endTime,
    is_all_day: event.isAllDay,
    location: event.location?.displayName || null,
    organizer: event.organizer?.emailAddress?.name || null,
    source_link: event.webLink,
    date,
    last_synced_at: new Date().toISOString(),
  };
};

/**
 * Build ISO datetime strings for calendarView query
 */
export const buildCalendarViewDateRange = (
  startDate: Date = new Date(),
  daysAhead: number = 7
) => {
  const start = startOfDay(startDate);
  const end = endOfDay(addDays(startDate, daysAhead - 1));

  return {
    startDateTime: formatISO(start),
    endDateTime: formatISO(end),
  };
};
