// Types for Timeblock data
export interface Timeblock {
  id: string;
  user_id: string;
  outlook_event_id: string;
  subject: string;
  start_time: string; // ISO 8601
  end_time: string; // ISO 8601
  is_all_day: boolean;
  location: string | null;
  organizer: string | null;
  source_link: string | null;
  date: string; // YYYY-MM-DD
  last_synced_at: string; // ISO 8601 timestamp
  created_at?: string;
  updated_at?: string;
}

// Type for Microsoft Graph calendar event
export interface GraphCalendarEvent {
  id: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  isAllDay: boolean;
  location?: {
    displayName?: string;
  };
  organizer?: {
    emailAddress?: {
      name?: string;
    };
  };
  webLink: string;
}

// Type for sync result
export interface SyncResult {
  inserted: number;
  updated: number;
  errors: Array<{
    event_id: string;
    error: string;
  }>;
  last_synced_at: string;
}

// Type for Task created from timeblock
export interface TaskFromTimeblock {
  id?: string;
  title: string;
  date: string;
  source: 'Calendar';
  source_link: string;
  status: 'Open';
  user_id?: string;
  created_at?: string;
}

// Type for authentication state
export interface AuthState {
  isAuthenticated: boolean;
  user: {
    id: string;
    displayName: string;
    userPrincipalName: string;
  } | null;
  accessToken: string | null;
}

// Type for planner state
export interface PlannerState {
  timeblocks: Timeblock[];
  loading: boolean;
  lastSyncedAt: string | null;
  error: string | null;
}
