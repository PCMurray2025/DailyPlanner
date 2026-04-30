# Daily Planner

A web-based calendar synchronization tool that automatically pulls Outlook calendar events and displays them as timeblocks in a daily planner interface.

## Features

- **Microsoft Outlook Integration**: Automatically sync calendar events for today + next 7 days
- **Timeblock Visualization**: View events as visual timeblocks grouped by day
- **Task Conversion**: Convert calendar events to tasks with one click
- **Read-Only Safety**: Never modifies your Outlook calendar
- **Automatic Sync**: Re-syncs on app open
- **Timezone Support**: Handles your local timezone correctly

## Tech Stack

- **Frontend**: Next.js 14 (React 18)
- **Authentication**: Microsoft Authentication Library (MSAL)
- **Calendar API**: Microsoft Graph API
- **Database**: Supabase
- **State Management**: Zustand

## Getting Started

### Prerequisites

- Node.js 18+
- Microsoft 365 account
- Supabase project

### Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Set up environment variables (see `.env.local.example`)

4. Run development server:
   ```bash
   npm run dev
   ```

5. Open http://localhost:3000

## Environment Variables

```
NEXT_PUBLIC_MSAL_CLIENT_ID=<your_client_id>
NEXT_PUBLIC_MSAL_AUTHORITY=https://login.microsoftonline.com/common
NEXT_PUBLIC_SUPABASE_URL=<your_supabase_url>
NEXT_PUBLIC_SUPABASE_ANON_KEY=<your_anon_key>
SUPABASE_SERVICE_ROLE_KEY=<your_service_key>
```

## Database Schema

The `timeblocks` table stores synced Outlook events:

- `id` - UUID primary key
- `user_id` - Supabase user reference
- `outlook_event_id` - Unique identifier from Outlook (used for upsert)
- `subject` - Event title
- `start_time` - Event start (ISO 8601 datetime)
- `end_time` - Event end (ISO 8601 datetime)
- `is_all_day` - Boolean flag for all-day events
- `location` - Event location
- `organizer` - Event organizer name
- `source_link` - Deep link to event in Outlook
- `date` - Date only (derived from start_time)
- `last_synced_at` - Timestamp of last sync

## Usage

### Viewing Your Schedule

1. Log in with Microsoft account
2. Calendar syncs automatically
3. View events in Planner tab (grouped by day)
4. See all upcoming events in Upcoming tab

### Converting Events to Tasks

Click "Convert to Task" on any timeblock to:
- Create a new task with the event title
- Automatically link back to the calendar event
- Set the task date to match the event date

### Manual Sync

Click "Sync Calendar Now" button to refresh events manually.

## Architecture

```
src/
├── app/                      # Next.js app directory
│   ├── layout.tsx
│   ├── page.tsx
│   └── api/                  # API routes
│       ├── auth/
│       └── calendar/
├── components/               # React components
│   ├── Auth/
│   ├── PlannerTabs/
│   ├── Timeblock/
│   └── TaskConversion/
├── lib/                      # Utilities
│   ├── msal.ts              # MSAL configuration
│   ├── graph.ts             # Microsoft Graph client
│   ├── supabase.ts          # Supabase client
│   ├── calendar.ts          # Calendar sync logic
│   └── tz.ts                # Timezone utilities
├── hooks/                    # Custom React hooks
├── store/                    # Zustand state management
└── types/                    # TypeScript types
```

## Security

- ✅ Read-only access to Outlook calendar
- ✅ No delete/edit operations on calendar
- ✅ MSAL handles OAuth 2.0 securely
- ✅ Calendars.ReadBasic permission scope only
- ✅ User data isolated by user_id
