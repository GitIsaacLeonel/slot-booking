# Slot Booking

A booking page with approval, built on Google Apps Script.

Guests pick a free slot, the time is held in your Google Calendar, and you approve or reject by email. After approval the guest gets a calendar invite and a confirmation with a Google Meet, Microsoft Teams or Zoom link.

## Features

- Free slots checked against Google Calendar, extra Google calendars and Outlook (Microsoft 365)
- Pending hold in the calendar until you decide, released automatically after a timeout
- Approve with Google Meet, Teams or Zoom, reject with an optional message
- Guest sees times in their own time zone
- Hidden admin page (secret key) for all settings, connections and a health check
- No server to host: runs free on your Google account

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Server logic (JavaScript, Apps Script V8) |
| `Booking.html` | Public booking page |
| `Approve.html` | Review page for approving or rejecting |
| `Admin.html` | Hidden admin page |
| `appsscript.json` | Manifest (Calendar advanced service, web app settings) |
| `SETUP.md` | Step by step setup guide |

## Setup

See [SETUP.md](SETUP.md).

No credentials are stored in this repository. Microsoft and Zoom secrets live in the Apps Script project properties only.
