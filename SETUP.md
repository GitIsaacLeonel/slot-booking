# Booking with approval: setup guide

## How it works

1. A guest opens your booking page and sees free slots (in their own time zone). A slot is only shown when you are free in Google Calendar AND in Outlook, with a 10 minute buffer.
2. The guest picks a slot and enters name, email and topic. A yellow **[PENDING]** hold appears in your Google Calendar so nobody else can take the slot. The guest gets a "request received" email.
3. You get an email with the buttons **Approve with Google Meet**, **Approve with Teams**, **Approve with Zoom** and **Reject** (Teams and Zoom only appear once they are connected). Each opens a review page where you confirm with one click. (Link previews in Outlook or Gmail cannot trigger anything by accident.)
4. On approval:
   - **Meet**: the hold becomes the real event, a Meet link is created and Google sends the invite. The time is also blocked in Outlook.
   - **Teams**: an Outlook event with a Teams link is created and Outlook sends the invite. The Google hold stays as a block with the Teams link.
   - **Zoom**: a Zoom meeting is created, the hold becomes the real event with the Zoom link and Google sends the invite. The time is also blocked in Outlook.
   - The guest also gets a confirmation email with a "Join" button.
5. On rejection the hold is removed and the guest gets a polite email (with your optional message) and a link to pick another time.
6. Unanswered requests are released automatically after 48 hours, or 1 hour before the slot starts.

## Part 1: Google (about 10 minutes)

1. Go to **script.google.com** with the Google account whose calendar should be used. Click **New project** and name it "Booking".
2. Click the gear icon (**Project Settings**) and tick **Show "appsscript.json" manifest file in editor**. Also check the time zone is Europe/Berlin.
3. Back in the **Editor**:
   - Replace the content of `appsscript.json` with the file from this package.
   - Replace the content of `Code.gs` with the file from this package.
   - Click **+ > HTML**, name it `Booking` (without .html) and paste `Booking.html`.
   - Click **+ > HTML**, name it `Approve` and paste `Approve.html`.
   - Click **+ > HTML**, name it `Admin` and paste `Admin.html`.
4. Nothing to edit in the code. All settings are changed later on the admin page (step 9). The `DEFAULTS` block at the top of `Code.gs` is only the starting point.
5. Select the function **testSetup** in the toolbar and click **Run**. Google asks for permissions (calendar, email, external requests). Click **Advanced > Go to Booking (unsafe)**. This warning is normal for your own scripts. Check the log shows a number of free slots.
6. Click **Deploy > New deployment**, type **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy** and copy the **Web app URL**. This is your booking link.
7. In **Project Settings > Script properties** add `WEB_APP_URL` with that URL (makes sure email links always point to the live version).
8. Select **installTriggers** and click **Run** once. This sets up the hourly cleanup.
9. Select **setupAdmin** and click **Run**. The log shows your private admin link. Bookmark it and keep it to yourself.
10. Open the admin link, check your settings, click **Run check** at the bottom, then open the booking link in a private browser window and book a test slot with another email address.

**Updating later:** after changing code, use **Deploy > Manage deployments > Edit (pencil) > Version: New version**. This keeps the same URL. Your saved settings stay as they are.

## The admin page

Your private link looks like `https://script.google.com/macros/s/.../exec?action=admin&key=...`. Without the right key the normal booking page is shown, so nobody can tell an admin page exists.

What you can do there:

- **Pause bookings** with one switch, and set the message guests see while paused
- **Calendars:** pick the main Google calendar, switch the busy check for Google and Outlook on or off, add other Google calendars (family, team) whose busy times should count, and choose whether Meet and Zoom meetings are also blocked in Outlook
- **Meeting platforms:** switch Meet, Teams and Zoom on or off as approval options
- **Availability:** working hours per weekday, meeting length, buffer, minimum notice, days ahead, time zone
- **Requests:** expiry time, open requests per guest, extra confirmation email
- **Open and upcoming meetings** with a Review link for pending requests and a Join link for confirmed ones
- **Connections:** enter or replace the Microsoft 365 and Zoom credentials (secrets are never shown again), or disconnect them
- **Health check:** tests Google, Outlook, Zoom and shows how many slots guests currently see

Pending and confirmed requests always block their slot, even if you switch a calendar check off, so you cannot be double booked by your own requests.

**If the admin link ever leaks:** run **rotateAdminKey** in the editor. The old link stops working right away and the log shows the new one.

## Part 2: Outlook and Teams (about 15 minutes)

Requirement: a **Microsoft 365 work or school account** (Teams via API does not work with personal outlook.com accounts) and admin rights in that tenant, or an admin who can give consent.

1. Go to **entra.microsoft.com > Applications > App registrations > New registration**.
   - Name: "Booking App"
   - Supported account types: **Single tenant**
   - No redirect URI. Click **Register**.
2. On the overview page copy **Application (client) ID** and **Directory (tenant) ID**.
3. **Certificates & secrets > New client secret**, pick 24 months, and copy the **Value** right away (it is shown only once).
4. **API permissions > Add a permission > Microsoft Graph > Application permissions**, search **Calendars.ReadWrite**, add it, then click **Grant admin consent**.
5. Open your admin page, go to **Connections > Microsoft 365** and enter Tenant ID, Client ID, the secret Value and your Outlook address (e.g. you@company.com). Click **Save connections**.
6. Click **Run check**. Outlook should show OK. Then switch on the Outlook and Teams options you want under Calendars and Meeting platforms.

(Alternative without the admin page: add `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET` and `MS_USER` under **Project Settings > Script properties**.)

**Recommended security step:** application permissions allow access to every mailbox in the tenant by default. Restrict the app to your own mailbox in Exchange Online with RBAC for Applications (or the older Application Access Policy). Your IT admin will know this, it takes one PowerShell command.

**Remember:** the client secret expires. Put a reminder in your calendar to renew it before the end date and update `MS_CLIENT_SECRET`.

## Part 3: Zoom (about 10 minutes, optional)

Works with free and paid Zoom accounts. Note that free accounts limit meetings to 40 minutes. You need to be the account owner or admin to create the app.

1. Go to **marketplace.zoom.us**, sign in, then **Develop > Build App** and choose **Server-to-Server OAuth App**. Name it "Booking App".
2. On the **App Credentials** page copy **Account ID**, **Client ID** and **Client Secret**.
3. **Information**: fill in the required fields (company name, your name and email).
4. **Scopes > Add Scopes**: under Meeting add **meeting:write:meeting:admin** (create a meeting for a user). On older accounts it is called **meeting:write:admin**.
5. **Activation**: click **Activate your app**.
6. Open your admin page, go to **Connections > Zoom** and enter Account ID, Client ID, Client Secret and the email of your Zoom user (the host of the meetings). Click **Save connections**.
7. Click **Run check**. Zoom should show OK.

(Alternative without the admin page: add `ZOOM_ACCOUNT_ID`, `ZOOM_CLIENT_ID`, `ZOOM_CLIENT_SECRET` and `ZOOM_USER` under **Project Settings > Script properties**.)

When you approve with Zoom, a Zoom meeting is created (waiting room on, guests cannot join before you), the calendar hold becomes the real Google event with the Zoom link, and Google sends the invite to the guest. If Outlook is connected, the time is also blocked there.

## Good to know

- **Without Part 2 or 3** everything works with Google Calendar and Meet only. The Teams and Zoom buttons simply do not appear until you connect them.
- **Fail safe:** if Outlook cannot be reached, the booking page shows "Could not load times" instead of risking a double booking.
- **Emails** go out from your Gmail. A free Gmail account can send about 100 script emails per day, which is plenty for bookings.
- **Spam protection:** a hidden field catches bots and one email address can have at most 2 open requests.
- **Overview of all requests:** run **listRequests** in the editor and look at the log.
- **Embedding:** you can link the booking URL from your website, email signature or LinkedIn profile.
- **Google Workspace accounts:** if "Anyone" is not available under access, your Workspace admin has blocked public web apps. A personal Gmail account works without restriction.
