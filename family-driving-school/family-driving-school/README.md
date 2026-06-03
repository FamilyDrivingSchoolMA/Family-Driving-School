# Family Driving School — Website

**Family Driving School Inc.** | 102 Main St, Brockton, MA 02301 | (508) 588-4503

---

## Project Structure

```
family-driving-school/
├── index.html        # Main landing page
├── services.html     # Individual services & pricing page
├── admin.html        # Staff admin panel (schedule updates)
├── hero-video.mp4    # Hero background video
├── images/
│   ├── logo.png          # School logo
│   ├── antonio-pires.jpg # Instructor photo
│   └── maria-santos.jpg  # Instructor photo
└── README.md         # This file
```

---

## Pages

### `index.html` — Main Site
The primary landing page. Contains:
- Hero section with video background
- Our Programs — 4 cards (Classroom Training, Behind the Wheel & Observation, Complete Driver's Ed, Private Lessons)
- Why Us section
- Testimonials
- Pricing packages (Teen Driver Ed, Complete Package, Private Lessons)
- FAQ
- Policy section (Behind the Wheel & Road Test policies)
- Team section
- CTA / contact section
- Enrollment modal (multi-step form → Google Sheets)
- Footer with social links (Facebook, LinkedIn)

### `services.html` — Services Page
Linked from the main nav as "Services". Shows all individual à la carte services with pricing:
- Driver's Ed Classroom ($280)
- Parent/Guardian Class (included in classroom payment)
- Private Lessons ($60/hr)
- Behind the Wheel ($50–$60/hr)
- Adult Behind the Wheel ($50/hr)
- Road Test ($120)
- Observation (included in Behind the Wheel)
- Adult Retraining (call for pricing)
- Door-to-Door Pick Up (fee for out-of-range locations)

Also displays the accelerated week schedule (fetched dynamically from Google Sheets config).

### `admin.html` — Admin Panel
Password-protected staff page for updating the accelerated classroom week schedule. See [Admin Panel](#admin-panel) section below.

---

## Our Programs (index.html)

The "Our Programs" section has 4 cards in a 2×2 grid. Students can book any component separately or all together:

| Card | Description |
|------|-------------|
| Classroom Training | 30 required classroom hours. Regular session: 3 Saturdays & 2 Sundays. Accelerated option also available. |
| Behind the Wheel & Observation | 12 hrs supervised driving + 6 hrs backseat observation. Door-to-door pick up available. |
| Complete Driver's Ed | All three components combined (30 + 12 + 6 hrs). |
| Private Lessons | One-on-one sessions for any experience level. |

---

## Accelerated Week Schedule

The accelerated classroom week date is managed dynamically — **no code changes needed** to update it.

- **Current schedule:** April 20–24, 8:00am–2:30pm at 102 Main St, Brockton, MA 02301
- Displayed on both `index.html` (Programs & FAQ sections) and `services.html`
- Updated via the [Admin Panel](#admin-panel)
- Dates are fetched from Google Sheets on page load and cached in `localStorage` for instant display on return visits

---

## Enrollment Form

Multi-step modal form on `index.html`. Submissions go directly to Google Sheets.

| Panel | Fields |
|-------|--------|
| 1 — Personal Info | First name, last name, date of birth, phone, email, street address, apt/unit, city, state, ZIP |
| 2 — Program | Course of interest, preferred start date, schedule preference, learner's permit status, driving experience |
| 3 — Emergency Contact | Emergency contact name, relation & phone; preferred language, how they heard about us, comments |

### Course options (dropdown)
- Complete Driver's Ed — Classroom + BTW + Observation
- Classroom Training Only (30 hrs)
- Behind the Wheel & Observation (12 + 6 hrs)
- Private Lessons
- Road Test

---

## Google Sheets Setup

### What it does
The Google Sheet serves two purposes:
1. **Collects enrollment form submissions** from the website
2. **Stores the accelerated week schedule config** updated via the admin panel

### Sheet tabs
| Tab | Purpose |
|-----|---------|
| Sheet1 (default) | Enrollment form submissions |
| Config | Accelerated week date & location (auto-created on first admin save) |

### Enrollment columns (Sheet1)
| Column | Value |
|--------|-------|
| Timestamp | Auto-generated |
| First Name | |
| Last Name | |
| Date of Birth | |
| Email | |
| Phone | |
| Street | |
| Apt/Unit | |
| City | |
| State | |
| ZIP | |
| Course | |
| Preferred Start | |
| Schedule | |
| Has Permit | |
| Permit # | |
| Experience | |
| Emergency Name | |
| Emergency Relation | |
| Emergency Phone | |
| Preferred Language | |
| How Heard | |
| Comments | |

> **Note:** If adding columns manually, insert Street, Apt/Unit, City, State, ZIP where the old Address column was.

### Apps Script
The Google Apps Script connects the website to the sheet. It handles:
- `doPost` — enrollment form submissions + admin config updates
- `doGet` — returns the current accelerated week schedule as JSON

**To update the Apps Script:**
1. Open your Google Sheet → **Extensions → Apps Script**
2. Replace all code with the script below
3. Click **Deploy → Manage Deployments → edit pencil → New version → Deploy**

```js
function doPost(e) {
  var d = e.parameter;

  if (d.action === 'updateConfig') {
    var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!cfg) {
      cfg = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
      cfg.appendRow(['Key', 'Value']);
    }
    var keys = { accelDate: d.accelDate, accelLocation: d.accelLocation };
    var rows = cfg.getDataRange().getValues();
    for (var k in keys) {
      var found = false;
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === k) { cfg.getRange(i + 1, 2).setValue(keys[k]); found = true; break; }
      }
      if (!found) cfg.appendRow([k, keys[k]]);
    }
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp','First Name','Last Name','Date of Birth',
      'Email','Phone','Street','Apt/Unit','City','State','ZIP',
      'Course','Preferred Start','Schedule',
      'Has Permit','Permit #','Experience',
      'Emergency Name','Emergency Relation','Emergency Phone',
      'Preferred Language','How Heard','Comments'
    ]);
  }
  sheet.appendRow([
    d.timestamp || new Date().toLocaleString(),
    d.firstName, d.lastName, d.dob,
    d.email, d.phone, d.street, d.apt, d.city, d.state, d.zip,
    d.course, d.startDate, d.schedule,
    d.hasPermit, d.permitNumber, d.experience,
    d.emergencyName, d.emergencyRelation, d.emergencyPhone,
    d.language, d.howHeard, d.comments
  ]);
  return ContentService.createTextOutput('OK');
}

function doGet(e) {
  var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var out = {
    accelDate: 'April 20\u201324 \u2022 8:00am\u20132:30pm',
    accelLocation: '102 Main St, Brockton, MA 02301'
  };
  if (cfg) {
    var rows = cfg.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (out.hasOwnProperty(rows[i][0])) out[rows[i][0]] = rows[i][1];
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### Deployment URL
The Apps Script URL is stored in all three files:
- `index.html` — `const SHEETS_URL = '...'`
- `services.html` — `const SHEETS_URL = '...'`
- `admin.html` — `const SHEETS_URL = '...'`

All three must match. If you redeploy, the URL stays the same — no changes needed in the code.

---

## Admin Panel

**URL:** `yoursite.com/admin.html`

**Password:** `FamilyDS2025`
> Change this by editing the `ADMIN_PASSWORD` line in `admin.html`

### What it does
Lets staff update the **accelerated classroom week** dates without touching any code.

### How to use
1. Go to `admin.html` and log in
2. Select the start and end dates for the next accelerated week
3. Set the start and end times
4. Confirm or update the location
5. Click **Save Changes**
6. The website updates automatically within seconds

### How it works
- Admin panel POSTs the new dates to the Apps Script
- Apps Script writes them to the **Config** tab in Google Sheets
- Website fetches config on every page load and updates the displayed dates
- Dates are cached in `localStorage` so they show instantly on return visits — no blank flash

---

## Policy Section

Displayed on `index.html` between the Team and CTA sections. Two policy cards:

**Behind the Wheel**
- Pay per session or upfront — no payment after the fact
- 24-hour cancellation required or $30 fee applies
- Pick-up location must be confirmed to hold time slot
- $3 processing fee for third-party payment apps

**Road Test**
- All outstanding payments must be settled before scheduling
- 72-hour cancellation required
- No-shows forfeit all road test fees (paid directly to RMV)

---

## Social Links

| Platform | URL |
|----------|-----|
| Facebook | https://www.facebook.com/profile.php?id=61581270393447 |
| LinkedIn | https://www.linkedin.com/in/antonio-pires-1643a226b/ |

Icons appear in the footer of `index.html` with hover color effects.

---

## Mobile Responsiveness

Both `index.html` and `services.html` are fully responsive:

| Breakpoint | Changes |
|------------|---------|
| 1024px | Pricing stacks to single column |
| 900px | Services grid goes 2-column |
| 768px | Nav collapses to hamburger, hero stacks, footer stacks |
| 560px | Services grid goes single column |
| 480px | CTA buttons stack vertically, stat numbers scale down |

---

## Hosting & Domain

The site is plain HTML — no server or build process required.

**Recommended: GitHub Pages (free)**
1. Go to your GitHub repo → **Settings → Pages**
2. Set branch to **main** → **/ (root)** → Save
3. Site is live at `jayexs.github.io/family-driving-school`

**Using your existing Wix domain with GitHub Pages**
1. Enable GitHub Pages (above)
2. In Wix domain settings → **DNS settings**
3. Point DNS records to GitHub Pages
4. GitHub Pages serves the site under your Wix-owned domain

**Costs**
- GitHub Pages hosting — free
- Wix domain — already owned
- Google Workspace — already owned
- Stripe (when payments are added) — no monthly fee, 2.9% + 30¢ per transaction

---

## Transferring to Business Email

1. **Google Sheet** — Share with business email → File → Make a copy under business account
2. **Apps Script** — Open new sheet → Extensions → Apps Script → paste script → redeploy
3. **Update SHEETS_URL** — If URL changes, update in `index.html`, `services.html`, and `admin.html`
4. **Upload files** — Upload all HTML/image/video files to hosting
5. **Admin password** — Update `ADMIN_PASSWORD` in `admin.html`

---

## Planned Features

- **Online payments** — Stripe Payment Links per service
- **Scheduling** — Google Calendar Appointment Schedules (included with Google Workspace)
- **Post-payment booking** — After Stripe payment, student is redirected to a booking page with their Google Calendar appointment link

---

*Last updated: March 2026*
