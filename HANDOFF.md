# SLS Robuster — Handoff Guide

**Who this is for:** anyone who needs to install, use, or hand off this Mac app — even if you have never seen it before and are not a programmer.

**What this app does in one sentence:** SLS Robuster helps staff set live YouTube stream titles and thumbnails (and optionally update the school website) during School of Leadership Studies class broadcasts.

---

## Table of contents

1. [What is SLS Robuster?](#1-what-is-sls-robuster)
2. [Where to find it on GitHub](#2-where-to-find-it-on-github)
3. [Install on a Mac (step by step)](#3-install-on-a-mac-step-by-step)
4. [First-time setup after install](#4-first-time-setup-after-install)
5. [How the app works (tour by tour)](#5-how-the-app-works-tour-by-tour)
6. [Typical day-of-class workflow](#6-typical-day-of-class-workflow)
7. [Where everything is saved](#7-where-everything-is-saved)
8. [Getting newer versions](#8-getting-newer-versions)
9. [Common problems](#9-common-problems)
10. [For IT / developers (short)](#10-for-it--developers-short)

---

## 1. What is SLS Robuster?

SLS Robuster is a **desktop app for Mac**. You open it like any other app (Safari, Zoom, etc.). You do **not** need to install Python or write code to use it day to day.

It connects to:

| Service | What Robuster uses it for |
|---------|---------------------------|
| **YouTube** | Rename the live stream, upload a thumbnail, put the video in an archive playlist, show “is it live?” status |
| **DreamClass** (optional) | Show today’s class schedule so you can pick titles/instructors quickly |
| **WordPress + SLS Live plugin** (optional) | Push the live video ID to the school website when you hit Apply |

The main job on broadcast day:

1. Pick class title, instructor, and date.
2. Preview the title and thumbnail.
3. Click **Apply to live stream**.
4. Robuster updates YouTube (and the website, if that is turned on).

---

## 2. Where to find it on GitHub

There are **two** GitHub places. They do different jobs.

### A. Download the app (this is what most people need)

**Public updates repo:**  
[https://github.com/mrcodeimator/SLS-Robuster-Updates](https://github.com/mrcodeimator/SLS-Robuster-Updates)

- Anyone with the link can open it (no special GitHub login required for downloads).
- **Releases** page holds the Mac zip files (the actual installers).
- Current release (as of this guide): **v2.2.0**  
  Direct Releases page:  
  [https://github.com/mrcodeimator/SLS-Robuster-Updates/releases](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases)
- Example download for Apple Silicon Macs:  
  [SLS_ROBUSTER-macos-arm64.zip (v2.2.0)](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases/download/v2.2.0/SLS_ROBUSTER-macos-arm64.zip)

### B. Source code (for developers / IT only)

**Private source repo:**  
[https://github.com/mrcodeimator/SLS-ROBUSTER-V2](https://github.com/mrcodeimator/SLS-ROBUSTER-V2)

- This is the programming project (code).
- It is **private** — you need GitHub access from the organization account that owns it.
- Day-to-day users should **not** need this. Install from the public Updates repo above.

| If you want… | Go here |
|--------------|---------|
| Install or update the Mac app | Public: **SLS-Robuster-Updates** → Releases |
| Change how the program is written | Private: **SLS-ROBUSTER-V2** |

---

## 3. Install on a Mac (step by step)

You need a **Mac**. The packaged app is built for macOS (current public zip is for Apple Silicon / `arm64`).

### Step 1 — Download

1. Open this page in a browser:  
   [https://github.com/mrcodeimator/SLS-Robuster-Updates/releases](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases)
2. Open the newest release (for example **SLS Robuster 2.2.0**).
3. Under **Assets**, download the `.zip` file (name looks like `SLS_ROBUSTER-macos-arm64.zip`).

### Step 2 — Unzip

1. Find the zip in your **Downloads** folder.
2. Double-click it. macOS will create a folder or an app named **`SLS_ROBUSTER.app`**.

### Step 3 — Put it in Applications

1. Open **Finder**.
2. Open **Applications**.
3. Drag **`SLS_ROBUSTER.app`** into Applications.  
   (You can leave it in Downloads, but Applications is the usual place.)

### Step 4 — Open it the first time (Gatekeeper)

The app is not Apple-notarized yet, so macOS may block a normal double-click the first time.

1. In Finder, go to **Applications**.
2. **Right-click** (or Control-click) **`SLS_ROBUSTER`**.
3. Choose **Open**.
4. In the warning dialog, click **Open** again.

After that, you can open it normally from Launchpad, Spotlight, or Applications.

### Step 5 — Confirm it launched

You should see a dark window titled **SLS Robuster**, with a left sidebar and tabs such as **Retitle**, **Schedule**, and **Stream Info**.

---

## 4. First-time setup after install

Do these once per Mac (or after a wipe). Settings are remembered for that user account.

### 4.1 Sign in each YouTube channel

Robuster needs permission to change each channel’s live stream.

1. Click **Settings** (gear / Settings in the sidebar).
2. Open the **Channels** tab.
3. Click a channel chip (for example **SLS Year 1**).
4. Click **Sign In / Re-auth**.
5. A browser window opens — sign in with the Google account that **owns or manages** that YouTube channel.
6. Approve access when Google asks.
7. Repeat for every channel you will use (Year 1, Year 2, Intern, etc.).

A **✓** on the channel means login is saved. A **✗** means you still need to sign in.

### 4.2 Check channel details (if someone asks)

For each channel you can set:

- **Display name** — label in the app  
- **Channel ID** — YouTube channel ID (starts with `UC…`)  
- **Playlist** — archive playlist (dropdown after you are signed in)  
- **WordPress profile** — e.g. `year1` / `year2` if the website should update for that channel  
- **Thumbnail template** — optional custom background image path  

Click **Save Channel** when done.

### 4.3 DreamClass (class schedule) — optional but useful

Used by the **Schedule** tab.

1. Settings → **DreamClass** tab.
2. Enter the API key, school code, and tenant your organization uses.
3. Save.

If this is empty, Schedule will not load live class data; Retitle still works with manual dropdowns.

### 4.4 WordPress / school website — optional

Used so **Apply** can update the live embed on the website (SLS Live plugin).

1. Settings → **WordPress** tab.
2. Turn on **Enable WordPress push**.
3. Enter the site URL (example: `https://yoursite.com` — no trailing slash).
4. Enter the **secret key** that matches the WordPress plugin settings.
5. Click **Test Connection**, then **Save WordPress Settings**.

Each channel that should update the site needs a **WordPress profile** filled in under Channels.

### 4.5 General preferences (optional)

Settings → **General**:

- **Preview Mode** — practice without changing YouTube (see below)  
- **Start in Preview Mode**  
- **Launch at Login** — open Robuster when you log into the Mac  
- **Keep window on top**  
- **Check for Updates** — look for a newer app version  

---

## 5. How the app works (tour by tour)

### Left sidebar

| Control | Meaning |
|---------|---------|
| **Hamburger (☰)** | Expand / collapse the sidebar |
| **Live Status** | Shows each channel you chose to watch; Live / not live |
| Checkboxes next to channels | Which channels to monitor for “Check if Live” |
| **Check if Live** | Repeatedly checks YouTube until monitored channels are live (then stops) |
| **Preview** | When on, Apply only *logs* what it would do — no real YouTube/WordPress changes |
| **Settings** | Channels, presets, DreamClass, WordPress, General |

### Top navigation

Main work areas:

1. **Retitle** — compose title + thumbnail and apply to live streams  
2. **Schedule** — today’s (or any day’s) classes from DreamClass  
3. **Stream Info** — scan for live broadcasts, copy links / embed codes  

An **Activity Log** at the bottom shows what Robuster just did (successes and errors). You can collapse it.

### Retitle (main daily screen)

**Title composer**

- Choose **class title**, **instructor**, and **date** from dropdowns (lists come from Presets).
- Robuster builds a title like:  
  `Class Title | Instructor | Date`
- The **thumbnail preview** updates as you change fields.

**Apply to**

- Check which YouTube channels should receive the change (Year 1, Year 2, etc.).

**Apply to live stream** — for each selected channel, Robuster tries to:

1. Find the **active or testing** live broadcast on that channel.  
2. **Rename** it to the composed title.  
3. **Add** the video to that channel’s archive **playlist** (if a playlist is set).  
4. **Upload** the generated thumbnail.  
5. If WordPress is enabled and the channel has a profile: **push** the video ID/title to the website.

If Preview is on, none of those changes happen — you only see messages in the Activity Log starting with `[Preview]`.

### Schedule

- Shows a timeline (roughly 8 AM–5 PM) with classes from DreamClass.
- **Family Center** and **Chapel** appear in columns.
- Click the date to open a calendar and jump to another day.
- You can send a class into Retitle (pre-fill title/instructor) when that action is available on a class chip.

### Stream Info

- **Scan** channels for an active broadcast.
- Shows **Video ID**, **Watch URL**, **Embed code**, **Title**, **Started** time.
- **Copy** buttons for quick paste into other tools.
- **Push to WordPress** on a card if that channel has a WordPress profile.
- **Recover Missed Streams** helps with streams that were missed earlier (secondary recovery tool).

### Settings → Presets

Editable lists of:

- Class titles  
- Instructors  
- Dates  

These feed the Retitle dropdowns. Add or remove items as the semester changes.

### Settings → Channels

Add / edit / remove channels, Sign In, playlists, WordPress profile, “Show in Monitor.”

---

## 6. Typical day-of-class workflow

Use this as a checklist for operators.

1. Open **SLS_ROBUSTER** from Applications.  
2. Confirm **Preview** is **off** when you want real updates (on only for practice).  
3. In the sidebar, check the channels you care about; optionally click **Check if Live**.  
4. Open **Retitle**.  
5. Pick class title, instructor, and date (or prefill from **Schedule**).  
6. Confirm the thumbnail preview looks right.  
7. Under **Apply to**, check Year 1 / Year 2 / etc. as needed.  
8. Click **Apply to live stream**.  
9. Watch the **Activity Log** for “Retitled…”, “Archived…”, thumbnail, and WordPress messages.  
10. Optionally open **Stream Info** → Scan to copy the watch URL or embed for someone else.

If something fails with “No token” → Settings → Channels → **Sign In / Re-auth** for that channel.

---

## 7. Where everything is saved

### The app itself

| What | Usual location |
|------|----------------|
| The program | `/Applications/SLS_ROBUSTER.app` (or wherever you dragged it) |

Updating or replacing this app **does not** erase your settings or YouTube logins.

### Your settings and logins (important)

All personal/org data for the Mac user lives here:

```text
~/Library/Application Support/SLSRobuster/
```

In Finder:

1. Open Finder.  
2. Menu **Go** → **Go to Folder…**  
3. Paste: `~/Library/Application Support/SLSRobuster`  
4. Press Return.

| File / item (examples) | What it is |
|------------------------|------------|
| `channels.json` | Your channel list (names, IDs, playlists, WordPress profiles, monitor flags) |
| `retitle_presets.json` | Class titles, instructors, dates used in dropdowns |
| `general_settings.json` | Preview-on-start, stay on top, launch at login preference, etc. |
| `dreamclass_settings.json` | DreamClass API key / school code / tenant |
| `wordpress_settings.json` | WordPress site URL, secret key, enable flag |
| `app_config.json` | App wiring (update URL, default client secret name, etc.) |
| `client_secret….json` | Google OAuth client file (copied from the app on first run) |
| `token-UC….pickle` | **YouTube login** for each channel (one file per channel ID) |
| `thumbnail_layout.json` | Where text sits on the thumbnail |
| `thumbnail_preview.jpg` / `thumbnail_apply.jpg` | Recent thumbnail images Robuster generated |

**Treat the `token-*.pickle` files and secret keys as sensitive.** Do not email them or post them publicly. If you copy this folder to a new Mac, you may move settings — but YouTube may still ask you to sign in again.

### Launch at Login (if enabled)

| What | Location |
|------|----------|
| Login helper | `~/Library/LaunchAgents/com.sls.robuster.plist` |
| Log file | `~/Library/Logs/SLSRobuster.log` |

### What is *not* in that folder

- The **source code** (only on GitHub / a developer machine).  
- The **public zip** (only on the Updates GitHub Releases page until someone downloads it).

---

## 8. Getting newer versions

### For normal users (easiest)

If you already have **v2.2.0** or later as a real `.app`:

1. Open Robuster.  
2. Settings → **General** → **Check for Updates**,  
   **or** use the update banner if it appears after launch.  
3. If a newer version exists, confirm download/install and restart when asked.  
4. macOS may ask for your password if the app lives in **Applications** — that is normal.

Your settings folder (`Application Support/SLSRobuster`) stays put.

### First install of 2.2.0 (or any brand-new Mac)

Still: download zip from GitHub Releases → unzip → drag to Applications (section 3).

### How updates are published (plain English)

```text
Developers change code in private repo (SLS-ROBUSTER-V2)
        ↓
Someone builds a new Mac zip
        ↓
Zip is uploaded to public repo Releases (SLS-Robuster-Updates)
        ↓
version.json on that public repo is updated
        ↓
Installed apps see “update available” and can download the zip
```

A **pull request / merge on GitHub is not the same as a public update.** Until the zip and `version.json` are published on **SLS-Robuster-Updates**, people’s Macs will not get the new build.

More detail for IT: [UPDATES.md](UPDATES.md) and [BUILD.md](BUILD.md).

---

## 9. Common problems

| What you see | What to try |
|--------------|-------------|
| macOS says the app can’t be opened | Right-click app → **Open** → **Open** again |
| “No token” / auth errors | Settings → Channels → **Sign In / Re-auth** |
| Apply does nothing useful | Turn **Preview** off; confirm a stream is live or in testing on YouTube Studio |
| Wrong channel updated | Check **Apply to** boxes on Retitle |
| Website didn’t update | Settings → WordPress enabled? Channel has WordPress profile? Test Connection |
| Schedule empty | DreamClass settings filled in? Internet working? |
| “You’re on the latest” but IT just shipped | Public `version.json` may not be updated yet — ask whoever publishes releases |
| Update download fails / 404 | Zip must be on **public** SLS-Robuster-Updates, not only the private code repo |
| Need a clean start | Quit app; optionally delete `~/Library/Application Support/SLSRobuster/` (you will lose logins and settings) |

---

## 10. For IT / developers (short)

| Topic | Document |
|-------|----------|
| Build the `.app` and zip on a Mac | [BUILD.md](BUILD.md) |
| How in-app updates and the two repos work | [UPDATES.md](UPDATES.md) |
| Version history (user-facing notes) | `CHANGELOG.txt` in the project root |
| Public install / updates | [mrcodeimator/SLS-Robuster-Updates](https://github.com/mrcodeimator/SLS-Robuster-Updates) |
| Private source | [mrcodeimator/SLS-ROBUSTER-V2](https://github.com/mrcodeimator/SLS-ROBUSTER-V2) |

**Important packaging note:** Google OAuth `client_secret*.json` is required on the build machine and is bundled into the release zip so end users can Sign In. That file must **not** be committed to git. User tokens (`token-*.pickle`) are created only on each person’s Mac after Sign In.

App bundle id: `com.sls.robuster`.

---

## Quick reference card

| Need | Answer |
|------|--------|
| Download app | [SLS-Robuster-Updates Releases](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases) |
| Source code | [SLS-ROBUSTER-V2](https://github.com/mrcodeimator/SLS-ROBUSTER-V2) (private) |
| Install | Unzip → drag `SLS_ROBUSTER.app` to Applications → Right-click → Open |
| Settings / logins live in | `~/Library/Application Support/SLSRobuster/` |
| Main daily action | Retitle → compose → Apply to live stream |
| Practice safely | Turn on **Preview** |
| Update later | Settings → General → Check for Updates (after 2.2.0) |

---

*This handoff describes SLS Robuster as of version **2.2.0**. If the app’s version badge (sidebar / Settings) shows a newer number, prefer that build’s behavior and the latest notes on the Updates Releases page.*
