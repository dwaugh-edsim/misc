# About the Reunion Playlist Portal (`index.html`)

This document explains the architecture, design, and functionality of the `index.html` file to assist other AI models in understanding and maintaining the codebase.

---

## 📋 Context & Purpose
The `index.html` file serves as the interactive frontend portal for the **Classes of '86, '87, and '88 40th High School Reunion Playlist**. It allows attendees to review a curated list of 80s party songs, upvote/downvote tracks, and suggest new songs to play between the live band's sets.

---

## 🎨 Design System (80s Retro Synthwave)
* **Aesthetics:** High-energy, polished dark-mode synthwave style using deep purples, hot pinks, and neon blues.
* **Layout:** Centered `800px` content column with a CSS-rendered wireframe grid background overlay.
* **Header:** Features a custom hero banner image (`retro_80s_header.png`) with glowing neon text shadows.
* **Collage:** Staggers Polaroid-style memory photos (from the `danceimages/` folder) down the left and right margins of the viewport with random rotations. 
* **Responsive Guardrails:** The scattered Polaroids hide on screens narrower than `1080px` to prevent overlapping central content. The main song grid shifts down to 1 column on mobile devices.

---

## ⚙️ Key Functionalities

### 1. Data Fetching & Webhook Sync
* **API URL:** Hardcoded to a Google Apps Script Web App:
  `https://script.google.com/macros/s/AKfycbzm1U5F9W50BleZirSm2iHjKBmkgXywdByneFiSsehjP1TnM4AB8S3W5UOo0GvB4KaAfA/exec`
* **GET Flow:** On load, the page sends a GET request to the webhook. It expects a JSON object containing:
  - `playlist`: An array of song objects (`id`, `title`, `artist`, `year`, `votes`).
  - `images`: An array of paths for the Polaroid collage.
* **Fallback:** If the network request fails, it defaults to a built-in static array of 44 curated songs and local mock data.

### 2. Search & Filtering
* A live search input filters songs by `title`, `artist`, or `year` in real-time by toggling CSS display attributes.

### 3. Upvoting/Downvoting
* Users can cast upvotes or downvotes.
* **Toggle Logic:** If a user clicks the same vote button a second time, the vote is undone.
* **Local Storage:** User votes are tracked locally (`localStorage.getItem("votesMap")`) to visually highlight voted buttons and prevent duplicate voting.
* **Syncing:** Votes are sent via POST `no-cors` to the webhook with payload: `{ action: "vote", id: "<id>", type: "up" | "down" }`.

### 4. Simplified Song Suggestions
* A streamlined form at the bottom allows users to submit suggestions.
* **Requirement:** Only requires a **Song Title**.
* **Syncing:** Submits via POST `no-cors` to the webhook with payload: `{ action: "suggest", title: "<title>", artist: "", year: null, comment: "" }`.
