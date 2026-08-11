# SIMATS - SYNORA '26 Certificate Sharing Portal & Registration System

An advanced, high-capacity web portal for **SYNORA '26** (National Level Symposium conducted by the Department of Nanobiomaterials, SIMATS Engineering). 

This portal features fast, parallel certificate rendering, instant email dispatching via Google Apps Script, an executive BI analytics dashboard, dynamic status tracking, and Netlify deployment support.

---

## 🌟 Key Features

* **Instant Parallel PDF Generation**: Renders high-resolution vector PDF certificates in parallel using `jsPDF`.
* **1-Click & Batch Email Dispatcher**: Dispatches certificates directly to team leaders' emails via Google Apps Script with auto-retry and concurrency chunking for 150+ teams.
* **Smart Skip Protection**: Automatically skips teams that have already received certificates (`status === 'Emailed'`).
* **Role-Based Access Control**: Restricts certificate preview, ZIP downloads, and mail dispatch to authenticated organizers while giving participants a clean registration confirmation summary.
* **Real-Time Status Synchronization**: Updates local database, UI badges, and Google Sheets Column H to `Emailed` instantly.
* **Certificate Authenticator**: Anyone can verify certificate authenticity by entering a unique Verification ID (e.g. `SYN26-X8F9A2`).
* **Executive BI Analytics Dashboard**: Features live progress bars, metrics cards, and an interactive HTML5 Canvas Doughnut Chart displaying delivery ratios and institutional statistics.
* **Cosmic UI & Delight Effects**: Floating gold particle constellation canvas background (`#particle-canvas`) and metallic confetti celebration bursts upon registration.

---

## 🛠️ Project Structure

```
├── index.html        # Single-Page Application (Frontend, Dashboard, PDF Renderer, UI)
├── code.gs           # Google Apps Script Backend (Web App API, Google Sheets & MailApp)
├── netlify.toml      # Netlify deployment configuration & security headers
├── manifest.json     # PWA Web App Manifest
└── template.png      # High-resolution certificate background template
```

---

## 🚀 Live Netlify Deployment

The portal is configured for instant Netlify deployment with SPA routing and asset caching headers.
* **Netlify Config**: `netlify.toml`
