# Email Automation Scripts Repository

## Overview

This repository is used to **store and organize scripts related to sending emails**, built with **Google Apps Script** and connected to **Google Sheets**.
The scripts are intended for automating common communication tasks such as membership renewals, reminders, and notifications.

Each folder groups scripts by **purpose or campaign**, making them easier to maintain, reuse, and update.

---

## Repository Structure

```text
.
├── renewals/
│   └── Renovaciones 2026/
│       └── sendPersonalizedEmailsRenovaciones.gs
│
├── cards/
│   └── Pick up your card/
│       └── sendPersonalizedEmails.gs
│
└── README.md
```

---

## Folder Description

### `renewals/`

Scripts related to **membership_renewal**.

* **Renovaciones 2026/**

  * Automation for sending personalized renewal emails.
  * Reads data from Google Sheets.
  * Applies business logic (renewed, fee pending, zero fee, etc.).
  * Sends HTML emails with payment instructions when applicable.

---

### `pickup_ids_from_office/`

Scripts related to **membership card management**.

* **Pick up your card/**

  * Sends reminder emails to members who have not yet collected their AEGEE card.
  * Uses bilingual (Spanish / English) email content.
  * Filters recipients based on a “pending card” column in Google Sheets.

---

## General Notes

* All scripts:

  * Are designed to be run from **Google Apps Script**.
  * Read data from **Google Sheets**.
  * Use `MailApp` to send emails.
* Each script folder should include:

  * The `.gs` file(s)
  * A small README explaining how to configure and run the script

---

## Purpose of this Repository

* Centralize all **email-related automation scripts**
* Make scripts:

  * Reusable
  * Easy to audit
  * Easy to adapt for future campaigns
* Avoid duplication and reduce configuration errors

---

If you plan to add new scripts, create a new folder describing the campaign or purpose and include a short README explaining usage and configuration.
