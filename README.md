\# Smart AI Lead Management System



An AI-powered lead management automation system built using Google Apps Script and Gemini AI.



\## Overview

This project automates customer lead management by collecting information through Google Forms, generating personalized email responses using Gemini AI, and sending emails automatically through Gmail.



\## Technologies Used

\- Google Apps Script

\- Google Gemini API

\- Google Forms

\- Google Sheets

\- Gmail



\## Features

\- Automated lead collection

\- AI-generated personalized emails

\- Automatic email delivery

\- Google Sheets lead management

\- Automation status tracking

\- Error handling and retry mechanism



\## Workflow

1\. Customer submits a Google Form.

2\. Google Apps Script processes the submission.

3\. Customer details are extracted from the spreadsheet.

4\. Gemini AI generates a personalized email.

5\. Gmail sends the email automatically.

6\. The response and status are saved in Google Sheets.



\## Setup

1\. Create a Google Form and link it to Google Sheets.

2\. Open the Apps Script project.

3\. Add `GEMINI\_API\_KEY` under Script Properties.

4\. Run `setupTrigger()` once and authorize permissions.

5\. Submit a test form response.



\## Security

API keys, passwords, and private customer information must never be committed to this repository.



\## Author

Yogesh

