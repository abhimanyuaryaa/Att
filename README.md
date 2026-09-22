# Attendance Vercel + Google Apps Script

Architecture:
Vercel frontend -> Vercel API proxy -> Google Apps Script -> Google Sheet

No Google Cloud service-account key is required.

## One-time Apps Script update

Replace the existing Code.gs with `apps-script/Code.gs`, then deploy a NEW VERSION of the existing Apps Script web app.

Keep the web app accessible to anyone with the URL.

The configured Apps Script URL is:
https://script.google.com/macros/s/AKfycby4gsETKJmTCDopdfY7J7R_sAO8ylyDamFKiP94ANcyZ_KHjVjunUuZMb9LD1YY7BWeBQ/exec

The Vercel project already contains the matching API token and Apps Script URL, so no Spreadsheet ID, Sheet Name, Year, or Vercel environment variable is required.

Keep your GitHub repository private because the API token is stored in the server-side files.
