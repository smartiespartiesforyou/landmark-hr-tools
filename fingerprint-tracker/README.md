# Fingerprint Tracker — development branch

This branch is reserved for the Landmark South fingerprint appointment and NCAC-code tracking pilot.

**Do not deploy this branch with actual employee/applicant data.**

Pilot source ZIP: `fingerprint-tracker-pilot.zip`, provided in the ChatGPT conversation on October 13, 2026. It has not yet been copied into this branch.

The pilot supports:
- Applicant, position, NCAC code, scheduled appointment date/time
- Spreadsheet CSV import and export
- Receipt tracking and review
- Planned Twilio MMS receipt intake

Before any live use: add individual HR sign-ins/MFA, audit logging, backups, retention/deletion, verified applicant identity matching, and an approved deployment configuration. Do not commit credentials or real employee records.

Recommended long-term arrangement: a dedicated private `fingerprint-tracker` repository and managed web service, separate from the existing HR Tools production system.
