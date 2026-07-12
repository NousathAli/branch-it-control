# Security Notice

This project is currently a functional prototype and should not be used to store live company passwords in a public deployment without further security work.

Current limitations include:

- Application users are managed in client-side application data.
- Firebase anonymous authentication is used for shared database access.
- Firestore data is stored largely in a single application document.
- The credential vault uses browser-side encryption and a user-provided passphrase.
- Frontend environment variables and Firebase web configuration are not secrets.

Before production credential use:

1. Migrate users to Firebase Authentication or Microsoft Entra ID.
2. Enforce role-based Firestore Security Rules.
3. Store credential records separately from general app data.
4. Perform encryption/decryption through a trusted backend or managed secret store.
5. Add immutable audit logs for reveal, copy, edit, import, and export actions.
6. Enable MFA for privileged administrators.
7. Review Firebase App Check and rate limiting.

Never commit exported credential spreadsheets, database backups, `.env`, or production secrets.
