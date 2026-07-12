# Setup and Deployment

## Requirements

- Node.js LTS
- npm
- Firebase CLI for deployment

## Local setup

```powershell
npm install
copy .env.example .env
npm run dev
```

## Build

```powershell
npm run build
```

## Firebase setup

1. Create a Firebase Web App.
2. Enable Anonymous Authentication if using the current prototype sync.
3. Create a Firestore database.
4. Copy the Firebase Web configuration into `.env`.
5. Copy `.firebaserc.example` to `.firebaserc` and set the project ID.
6. Review Firestore Security Rules before deployment.

## Deploy Hosting

```powershell
firebase login
firebase use your-firebase-project-id
npm run build
firebase deploy --only hosting
```
