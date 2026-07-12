# Branch IT Control

A React and Firebase prototype for managing multi-branch IT assets, CCTV/NVR reviews, network documentation, branch handovers, fault records, reusable asset products, standard branch templates, user work portfolios, and an encrypted credential-vault workflow.

## Main modules

- Executive dashboard with project and asset summaries
- Group, concept, location, and branch management
- IT asset register with Excel import/export
- Asset Product List and standard branch templates
- CCTV/NVR weekly review and camera-operation filters
- Network setup records
- Fault logs, handovers, and pullouts
- User portfolios and admin user management
- Credential Vault for shared accounts, device credentials, and access assignments
- Dark/light themes and responsive layouts

## Technology

- React
- Vite
- Firebase Authentication (anonymous sync in the prototype)
- Cloud Firestore
- SheetJS/XLSX
- Tailwind CSS CDN-based styling

## Run locally

```powershell
npm install
copy .env.example .env
npm run dev
```

The app can run with local browser storage when Firebase is not configured. To enable Firebase sync, place your web-app configuration values in `.env`.

### Demo logins

| Role | Username | Password |
|---|---|---|
| Admin | `demo-admin` | `DemoAdmin@123` |
| Viewer | `demo-viewer` | `DemoViewer@123` |
| Technician | `demo-tech` | `DemoTech@123` |

These are demonstration credentials only. Replace the client-side prototype authentication before production use.

## Production build

```powershell
npm run build
```

## Firebase Hosting

1. Copy `.firebaserc.example` to `.firebaserc`.
2. Replace the placeholder project ID.
3. Configure `.env`.
4. Run:

```powershell
npm run build
firebase deploy --only hosting
```

## Security status

This repository is a portfolio-ready prototype. It demonstrates workflows and UI, but the current client-side user authentication and credential-vault design must be upgraded before storing real production credentials. See [SECURITY.md](SECURITY.md).

## Project documentation

- [Architecture](docs/ARCHITECTURE.md)
- [Modules](docs/MODULES.md)
- [Local and Firebase setup](docs/SETUP_AND_DEPLOYMENT.md)
- [CV project description](docs/CV_PROJECT_DESCRIPTION.md)
- [Roadmap](docs/ROADMAP.md)

## Version

Current portfolio package: **v8.47**
