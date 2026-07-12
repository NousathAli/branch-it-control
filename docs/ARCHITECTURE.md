# Architecture

## Frontend

The application is a React single-page application built with Vite. Most current functionality is implemented in `src/App.jsx`, with Firebase initialization in `src/firebase.js`.

## Data flow

- React state manages branches, locations, assets, products, templates, network records, faults, handovers, portfolios, and vault records.
- Local browser storage provides offline persistence.
- When Firebase is configured, the prototype synchronizes a combined database snapshot to Firestore.
- Excel import/export uses SheetJS.

## Recommended production refactor

- Split `App.jsx` into feature modules and reusable components.
- Use separate Firestore collections for branches, assets, products, users, faults, and vault records.
- Replace last-write-wins snapshot sync with record-level writes and transactions.
- Introduce formal authentication, role claims, security rules, audit logs, and tests.
