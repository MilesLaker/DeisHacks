# iOS Native App - Implementation Plan & Status

## Checkpoint 1: Repo Analysis (Completed)

I have analyzed the existing React Native (`frontend/`) and Apps Script (`backend/`) code.

### 1. Structure
- **Navigation**: Simple Stack (`Home` -> `ServiceEntry`, `NewUserSetup`, `SearchGuest`, `Clothing`).
- **State**: Redux with local persistence (`guest` slice, `sync` slice).
- **Offline Strategy**: Optimistic updates + `syncSlice` queue (POST to backend with retry).

### 2. Backend Contract (Verified)
Base URL: *(Defined in Config)*

| Action | Payload | Source Screen |
| :--- | :--- | :--- |
| `LOG_SERVICE` | `{ guestId, services: { shower, laundry, meals, hygieneKits }, timestamp }` | `ServiceEntryScreen` |
| `ANONYMOUS_ENTRY` | `{ meals, timestamp }` | `ServiceEntryScreen` |
| `UPDATE_GUEST` | `{ id, name?, programs: {...}, ... }` | `NewUserSetupScreen` |
| `CLOTHING_PURCHASE` | `{ guestId, quantity, timestamp }` | `ClothingScreen` |
| `GET_BUDGET` | `{ guestId }` | `ClothingScreen` |
| `REPLACE_CARD` | `{ oldId, newId }` | `SearchGuestScreen` |
| `LOG_SERVICE` (Verify) | *Returns `{ status: 'success', log: { written: 1 } }`* | `syncSlice.ts` |

### 3. Special Logic
- **NFC Scan**:
    - Known Guest -> Alert(Service or Clothing).
    - Unknown Tag -> Alert(New Guest or Replacement).
- **Clothing**:
    - Fetches budget on mount.
    - Decrements locally (optimistic) and queues request.
- **Service Entry**:
    - "First Visit" calculated by backend, not app.
    - App sends simple counters.
- **Sync**:
    - Serial queue.
    - Deduplication via `eventId` (UUID).

---

## Checkpoint 2: Project Skeleton (In Progress)
- [x] Directory Structure Created
- [ ] App Entry Point (`CDCWNFCSwiftAppApp.swift`)
- [ ] Navigation (`AppNavigation.swift`)
- [ ] Networking (`NetworkClient.swift`)
- [ ] Config (`AppConfig.swift`)
