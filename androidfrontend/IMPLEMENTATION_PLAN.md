# Android Implementation Plan

## Checkpoint 1: Project Setup (In Progress)
- [x] Create Directory Structure (`androidfrontend/`)
- [ ] Create `AndroidManifest.xml` (Permissions: INTERNET, NFC)
- [ ] Create `build.gradle.kts` (Dependencies: Compose, lifecycle, networking)
- [ ] Create `MainActivity.kt` (Entry Point, Theme)

## Checkpoint 2: Core Models & Sync (Pending)
- [ ] `PayloadModels.kt`: Data classes matching backend contract.
- [ ] `SyncService.kt`: Retrofit/OkHttp client + Offline Queue logic.
- [ ] `GuestStore.kt`: Local persistence (SharedPreferences/Room).

## Checkpoint 3: NFC Logic (Pending)
- [ ] `NFCService.kt`: `NfcAdapter` management, Foreground Dispatch.
- [ ] NDEF Text Record parsing logic (Status byte handling).
- [ ] Regex validation: `^guest_\d{1,6}$`.

## Checkpoint 4: UI & Navigation (Pending)
- [ ] `ScanScreen.kt`: Ready state, Scan Button (Debug vs Real).
- [ ] `ServiceEntryScreen.kt`: Service toggles & counters.
- [ ] `ClothingStoreScreen.kt`: Budget fetch & purchase.
- [ ] `NewGuestScreen.kt`: Registration form.
- [ ] `NavManager.kt`: Centralized routing.

## Checkpoint 5: Verification (Pending)
- [ ] Test on Emulator (Simulate Scan).
- [ ] Verify Code Parity with iOS/RN logic.
