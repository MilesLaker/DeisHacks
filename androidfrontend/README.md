# NFC Services App (Android)

Android implementation of the NFC Services App, built with Kotlin and Jetpack Compose.

## Structure
*   **`com.cdcw.nfc.models`**: Data classes matching the backend JSON contract.
*   **`com.cdcw.nfc.network`**: `SyncService` handling offline queue and Retrofit calls.
*   **`com.cdcw.nfc.nfc`**: `NFCService` handling Foreground Dispatch and NDEF Text parsing.
*   **`com.cdcw.nfc.ui`**: Jetpack Compose screens (`ScanScreen`, `ServiceEntry`, etc.).
*   **`com.cdcw.nfc.navigation`**: Centralized `NavHost`.

## Setup
1.  Open `androidfrontend/` in Android Studio.
2.  Wait for Gradle Sync.
3.  **Physical Device**:
    *   Connect Android phone with NFC.
    *   Enable Developer Mode.
    *   Runs `MainActivity` which enables NFC Foreground Dispatch.
4.  **Emulator**:
    *   Run app.
    *   Use "Simulate Scan" debug button at the bottom of Home Screen.

## Parity Notes
*   **NFC**: Scans NDEF Text records, strictly validates `guest_###` regex.
*   **Sync**: Queues offline requests to SharedPreferences (Stubbed) and retries.
*   **UI**: Matches iOS flow (Home -> Service Entry / Clothing).
