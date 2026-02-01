# CDCW NFC Guest Management System

A robust mobile application designed for CDCW to manage guest services, track budgets, and streamline check-ins using NFC technology. The system utilizes a **React Native** mobile app for staff and a **Google Sheets** backend for data storage and reporting.

## 🚀 Features

-   **NFC Card Integration**: Instantly pull up guest profiles by tapping their ID card.
-   **Service Logging**: Track showers, laundry, meals, and hygiene kits with a single tap.
-   **Budget Management ("Felton Bucks")**: Automatically track and deduct clothing purchases from guest budgets.
-   **Offline-First Architecture**: Actions are queued locally and synced when internet connectivity is available, ensuring operations never stop.
-   **Anonymous Entry**: Fast-track entry for guests without IDs.
-   **Cloud Backend**: Leverages Google Sheets for free, accessible, and real-time database management without complex infrastructure.

## 🛠 Tech Stack

### Frontend (Mobile App)
-   **Framework**: React Native (via **Expo**)
-   **Language**: TypeScript
-   **State Management**: Redux Toolkit + Redux Persist (for offline queueing)
-   **NFC**: `react-native-nfc-manager`
-   **Navigation**: React Navigation

### Backend (Serverless)
-   **Logic**: Google Apps Script (GAS)
-   **Database**: Google Sheets (`Guests`, `Log`, `Analytics` tabs)
-   **API**: Web App deployment of GAS (JSON-over-HTTP)

## 📂 Project Structure

```text
/
├── backend/
│   └── sheets/
│       └── code.gs       # Google Apps Script logic (Deploy as Web App)
├── frontend/
│   ├── App.tsx           # Entry point
│   ├── app/              # Screens and Navigation
│   │   ├── Navigation.tsx
│   │   └── screens/      # React Native Screens (Home, ServiceEntry, etc.)
│   ├── redux/            # State & Logic
│   │   └── slices/       # syncSlice.ts handles the Offline Queue & API
│   └── services/         # Hardware integration (NFC)
```

## ⚡ Setup & Installation

### 1. Backend Configuration (Google Sheets)
1.  Create a new **Google Sheet**.
2.  Create three tabs: `Guests`, `Log`, `Analytics`.
    -   *Guests Columns*: `Guest ID`, `Name`, `Felton Bucks Budget`, `Last Visit`
    -   *Log Columns*: `Date`, `Time`, `Guest ID`, `Service Type`, `Event ID`
3.  Open **Extensions > Apps Script**.
4.  Copy the contents of `backend/sheets/code.gs` into the script editor.
5.  **Critical Config**:
    -   Update `SHEET_ID` in `code.gs` with your Google Sheet ID.
6.  **Deploy**:
    -   Click `Deploy` > `New deployment`.
    -   Type: `Web app`.
    -   Execute as: `Me` (your account).
    -   Who has access: `Anyone` (or `Anyone with Google account` depending on security needs).
7.  **Copy URL**: Save the "Web App URL" (ends in `/exec`).

### 2. Frontend Configuration
1.  Navigate to the frontend directory:
    ```bash
    cd frontend
    ```
2.  Install dependencies:
    ```bash
    npm install
    ```
3.  **Link Backend**:
    -   Open `frontend/redux/slices/syncSlice.ts`.
    -   Replace `const API_URL = '...'` with your **Google Web App URL** from step 1.

### 3. Running the App
Start the Expo development server:
```bash
npm start
```
-   **Android**: Press `a` (Requires Android Emulator or connected device).
-   **iOS**: Press `i` (Requires Xcode or iPhone).
-   **Physical Device**: Install "Expo Go" app and scan the QR code.

> **Note**: For NFC functionality to work, you must use a physical Android/iOS device with NFC support. Simulators cannot test NFC taps.

## 📱 detailed Usage

### Service Entry
1.  **Tap Card**: Staff taps a guest's NFC card on the phone.
2.  **View Profile**: Guest name and budget appear.
3.  **Log Services**: Toggle "Shower", "Laundry" or increment "Meals".
4.  **Confirm**: Hitting "Submit" adds the entry to the sync queue.

### Clothing Store
1.  Select "Clothing Store" after scanning.
2.  Enter quantity of items.
3.  App checks local budget (if synced) or server budget.
4.  Deducts "Felton Bucks" and updates the sheet.

### Syncing
-   The app attempts to sync automatically on mount.
-   Manual trigger available in the hidden **Debug Menu** (visible in `__DEV__` mode).
-   Failed requests are retried automatically with exponential backoff logic (implied in code).

## 🛡 License
This project is open-source.
