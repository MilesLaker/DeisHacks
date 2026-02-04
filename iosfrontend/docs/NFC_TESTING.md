# NFC Testing Guide

## 1. Requirements
*   **Device**: Physical iPhone 7 or newer (Simulator does NOT support CoreNFC).
*   **Apple ID**: Must be signed in with a valid provisioning profile (Free or Paid).
*   **Capabilities**: Project must have "Near Field Communication Tag Reading" capability enabled.
*   **Info.plist**: Must contain `NFCReaderUsageDescription` key.

## 2. Preparing the Cards
You need NDEF-formatted NFC tags (NTAG213, 215, 216 etc.).
Write a **Text Record** to the tag with the following content:
*   `guest_001`
*   `guest_12345`

*Note: The app expects the text to match the regex `^guest_\d{1,6}$`.*

## 3. Testing Steps (On Device)
1.  Open the App.
2.  Observe the status: **"Ready to Scan"**.
3.  Tap the large **"Tap to Scan"** button in the center.
    *   *iOS System Prompt should appear: "Ready to Scan / Hold card near phone".*
4.  Hold the NFC card to the top-back of the phone.
5.  **Success Flow**:
    *   System Prompt shows checkmark "Card read."
    *   App displays "Guest Found" or "Guest Not Found" alert.
    *   **Debug Label** (if Debug build): Shows "Last NFC: guest_###".
6.  **Failure Flow (Invalid Card)**:
    *   System Prompt shows red error "Invalid card. Please use a CDCW guest card."
    *   **Debug Label**: Shows specific error (e.g., "Payload mismatch").

## 4. Simulator Testing
NFC is disabled on Simulator. Use the yellow **debug box** at the bottom:
1.  Enter `guest_999` in the text field.
2.  Tap **Simulate**.
3.  Observe the same Alerts as the physical scan.

## 5. Troubleshooting
*   **"Session Invalidated" immediately**: Ensure `Info.plist` has the usage string.
*   **"Scanning not supported"**: You are on Simulator or iPad.
*   **No Read**: Remove phone case, ensure card is NDEF formatted (not raw).
