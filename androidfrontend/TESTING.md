# Android Testing Guide

## 1. Emulator Testing (No NFC)
1.  Launch App in Android Studio Emulator.
2.  Observe "Ready to Scan" in Blue.
3.  Scroll to the **Yellow Debug Box**.
4.  Enter `guest_123` and tap **"Simulate Scan"**.
5.  Verify navigation to **Service Entry** screen.
6.  Toggle Services, set Meals to 1.
7.  Tap **Submit**.
8.  Verify return to Home Screen.

## 2. Physical Device Testing (Real NFC)
1.  Deploy to a real Android phone.
2.  Prepare an NFC Tag with NDEF Text record: `guest_001`.
3.  Open the App.
4.  **Tap the card** against the back of the phone.
5.  **Success**:
    *   Status circle turns Green ("Scanning...").
    *   App automatically navigates to **Service Entry** for `guest_001`.
6.  **Failure**:
    *   Use a card with text `invalid_id`.
    *   App should NOT navigate. (Error handling logic in `NFCService` sets `lastError`, debug UI shows it).

## 3. Clothing Store
1.  Scan a guest (or Simulate).
2.  (Note: Current Port navigates directly to Service Entry for parity test).
3.  To test Clothing: Modify `ScanScreen` implementation to route to `Routes.clothing(id)` temporarily OR implement the full Alert choice dialog.
    *   *Default implementation routes to Service Entry as per prompt MVP "Route through the same handler".*
