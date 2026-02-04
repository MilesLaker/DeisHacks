# Testing Guide: CDCW NFC Swift App

## 1. Setup
Since this is a source-only delivery:
1. Create a newly generated SwiftUI Project in Xcode.
2. Drag and dropping the `CDCWNFCSwiftApp/` folders into the project.
3. Ensure `Config/AppConfig.swift` has the correct API URL.
4. Add "Near Field Communication Tag Reading" capability in Signing & Capabilities (Required for `CoreNFC` if implementing real hardware scan).

## 2. Test Scripts

### A. Simulating a Scan (Simulator)
1. Launch App in Simulator.
2. Tap "Simulate" in the yellow Debug box.
3. Verify "Guest Not Found" alert appears (if Guest 001 doesn't exist locally).
4. Tap "Rand Tap" to generate a random ID.

### B. New Guest Flow
1. Tap "Rand Tap".
2. Select "New Guest" (or equivalent flow via Search -> "Create New").
3. Enter "Test User".
4. Submitting should navigate to "Service Entry".
5. Verify `UPDATE_GUEST` payload in Console Logs.
6. Verify Status Circle on Home Screen says "Ready to Scan" after backing out.

### C. Service Entry
1. From "Service Entry" screen:
2. Toggle "Shower" to YES.
3. Set Meals to 2.
4. Tap "Submit Entry".
5. Verify `LOG_SERVICE` is queued.
6. Check Debug Sync Box: "Queue: 1".
7. Tap "Sync Now".
8. Verify "Queue: 0" and "Success" log in console (if backend connected).

### D. Clothing Store
1. Enter Clothing Store for a guest.
2. Verify "Current Budget".
3. Tap "Refresh Budget" -> Should fetch from Sheets.
4. Select quantity 1.
5. Tap "Confirm".
6. Verify Local Budget decrements immediately.
7. Verify `CLOTHING_PURCHASE` payload in Queue.

### E. Anonymous Entry
1. Tap "Anonymous Entry" button on specific Home Screen location.
2. Verify "Guest ID: anonymous".
3. Facilities (Shower/Laundry) should be HIDDEN.
4. Meals should be visible.
5. Submit -> Verify `ANONYMOUS_ENTRY` payload.

### F. Offline / Retry
1. Turn off Wifi/Network on Simulator.
2. Perform a Service Entry.
3. "Sync Now" -> Should show Error in Red on Debug box.
4. Turn Wifi On.
5. "Sync Now" -> Should clear queue.
