# Xcode Projects Setup Guide for NFC

Since this is a source-only delivery, you need to configure your Xcode project manually to enable NFC scanning.

## 1. Add the Capability
1. Open your project in Xcode.
2. Select the **Root Project** node in the file navigator (left sidebar).
3. Select your **App Target** in the main view.
4. Go to the **Signing & Capabilities** tab.
5. Click the **+ Capability** button (top left of the tab).
6. Search for and add **"Near Field Communication Tag Reading"**.
   * *This will automatically create a `.entitlements` file in your project.*

## 2. Update Info.plist
You must provide a privacy usage description, or the app will crash instantly when you try to scan.

1. Go to the **Info** tab (next to Signing & Capabilities).
2. Right-click anywhere in the list and select **Add Row**.
3. Key: `Privacy - NFC Scan Usage Description` 
   * *Raw Key Name: `NFCReaderUsageDescription`*
4. Value: `This app scans guest cards to record services privately.`

## 3. Verify Hardware
*   **Simulator**: Does NOT support NFC. Use the debug "Simulate" buttons in the app.
*   **Physical Device**: Sideload the app to your iPhone.
    *   Ensure your Provisioning Profile supports NFC (most free/paid accounts do automatically).

## 4. Build & Run
1. Select your physical device.
2. **Cmd + R** to run.
3. Tap "Tap to Scan" and verify the system prompt appears.
