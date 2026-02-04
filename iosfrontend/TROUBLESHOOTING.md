# Troubleshooting: NFC Errors

## 1. FrameWork Error: Missing Info.plist
**Error**: `Framework .../CoreNFC.framework did not contain an Info.plist`
**Fix**:
1.  Go to **General** tab > **Frameworks, Libraries, and Embedded Content**.
2.  Change **CoreNFC** from "Embed & Sign" to **"Do Not Embed"**.

## 2. Runtime Error: "NFC Scanning not supported on this device"
**Symptoms**: You are on a real iPhone (not Simulator), but the app immediately says "Not Supported" when you tap Scan.
**Cause**: The App **Entitlement** is missing. Even if your phone has NFC, the App is not "authorized" by Xcode to use it.

**Fix**:
1.  Go to **Signing & Capabilities** tab in Xcode.
2.  Check if **"Near Field Communication Tag Reading"** is listed.
    *   **If missing**: Click **+ Capability** and add it.
    *   **If present**: Click the "Trash" icon to delete it, then re-add it. This forces Xcode to regenerate the `.entitlements` file.
3.  **Check the Entitlements File**:
    *   Look in your file explorer for a file ending in `.entitlements` (e.g., `CDCWNFCSwiftApp.entitlements`).
    *   Open it. It usually looks like a Property List.
    *   Ensure it has the key: `Tag NFC Reader Session Formats` (Array) -> Item 0: `NDEF`.
4.  **Provisioning Profile**:
    *   Sometimes "Free/Personal" Apple ID profiles glitch with NFC.
    *   Try changing the **Bundle Identifier** (e.g., add `.v2`) to force a new profile generation.

## 3. Crash on Scan
**Error**: App closes immediately when tapping Scan.
**Cause**: Missing `Info.plist` description.
**Fix**:
1.  Go to **Info** tab.
2.  Ensure `Privacy - NFC Scan Usage Description` is present and has text.
