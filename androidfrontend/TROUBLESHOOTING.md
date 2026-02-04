# Android Troubleshooting Guide

## 1. "Run" Button is Greyed Out or Missing
**Symptoms**: The green Play button in the top toolbar is disabled or not visible.
**Causes & Fixes**:

1.  **Gradle Sync in Progress** (Most Common):
    *   Look at the bottom right of Android Studio. Is there a progress bar saying "Gradle Sync"?
    *   **Fix**: Wait for it to finish. The Run button only enables after a successful sync.

2.  **No Configuration Selected**:
    *   Look at the dropdown menu to the left of the Run button (or where it should be).
    *   It might say "Add Configuration..." or be empty.
    *   **Fix**: Click the dropdown -> Select **"app"**.
    *   *If "app" is missing*: Click "Edit Configurations" -> "+" -> "Android App" -> Module: "app".

3.  **Gradle Sync Failed**:
    *   Click the **"Build"** tab at the bottom of the screen.
    *   Are there red errors? (e.g., "SDK location not found", "Could not resolve...").
    *   **Fix**: You must fix these errors first. A failed sync prevents the app from being recognized.

## 2. Device Not Recognized
**Symptoms**: The device dropdown says "No Devices" or "Loading..." indefinitely.
**Fixes**:
1.  **Check USB Debugging**: Unplug and replug. Look for the "Allow USB debugging?" prompt on the phone screen.
2.  **Cable Issue**: Ensure you are using a data cable, not just a charging cable.
3.  **Restart ADB**:
    *   Open Terminal in Android Studio (bottom bar).
    *   Run: `adb kill-server`
    *   Run: `adb start-server`

## 3. "SDK Location not found"
**Cause**: `local.properties` file is missing.
**Fix**:
1.  Android Studio usually generates this automatically on launch.
2.  If not, create `androidfrontend/local.properties`.
3.  Add: `sdk.dir=/Users/YOUR_USER/Library/Android/sdk` (Verify your actual SDK path via Android Studio -> Settings -> Languages & Frameworks -> Android SDK).
