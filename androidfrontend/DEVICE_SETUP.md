# Android Device Setup Guide (Samsung)

To run the app on your Samsung phone, you need to enable Developer Mode and USB Debugging.

## 1. Enable Developer Options
1.  Open **Settings** on your phone.
2.  Scroll down and tap **About phone**.
3.  Tap **Software information**.
4.  Find **Build number** and tap it **7 times** rapidly.
5.  Enter your PIN/Pattern if prompted.
6.  You will see a toast message: *"Developer mode has been enabled."*

## 2. Enable USB Debugging
1.  Go back to the main **Settings** menu.
2.  Scroll to the very bottom; you should now see **Developer options**. Tap it.
3.  Scroll down to the **Debugging** section.
4.  Toggle **USB debugging** to **ON**.
5.  Confirm the warning dialog.

## 3. Connect to Computer
1.  Plug your phone into your computer via USB.
2.  On your phone, a prompt will appear: *"Allow USB debugging?"*
3.  Check **"Always allow from this computer"** and tap **Allow**.

## 4. Run from Android Studio
1.  Open the project in Android Studio.
2.  Look at the toolbar at the top.
3.  In the device dropdown (next to the Run button), you should now see your **Samsung ...** device listed instead of "Pixel Emulator".
4.  Select your phone.
5.  Click the Green **Run** (Play) button.

## 5. Enable NFC
1.  Pull down the Quick Settings shade on your phone.
2.  Ensure **NFC** is turned **ON**.
3.  The app requires NFC to be enabled to scan cards.
