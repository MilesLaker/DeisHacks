import NfcManager, { NfcTech, Ndef } from 'react-native-nfc-manager';
import { Platform } from 'react-native';

// Initialize
async function init() {
    if (Platform.OS === 'web') return; // Web support skipped
    await NfcManager.start();
}

// Read Tag ID
async function readTag(): Promise<string | null> {
    if (Platform.OS === 'web') return null;

    try {
        // Request technology
        await NfcManager.requestTechnology(NfcTech.Ndef);

        const tag = await NfcManager.getTag();
        return tag?.id || null;
    } catch (ex) {
        console.warn('NFC Read Error:', ex);
        return null;
    } finally {
        // Stop
        NfcManager.cancelTechnologyRequest();
    }
}

// Cancel
function cancel() {
    NfcManager.cancelTechnologyRequest();
}

export const NFCService = {
    init,
    readTag,
    cancel,
};
