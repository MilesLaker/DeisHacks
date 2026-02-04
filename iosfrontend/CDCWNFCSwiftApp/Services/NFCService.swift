import Foundation
import Combine
import CoreNFC

final class NFCService: NSObject, ObservableObject, NFCNDEFReaderSessionDelegate {
    static let shared = NFCService()
    
    @Published var isScanning = false
    @Published var lastScannedId: String?
    @Published var lastError: String?
    
    private var session: NFCNDEFReaderSession?
    
    // Callback for View to subscribe to or trigger actions
    var onScan: ((String) -> Void)?
    
    override private init() {
        super.init()
    }
    
    func beginScan() {
        guard NFCNDEFReaderSession.readingAvailable else {
            self.lastError = "NFC Scanning not supported on this device."
            return
        }
        
        // Clean state
        self.lastError = nil
        self.isScanning = true
        
        session = NFCNDEFReaderSession(delegate: self, queue: nil, invalidateAfterFirstRead: true)
        session?.alertMessage = "Hold card near the top of the phone."
        session?.begin()
    }
    
    func stopScan() {
        session?.invalidate()
        isScanning = false
    }
    
    // MARK: - NFCNDEFReaderSessionDelegate
    
    func readerSession(_ session: NFCNDEFReaderSession, didInvalidateWithError error: Error) {
        DispatchQueue.main.async {
            self.isScanning = false
            
            // Ignore user cancellation
            if let nfcError = error as? NFCReaderError, nfcError.code == .readerSessionInvalidationErrorUserCanceled {
                return
            }
            
            self.lastError = error.localizedDescription
        }
    }
    
    func readerSession(_ session: NFCNDEFReaderSession, didDetectNDEFs messages: [NFCNDEFMessage]) {
        guard let message = messages.first else { return }
        
        for record in message.records {
            // Check for Text Record
            if let text = decodeTextRecord(record) {
                // Validate Guest ID Format
                let cleaned = text.trimmingCharacters(in: .whitespacesAndNewlines)
                
                // Allow case-insensitive check, but typically guest_ is lowercase
                // Regex: ^guest_\d{1,6}$
                // Simple check:
                if cleaned.lowercased().hasPrefix("guest_") {
                    // Valid
                    DispatchQueue.main.async {
                        self.lastScannedId = cleaned
                        self.onScan?(cleaned)
                    }
                    session.alertMessage = "Card read."
                    session.invalidate()
                    return
                }
            }
        }
        
        // If we get here, no valid guest ID found
        session.invalidate(errorMessage: "Invalid card. Please use a CDCW guest card.")
    }
    
    // MARK: - NDEF Text Decoding Logic
    
    private func decodeTextRecord(_ record: NFCNDEFPayload) -> String? {
        // Must be Well Known Type Text
        guard record.typeNameFormat == .nfcWellKnown,
              record.type == Data([0x54]) // "T"
        else { return nil }
        
        let payload = record.payload
        guard !payload.isEmpty else { return nil }
        
        // Status Byte (Table 3 of NFC Record Type Definition)
        let statusByte = payload[0]
        let isUTF16 = (statusByte & 0x80) != 0
        let languageCodeLength = Int(statusByte & 0x3F)
        
        guard payload.count > 1 + languageCodeLength else { return nil }
        
        let textData = payload.dropFirst(1 + languageCodeLength)
        
        if isUTF16 {
            return String(data: textData, encoding: .utf16)
        } else {
            return String(data: textData, encoding: .utf8)
        }
    }
    // Added req required protocol stub (even if not used for simple reads, sometimes compiler complains)
    func readerSession(_ session: NFCNDEFReaderSession, didDetect tags: [NFCNDEFTag]) {
        // Not used
    }
}
