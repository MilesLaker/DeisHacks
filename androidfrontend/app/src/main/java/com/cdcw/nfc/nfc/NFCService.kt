package com.cdcw.nfc.nfc

import android.app.Activity
import android.app.PendingIntent
import android.content.Context
import android.content.Intent
import android.content.IntentFilter
import android.nfc.NdefMessage
import android.nfc.NdefRecord
import android.nfc.NfcAdapter
import android.util.Log
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import java.nio.charset.Charset

class NFCService(
    private val context: Context,
    private val onScan: (String) -> Unit
) {
    private var nfcAdapter: NfcAdapter? = NfcAdapter.getDefaultAdapter(context)
    
    private val _isScanning = MutableStateFlow(false)
    val isScanning: StateFlow<Boolean> = _isScanning
    
    private val _lastError = MutableStateFlow<String?>(null)
    val lastError: StateFlow<String?> = _lastError

    private val _lastScannedId = MutableStateFlow<String?>(null)
    val lastScannedId: StateFlow<String?> = _lastScannedId

    // Foreground Dispatch setup
    fun enableForegroundDispatch() {
        if (context !is Activity) return
        val intent = Intent(context, context.javaClass).apply {
            addFlags(Intent.FLAG_ACTIVITY_SINGLE_TOP)
        }
        val pendingIntent = PendingIntent.getActivity(
            context, 0, intent, 
            PendingIntent.FLAG_MUTABLE
        )
        
        val filters = arrayOf(IntentFilter(NfcAdapter.ACTION_NDEF_DISCOVERED).apply {
            try {
                addDataType("text/plain")
            } catch (e: IntentFilter.MalformedMimeTypeException) {
                throw RuntimeException("fail", e)
            }
        })
        
        nfcAdapter?.enableForegroundDispatch(context, pendingIntent, filters, null)
        _isScanning.value = true
    }

    fun disableForegroundDispatch() {
        if (context !is Activity) return
        nfcAdapter?.disableForegroundDispatch(context)
        _isScanning.value = false
    }
    
    fun processIntent(intent: Intent) {
        if (NfcAdapter.ACTION_NDEF_DISCOVERED == intent.action) {
            val rawMsgs = intent.getParcelableArrayExtra(NfcAdapter.EXTRA_NDEF_MESSAGES)
            if (rawMsgs != null) {
                val msgs = rawMsgs.map { it as NdefMessage }
                // Parse first message
                if (msgs.isNotEmpty()) {
                    val record = msgs[0].records[0]
                    parseRecord(record)
                }
            }
        }
    }
    
    private fun parseRecord(record: NdefRecord) {
        // Validate TNF and Type
        if (record.tnf == NdefRecord.TNF_WELL_KNOWN && record.type.contentEquals(NdefRecord.RTD_TEXT)) {
            try {
                val payload = record.payload
                // Get Status Byte
                val statusByte = payload[0]
                val encoding = if ((statusByte.toInt() and 0x80) == 0) "UTF-8" else "UTF-16"
                val languageCodeLength = statusByte.toInt() and 0x3F
                
                // Extract Text
                val text = String(
                    payload, 
                    languageCodeLength + 1, 
                    payload.size - languageCodeLength - 1, 
                    Charset.forName(encoding)
                )
                
                validateAndEmit(text)
                
            } catch (e: Exception) {
                _lastError.value = "Failed to decode NDEF: ${e.message}"
            }
        } else {
            _lastError.value = "Not a Text Record"
        }
    }
    
    private fun validateAndEmit(text: String) {
        val cleaned = text.trim()
        // Regex: ^guest_\d{1,6}$
        val regex = Regex("^guest_\\d{1,6}$", RegexOption.IGNORE_CASE)
        
        if (regex.matches(cleaned)) {
            _lastScannedId.value = cleaned
            _lastError.value = null
            onScan(cleaned)
        } else {
            _lastError.value = "Invalid ID Format: $cleaned"
        }
    }
}
