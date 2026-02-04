package com.cdcw.nfc

import android.content.Intent
import android.nfc.NfcAdapter
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.ui.Modifier
import com.cdcw.nfc.navigation.AppNavigation
import com.cdcw.nfc.nfc.NFCService

class MainActivity : ComponentActivity() {
    private lateinit var nfcService: NFCService

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        
        // Init NFC Service
        nfcService = NFCService(this) { guestId ->
            // Handle Global Scan
            // In a real app we might use a shared ViewModel or EventBus
            // For now, we'll let the Navigation manager handle this via a singleton or broadcast
        }

        setContent {
            MaterialTheme {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    AppNavigation(nfcService)
                }
            }
        }
    }
    
    override fun onResume() {
        super.onResume()
        nfcService.enableForegroundDispatch()
    }
    
    override fun onPause() {
        super.onPause()
        nfcService.disableForegroundDispatch()
    }
    
    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        // Pass to NFC Service
        nfcService.processIntent(intent)
    }
}
