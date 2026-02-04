package com.cdcw.nfc.ui

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.navigation.NavController
import com.cdcw.nfc.navigation.Routes
import com.cdcw.nfc.nfc.NFCService

@Composable
fun ScanScreen(navController: NavController, nfcService: NFCService) {
    var debugId by remember { mutableStateOf("guest_001") }
    val lastScanned by nfcService.lastScannedId.collectAsState()
    val isScanning by nfcService.isScanning.collectAsState()
    
    // Auto-Navigate on Scan
    LaunchedEffect(lastScanned) {
        lastScanned?.let { id ->
            handleScan(navController, id)
        }
    }
    
    Column(
        modifier = Modifier.fillMaxSize().padding(16.dp),
        horizontalAlignment = Alignment.CenterHorizontally,
        verticalArrangement = Arrangement.Center
    ) {
        Text("Service Entry", style = MaterialTheme.typography.displaySmall)
        
        Spacer(modifier = Modifier.height(32.dp))
        
        // Status Circle
        Box(contentAlignment = Alignment.Center) {
             CircularProgressIndicator(
                 modifier = Modifier.size(200.dp),
                 progress = 1f, // Full circle
                 color = if (isScanning) Color.Green else Color.Blue,
                 strokeWidth = 4.dp
             )
             
             if (isScanning) {
                 Text("Scanning...", color = Color.Green, style = MaterialTheme.typography.titleMedium)
             } else {
                 Text("Ready to Scan", color = Color.Blue, style = MaterialTheme.typography.titleMedium)
             }
        }
        
        Spacer(modifier = Modifier.height(32.dp))
        
        Button(
            onClick = { navController.navigate(Routes.serviceEntry("anonymous", false, true)) },
            colors = ButtonDefaults.buttonColors(containerColor = Color.Gray)
        ) {
            Text("Anonymous Entry")
        }
        
        Spacer(modifier = Modifier.weight(1f))
        
        // Debug Area
        Card(
            colors = CardDefaults.cardColors(containerColor = Color(0xFFFFF8E1)),
            modifier = Modifier.fillMaxWidth()
        ) {
            Column(modifier = Modifier.padding(16.dp)) {
                Text("DEBUG TOOLS", style = MaterialTheme.typography.labelSmall)
                
                TextField(
                    value = debugId,
                    onValueChange = { debugId = it },
                    label = { Text("Test ID") }
                )
                
                Button(onClick = {
                    handleScan(navController, debugId)
                }) {
                    Text("Simulate Scan")
                }
            }
        }
    }
}

fun handleScan(navController: NavController, id: String) {
    // Determine if guest exists logic would go here if we had a local store ready.
    // For parity with "Mock" flow or stateless flow:
    // Route to Service Entry.
    // iOS App does:
    // Alert(Found) -> ServiceEntry.
    // We'll skip the Alert for speed and just go to ServiceEntry.
    // But wait, the Prompt says "Route through the same handler".
    // We are.
    navController.navigate(Routes.serviceEntry(id, false, false))
}
