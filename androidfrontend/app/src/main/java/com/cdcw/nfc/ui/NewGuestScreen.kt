package com.cdcw.nfc.ui

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import androidx.navigation.NavController
import com.cdcw.nfc.models.APIAction
import com.cdcw.nfc.models.GuestPrograms
import com.cdcw.nfc.models.UpdateGuestPayload
import com.cdcw.nfc.network.SyncService
import java.time.Instant

@Composable
fun NewGuestScreen(navController: NavController, guestId: String) {
    var name by remember { mutableStateOf("") }
    
    Column(modifier = Modifier.padding(16.dp)) {
        Text("New Guest Registration", style = MaterialTheme.typography.headlineSmall)
        Text("ID: $guestId")
        
        Spacer(modifier = Modifier.height(16.dp))
        
        OutlinedTextField(
            value = name,
            onValueChange = { name = it },
            label = { Text("Name (Optional)") },
            modifier = Modifier.fillMaxWidth()
        )
        
        Spacer(modifier = Modifier.height(24.dp))
        
        Button(
            onClick = {
                val ts = Instant.now().toString()
                val p = UpdateGuestPayload(
                    id = guestId,
                    name = name.ifBlank { null },
                    programs = GuestPrograms(false, false, false),
                    createdAt = ts,
                    lastVisit = ts
                )
                SyncService.enqueue(APIAction.UPDATE_GUEST, p)
                
                // Navigate to Service Entry (Replace)
                // In Compose we pop then navigate
                navController.popBackStack()
                navController.navigate("service_entry/$guestId/true/false")
            },
            modifier = Modifier.fillMaxWidth()
        ) {
            Text("Register & Continue")
        }
    }
}
