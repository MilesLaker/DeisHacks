package com.cdcw.nfc.ui

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import androidx.navigation.NavController
import com.cdcw.nfc.models.APIAction
import com.cdcw.nfc.models.AnonymousEntryPayload
import com.cdcw.nfc.models.ServiceEntryPayload
import com.cdcw.nfc.models.ServicesDict
import com.cdcw.nfc.network.SyncService
import java.time.Instant

@Composable
fun ServiceEntryScreen(
    navController: NavController,
    guestId: String,
    isNew: Boolean,
    isAnonymous: Boolean
) {
    var shower by remember { mutableStateOf(false) }
    var laundry by remember { mutableStateOf(false) }
    var meals by remember { mutableIntStateOf(0) }
    var hygiene by remember { mutableIntStateOf(0) }

    Column(modifier = Modifier.padding(16.dp)) {
        Text("Service Entry", style = MaterialTheme.typography.headlineMedium)
        Text("Guest: $guestId")
        if (isNew) {
            Text("NEW GUEST", color = MaterialTheme.colorScheme.primary)
        }
        
        Spacer(modifier = Modifier.height(20.dp))
        
        if (!isAnonymous) {
            Row(verticalAlignment = Alignment.CenterVertically) {
                Text("Shower", modifier = Modifier.weight(1f))
                Switch(checked = shower, onCheckedChange = { shower = it })
            }
            Divider()
            Row(verticalAlignment = Alignment.CenterVertically) {
                Text("Laundry", modifier = Modifier.weight(1f))
                Switch(checked = laundry, onCheckedChange = { laundry = it })
            }
        }
        
        Divider()
        CounterRow("Meals", meals) { meals = it }
        if (!isAnonymous) {
            CounterRow("Hygiene Kits", hygiene) { hygiene = it }
        }
        
        Spacer(modifier = Modifier.height(30.dp))
        
        Button(
            onClick = {
                val ts = Instant.now().toString()
                if (isAnonymous) {
                    val p = AnonymousEntryPayload(meals, ts)
                    SyncService.enqueue(APIAction.ANONYMOUS_ENTRY, p)
                } else {
                    val s = ServicesDict(shower, laundry, meals, hygiene)
                    val p = ServiceEntryPayload(guestId, s, ts)
                    SyncService.enqueue(APIAction.LOG_SERVICE, p)
                }
                navController.popBackStack()
            },
            modifier = Modifier.fillMaxWidth()
        ) {
            Text("Submit Entry")
        }
    }
}

@Composable
fun CounterRow(label: String, value: Int, onValueChange: (Int) -> Unit) {
    Row(
        modifier = Modifier.fillMaxWidth().padding(vertical = 8.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        Text(label, modifier = Modifier.weight(1f))
        IconButton(onClick = { if (value > 0) onValueChange(value - 1) }) {
            Text("-", style = MaterialTheme.typography.headlineSmall)
        }
        Text("$value", style = MaterialTheme.typography.titleMedium)
        IconButton(onClick = { onValueChange(value + 1) }) {
            Text("+", style = MaterialTheme.typography.headlineSmall)
        }
    }
}
