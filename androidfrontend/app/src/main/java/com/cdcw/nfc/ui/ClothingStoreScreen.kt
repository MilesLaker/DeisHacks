package com.cdcw.nfc.ui

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.navigation.NavController
import com.cdcw.nfc.models.APIAction
import com.cdcw.nfc.models.ClothingPurchasePayload
import com.cdcw.nfc.network.SyncService
import java.time.Instant

@Composable
fun ClothingStoreScreen(navController: NavController, guestId: String) {
    var budget by remember { mutableIntStateOf(0) }
    var quantity by remember { mutableIntStateOf(0) }
    var isLoading by remember { mutableStateOf(true) }
    
    LaunchedEffect(guestId) {
        SyncService.fetchBudget(guestId) { b ->
            budget = b ?: 0
            isLoading = false
        }
    }
    
    Column(
        modifier = Modifier.padding(16.dp),
        horizontalAlignment = Alignment.CenterHorizontally
    ) {
        Text("Clothing Store", style = MaterialTheme.typography.headlineMedium)
        Text("Guest: $guestId")
        
        Spacer(modifier = Modifier.height(30.dp))
        
        Card(
            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.primaryContainer),
            modifier = Modifier.fillMaxWidth()
        ) {
            Column(
                modifier = Modifier.padding(20.dp),
                horizontalAlignment = Alignment.CenterHorizontally
            ) {
                Text("Current Budget")
                if (isLoading) {
                    CircularProgressIndicator()
                } else {
                    Text("$budget Felton Bucks", style = MaterialTheme.typography.displaySmall)
                }
            }
        }
        
        Spacer(modifier = Modifier.height(30.dp))
        
        if (!isLoading && budget > 0) {
            Text("Select Items")
            Row(
                verticalAlignment = Alignment.CenterVertically,
                horizontalArrangement = Arrangement.Center,
                modifier = Modifier.fillMaxWidth().padding(20.dp)
            ) {
                Button(onClick = { if (quantity > 0) quantity-- }) { Text("-") }
                Text(
                    "$quantity",
                    style = MaterialTheme.typography.displayMedium,
                    modifier = Modifier.padding(horizontal = 20.dp)
                )
                Button(onClick = { if (quantity < budget) quantity++ }) { Text("+") }
            }
            
            Button(
                onClick = {
                    if (quantity > 0) {
                        val ts = Instant.now().toString()
                        val p = ClothingPurchasePayload(guestId, quantity, ts)
                        SyncService.enqueue(APIAction.CLOTHING_PURCHASE, p)
                        navController.popBackStack()
                    }
                },
                enabled = quantity > 0,
                modifier = Modifier.fillMaxWidth()
            ) {
                Text("Confirm Purchase ($quantity)")
            }
        } else if (!isLoading) {
            Text("No budget available", color = Color.Red)
            Spacer(modifier = Modifier.height(20.dp))
            Button(onClick = { navController.popBackStack() }) {
                Text("Back")
            }
        }
    }
}
