package com.cdcw.nfc.navigation

import androidx.compose.runtime.Composable
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.rememberNavController
import com.cdcw.nfc.ui.*
import com.cdcw.nfc.nfc.NFCService

object Routes {
    const val HOME = "home"
    const val SERVICE_ENTRY = "service_entry/{guestId}/{isNew}/{isAnonymous}"
    const val NEW_GUEST = "new_guest/{guestId}"
    const val CLOTHING = "clothing/{guestId}"
    const val SEARCH = "search"
    
    fun serviceEntry(guestId: String, isNew: Boolean, isAnonymous: Boolean) =
        "service_entry/$guestId/$isNew/$isAnonymous"
        
    fun clothing(guestId: String) = "clothing/$guestId"
    fun newGuest(guestId: String) = "new_guest/$guestId"
}



@Composable
fun AppNavigation(nfcService: NFCService) {
    val navController = rememberNavController()

    NavHost(navController = navController, startDestination = Routes.HOME) {
        composable(Routes.HOME) {
            ScanScreen(navController, nfcService)
        }
        composable(Routes.SERVICE_ENTRY) { backStackEntry ->
            val guestId = backStackEntry.arguments?.getString("guestId") ?: ""
            val isNew = backStackEntry.arguments?.getString("isNew")?.toBoolean() ?: false
            val isAnon = backStackEntry.arguments?.getString("isAnonymous")?.toBoolean() ?: false
            ServiceEntryScreen(navController, guestId, isNew, isAnon)
        }
        composable(Routes.CLOTHING) { entry ->
            val guestId = entry.arguments?.getString("guestId") ?: ""
            ClothingStoreScreen(navController, guestId)
        }
        composable(Routes.NEW_GUEST) { entry ->
             val guestId = entry.arguments?.getString("guestId") ?: ""
             NewGuestScreen(navController, guestId)
        }
    }
}
