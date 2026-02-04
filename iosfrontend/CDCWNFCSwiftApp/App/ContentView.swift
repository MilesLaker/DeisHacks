import SwiftUI

enum AppRoute: Hashable {
    case serviceEntry(guestId: String, isNew: Bool, isAnonymous: Bool)
    case newUserSetup(guestId: String)
    case clothing(guestId: String)
    case searchGuest(newCardId: String?)
}

struct ContentView: View {
    @EnvironmentObject var nav: NavigationManager
    @EnvironmentObject var sync: SyncService
    
    var body: some View {
        NavigationStack(path: $nav.path) {
            HomeScreen()
                .navigationDestination(for: AppRoute.self) { route in
                    switch route {
                    case .serviceEntry(let guestId, let isNew, let isAnonymous):
                        ServiceEntryScreen(guestId: guestId, isNew: isNew, isAnonymous: isAnonymous)
                    case .newUserSetup(let guestId):
                        NewUserSetupScreen(guestId: guestId)
                    case .clothing(let guestId):
                        ClothingScreen(guestId: guestId)
                    case .searchGuest(let newCardId):
                        SearchGuestScreen(newCardId: newCardId)
                    }
                }
        }
    }
}
