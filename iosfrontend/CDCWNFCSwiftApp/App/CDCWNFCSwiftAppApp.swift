import SwiftUI

@main
struct CDCWNFCSwiftAppApp: App {
    @StateObject private var nav = NavigationManager()
    @StateObject private var sync = SyncService.shared
    
    var body: some Scene {
        WindowGroup {
            ContentView()
                .environmentObject(nav)
                .environmentObject(sync)
        }
    }
}
