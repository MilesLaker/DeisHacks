import SwiftUI

struct NewUserSetupScreen: View {
    let guestId: String
    @EnvironmentObject var nav: NavigationManager
    @EnvironmentObject var sync: SyncService
    @StateObject var guestStore = GuestStore.shared
    
    @State private var name = ""
    
    var body: some View {
        Form {
            Section(header: Text("New Guest Info")) {
                Text("NFC ID: \(guestId)").foregroundColor(.gray)
                TextField("Name (Optional)", text: $name)
            }
            
            Section {
                Button("Save & Continue") {
                    save()
                }
                .frame(maxWidth: .infinity, alignment: .center)
            }
        }
        .navigationTitle("New Guest")
    }
    
    func save() {
        let ts = Date().iso8601
        
        // 1. Create Guest Object
        let newGuest = Guest(
            id: guestId,
            name: name.isEmpty ? nil : name,
            programs: GuestPrograms(healthcare: false, seasonalNight: false, sustainability: false),
            feltonBucks: 0,
            lastVisit: ts
        )
        
        // Local Save
        guestStore.saveGuest(newGuest)
        
        // Queue Sync
        let payload = UpdateGuestPayload(
            id: guestId,
            name: name.isEmpty ? nil : name,
            programs: newGuest.programs,
            createdAt: ts,
            lastVisit: ts
        )
        
        sync.addToQueue(action: .updateGuest, payload: payload)
        
        // Navigate Replace: Go to ServiceEntry
        nav.replace(with: .serviceEntry(guestId: guestId, isNew: true, isAnonymous: false))
    }
}
