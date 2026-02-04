import SwiftUI

struct SearchGuestScreen: View {
    let newCardId: String? // If present, we are in "Replace Card" mode
    
    @EnvironmentObject var nav: NavigationManager
    @EnvironmentObject var sync: SyncService
    @StateObject var guestStore = GuestStore.shared
    
    @State private var query = ""
    @State private var showConfirm = false
    @State private var selectedGuest: Guest?
    
    // Filter
    var filteredGuests: [Guest] {
        let q = query.lowercased()
        return guestStore.guests.values.filter { guest in
            if q.isEmpty { return true }
            return guest.id.lowercased().contains(q) ||
                   (guest.name?.lowercased().contains(q) ?? false)
        }.sorted { $0.id < $1.id }
    }
    
    var body: some View {
        VStack {
            if let newId = newCardId {
                VStack {
                    Text("Replacing Card for NFC ID: \(newId)")
                        .bold().foregroundColor(.orange)
                    Text("Search for the guest's existing profile below.")
                        .font(.caption)
                }
                .padding()
                .background(Color.orange.opacity(0.1))
                .cornerRadius(8)
                .padding(.horizontal)
            }
            
            TextField("Search Name, ID...", text: $query)
                .textFieldStyle(RoundedBorderTextFieldStyle())
                .padding()
            
            List(filteredGuests) { guest in
                Button(action: {
                    if newCardId != nil {
                        selectedGuest = guest
                        showConfirm = true
                    } else {
                        // Just View
                        nav.navigate(to: .serviceEntry(guestId: guest.id, isNew: false, isAnonymous: false))
                    }
                }) {
                    VStack(alignment: .leading) {
                        Text(guest.name ?? "Unnamed")
                            .font(.headline)
                        Text("ID: \(guest.id)")
                            .font(.caption).foregroundColor(.gray)
                    }
                }
            }
            
            if newCardId != nil && filteredGuests.isEmpty && !query.isEmpty {
                Button("Not Found? Create New Guest") {
                    nav.replace(with: .newUserSetup(guestId: newCardId!))
                }
                .padding()
            }
        }
        .navigationTitle("Find Guest")
        .alert("Confirm Replacement", isPresented: $showConfirm) {
            Button("Link & Continue", role: .destructive) {
                if let oldG = selectedGuest, let newId = newCardId {
                    performReplacement(oldGuest: oldG, newId: newId)
                }
            }
            Button("Cancel", role: .cancel) {}
        } message: {
            if let g = selectedGuest {
                Text("Link new card (\(newCardId ?? "")) to guest \(g.name ?? g.id)?")
            }
        }
    }
    
    func performReplacement(oldGuest: Guest, newId: String) {
        // Queue Backend
        sync.addToQueue(action: .replaceCard, payload: ReplaceCardPayload(oldId: oldGuest.id, newId: newId))
        
        // Local Logic ? We don't have explicit Local Replace in RN slice actually?
        // RN slice just queues it and navigates.
        // But for UI consistency, we should clone the guest locally with new ID?
        // For strict parity: RN SearchGuestScreen.tsx just dispatched replaceCard and Navigated.
        
        nav.replace(with: .serviceEntry(guestId: newId, isNew: false, isAnonymous: false))
    }
}
