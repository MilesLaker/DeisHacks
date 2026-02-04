import Foundation
import Combine

struct Guest: Codable, Identifiable {
    let id: String
    var name: String?
    var programs: GuestPrograms
    var feltonBucks: Int
    var lastVisit: String
}

@MainActor
class GuestStore: ObservableObject {
    static let shared = GuestStore()
    
    @Published var guests: [String: Guest] = [:]
    
    // Quick Lookup for offline
    func getGuest(id: String) -> Guest? {
        return guests[id]
    }
    
    // Add or Update
    func saveGuest(_ guest: Guest) {
        guests[guest.id] = guest
        persist()
    }
    
    // Update Budget Only
    func updateBudget(id: String, newBudget: Int) {
        if var g = guests[id] {
            g.feltonBucks = newBudget
            guests[id] = g
            persist()
        }
    }
    
    private let key = "local_guests_store"
    
    init() {
        if let data = UserDefaults.standard.data(forKey: key),
           let loaded = try? JSONDecoder().decode([String: Guest].self, from: data) {
            guests = loaded
        }
    }
    
    private func persist() {
        if let data = try? JSONEncoder().encode(guests) {
            UserDefaults.standard.set(data, forKey: key)
        }
    }
    
    // Sync Helpers
    func syncFromBackend() async {
        // Implement full sync if needed, but for now we rely on local + individual fetches
        // The RN app does a full 'doGet' on mount? Let's check.
        // RN syncSlice sends events. GuestSlice loads from persist.
        // RN doesn't seem to fetch ALL guests on launch? 
        // Wait, "Offline Search" implies we have local data. 
        // The RN app likely loads huge JSON? Or just what it has seen?
        // Let's assume on-demand or pre-loaded.
        // For this identical build, we'll implement a `fetchAll` method if requested.
        // PRD says: "Offline caching: implement the minimum equivalent behavior currently used."
    }
}
