import Foundation
import Combine

struct SyncItem: Codable, Identifiable {
    let id: String // UUID
    let action: APIAction
    let payloadData: Data // Store as raw data to act as generic container
    let timestamp: Date
    var retryCount: Int
    var lastError: String?
}

@MainActor
class SyncService: ObservableObject {
    static let shared = SyncService()
    
    @Published var queue: [SyncItem] = []
    @Published var isSyncing = false
    
    private let queueKey = "offline_sync_queue"
    
    init() {
        loadQueue()
    }
    
    func addToQueue<T: Codable>(action: APIAction, payload: T) {
        let id = "evt_" + UUID().uuidString
        do {
            let data = try JSONEncoder().encode(payload)
            let item = SyncItem(id: id, action: action, payloadData: data, timestamp: Date(), retryCount: 0)
            queue.append(item)
            saveQueue()
            processQueue()
        } catch {
            print("[Sync] Failed to encode payload: \(error)")
        }
    }
    
    func processQueue() {
        guard !isSyncing, !queue.isEmpty else { return }
        
        isSyncing = true
        
        Task {
            // Process first item
            if let item = queue.first {
                do {
                    print("[Sync] Processing \(item.action.rawValue) ID: \(item.id)")
                    
                    // Decode payload generic logic? 
                    // NetworkClient needs T. We know Action. 
                    // We can decode to Dictionary for generic send.
                    
                    if let dict = try JSONSerialization.jsonObject(with: item.payloadData) as? [String: Any] {
                         // We rely on NetworkClient internal dictionary handling or define a special "GenericPayload"
                         // Let's modify NetworkClient to accept [String:Any] or handle it here.
                         // Actually, let's create a struct wrapper since NetworkClient uses Generics.
                         // Hack: Decode to specific types based on Action.
                        
                        let _ = try await sendItem(item)
                        
                        print("[Sync] Success: \(item.id)")
                        if !queue.isEmpty {
                             queue.removeFirst()
                             saveQueue()
                        }
                    }
                } catch {
                    print("[Sync] Failed: \(error)")
                    if !queue.isEmpty {
                        queue[0].retryCount += 1
                        queue[0].lastError = String(describing: error)
                        saveQueue()
                    }
                }
            }
            isSyncing = false
            
            // Recursive continue if success
            if !queue.isEmpty && queue.first?.retryCount == 0 {
                processQueue()
            }
        }
    }
    
    private func sendItem(_ item: SyncItem) async throws -> APIResponse {
        // Decode based on action
        let decoder = JSONDecoder()
        switch item.action {
        case .logService:
            let p = try decoder.decode(ServiceEntryPayload.self, from: item.payloadData)
            return try await NetworkClient.shared.send(action: item.action, payload: p, eventId: item.id)
        case .anonymousEntry:
             let p = try decoder.decode(AnonymousEntryPayload.self, from: item.payloadData)
             return try await NetworkClient.shared.send(action: item.action, payload: p, eventId: item.id)
        case .updateGuest:
            let p = try decoder.decode(UpdateGuestPayload.self, from: item.payloadData)
            return try await NetworkClient.shared.send(action: item.action, payload: p, eventId: item.id)
        case .clothingPurchase:
             let p = try decoder.decode(ClothingPurchasePayload.self, from: item.payloadData)
             return try await NetworkClient.shared.send(action: item.action, payload: p, eventId: item.id)
        case .replaceCard:
             let p = try decoder.decode(ReplaceCardPayload.self, from: item.payloadData)
             return try await NetworkClient.shared.send(action: item.action, payload: p, eventId: item.id)
        default:
             // GET requests usually don't go to queue, but if they do:
             throw NetworkError.unknown
        }
    }
    
    private func saveQueue() {
        if let data = try? JSONEncoder().encode(queue) {
            UserDefaults.standard.set(data, forKey: queueKey)
        }
    }
    
    private func loadQueue() {
        if let data = UserDefaults.standard.data(forKey: queueKey),
           let loaded = try? JSONDecoder().decode([SyncItem].self, from: data) {
            queue = loaded
        }
    }
}
