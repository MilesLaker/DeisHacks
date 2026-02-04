import Foundation

enum NetworkError: Error {
    case invalidURL
    case serverError(String)
    case decodingError
    case unknown
}

class NetworkClient {
    static let shared = NetworkClient()
    
    private init() {}
    
    func send<T: Codable>(action: APIAction, payload: T, eventId: String? = nil) async throws -> APIResponse {
        var request = URLRequest(url: AppConfig.apiURL)
        request.httpMethod = "POST"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        
        // Construct body manually or via wrapper. 
        // We simply construct a dictionary to ensure correct JSON structure for 'payload'
        // Since AnyCodable is tricky, let's use a simpler Dict approach for outgoing.
        
        let bodyDict: [String: Any] = [
            "action": action.rawValue,
            "payload": try! JSONSerialization.jsonObject(with: try! JSONEncoder().encode(payload))
        ]
        
        // If eventId exists (for dedup), inject it into payload if possible, or wrapper?
        // Backend `doPost` expects `action` and `payload`. 
        // Backend `logService` checks `payload.eventId`.
        // So we must inject eventId INTO the payload object if it's not there.
        
        var finalPayload = bodyDict["payload"] as! [String: Any]
        if let eid = eventId {
            finalPayload["eventId"] = eid
        }
        
        let finalBody: [String: Any] = [
            "action": action.rawValue,
            "payload": finalPayload
        ]
        
        let jsonData = try JSONSerialization.data(withJSONObject: finalBody)
        request.httpBody = jsonData
        
        let (data, response) = try await URLSession.shared.data(for: request)
        
        guard let httpResponse = response as? HTTPURLResponse, httpResponse.statusCode == 200 else {
             // Handle Google Redirects? usually UrlSession follows them.
             // If Google Script returns HTML error, decoding will fail.
             throw NetworkError.serverError("Non-200 Status")
        }
        
        do {
            let apiRes = try JSONDecoder().decode(APIResponse.self, from: data)
            return apiRes
        } catch {
            if let str = String(data: data, encoding: .utf8) {
                print("[Network] Decoding Error. Raw: \(str)")
            }
            throw NetworkError.decodingError
        }
    }
}
