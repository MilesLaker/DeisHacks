import Foundation

// MARK: - Enums
enum APIAction: String, Codable {
    case logService = "LOG_SERVICE"
    case updateGuest = "UPDATE_GUEST"
    case replaceCard = "REPLACE_CARD"
    case clothingPurchase = "CLOTHING_PURCHASE"
    case anonymousEntry = "ANONYMOUS_ENTRY"
    case getBudget = "GET_BUDGET"
}

// MARK: - Payloads

struct ServiceEntryPayload: Codable {
    let guestId: String
    let services: ServicesDict
    let timestamp: String
}

struct ServicesDict: Codable {
    let shower: Bool
    let laundry: Bool
    let meals: Int
    let hygieneKits: Int
}

struct AnonymousEntryPayload: Codable {
    let meals: Int
    let timestamp: String
}

struct UpdateGuestPayload: Codable {
    let id: String
    let name: String?
    let programs: GuestPrograms
    let createdAt: String
    let lastVisit: String
}

struct GuestPrograms: Codable {
    let healthcare: Bool
    let seasonalNight: Bool
    let sustainability: Bool
}

struct ClothingPurchasePayload: Codable {
    let guestId: String
    let quantity: Int
    let timestamp: String
}

struct ReplaceCardPayload: Codable {
    let oldId: String
    let newId: String
}

struct GetBudgetPayload: Codable {
    let guestId: String
}

// MARK: - Wrapper

struct APIRequest: Codable {
    let action: APIAction
    let payload: AnyCodable
}

// Helper to encode mixed types
struct AnyCodable: Codable {
    let value: Any
    
    init<T>(_ value: T?) {
        self.value = value ?? ()
    }
    
    func encode(to encoder: Encoder) throws {
        var container = encoder.singleValueContainer()
        
        switch value {
        case let i as Int: try container.encode(i)
        case let d as Double: try container.encode(d)
        case let s as String: try container.encode(s)
        case let b as Bool: try container.encode(b)
        case let c as ServiceEntryPayload: try container.encode(c)
        case let c as AnonymousEntryPayload: try container.encode(c)
        case let c as UpdateGuestPayload: try container.encode(c)
        case let c as ClothingPurchasePayload: try container.encode(c)
        case let c as ReplaceCardPayload: try container.encode(c)
        case let c as GetBudgetPayload: try container.encode(c)
        case let c as [String: AnyCodable]: try container.encode(c) // Recursive
        default:
            let context = EncodingError.Context(codingPath: container.codingPath, debugDescription: "AnyCodable value cannot be encoded")
            throw EncodingError.invalidValue(value, context)
        }
    }
    
    init(from decoder: Decoder) throws {
        // We only use this for encoding requests, but if needed we can impl decoding
        self.value = "" 
    }
}

// MARK: - Responses

struct APIResponse: Codable {
    let status: String
    let message: String?
    let buildId: String?
    
    // Fields specific to certain actions
    let budget: Int?
    let log: LogResponse?
    
    struct LogResponse: Codable {
        let written: Int?
        let row: Int?
    }
}
