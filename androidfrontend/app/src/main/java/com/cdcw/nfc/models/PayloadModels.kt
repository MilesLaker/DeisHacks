package com.cdcw.nfc.models

import kotlinx.serialization.Serializable
import kotlinx.serialization.json.JsonElement

// Enums
enum class APIAction {
    LOG_SERVICE,
    UPDATE_GUEST,
    REPLACE_CARD,
    CLOTHING_PURCHASE,
    ANONYMOUS_ENTRY,
    GET_BUDGET
}

// Payloads

@Serializable
data class ServiceEntryPayload(
    val guestId: String,
    val services: ServicesDict,
    val timestamp: String
)

@Serializable
data class ServicesDict(
    val shower: Boolean,
    val laundry: Boolean,
    val meals: Int,
    val hygieneKits: Int
)

@Serializable
data class AnonymousEntryPayload(
    val meals: Int,
    val timestamp: String
)

@Serializable
data class UpdateGuestPayload(
    val id: String,
    val name: String?,
    val programs: GuestPrograms,
    val createdAt: String,
    val lastVisit: String
)

@Serializable
data class GuestPrograms(
    val healthcare: Boolean,
    val seasonalNight: Boolean,
    val sustainability: Boolean
)

@Serializable
data class ClothingPurchasePayload(
    val guestId: String,
    val quantity: Int,
    val timestamp: String
)

@Serializable
data class ReplaceCardPayload(
    val oldId: String,
    val newId: String
)

@Serializable
data class GetBudgetPayload(
    val guestId: String
)

// Wrapper for API Request
@Serializable
data class APIRequest(
    val action: String,
    val payload: JsonElement // Changed from Map<String, Any> to JsonElement
)

// Guest Model (Local)
@Serializable
data class Guest(
    val id: String,
    var name: String?,
    var programs: GuestPrograms,
    var feltonBucks: Int,
    var lastVisit: String
)
