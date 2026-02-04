package com.cdcw.nfc.network

import android.content.Context
import android.content.SharedPreferences
import com.cdcw.nfc.models.*
import kotlinx.serialization.encodeToString
import kotlinx.serialization.json.Json
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import okhttp3.*
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.RequestBody.Companion.toRequestBody
import java.io.IOException
import java.util.UUID
import java.util.concurrent.Executors

// Singleton for simplicity
object SyncService {
    private const val API_URL = "https://script.google.com/macros/s/AKfycbxEdsw5CquVDGOb2NDjxok0-zgQZVfV1Ej5AyP-n_6UmaqJpb-ap7OOothrwr_Rq_7U/exec"
    private const val PREFS_NAME = "offline_queue"
    private const val QUEUE_KEY = "queue_data"

    private lateinit var prefs: SharedPreferences
    private val client = OkHttpClient()
    
    // Changed to internal (or public) to be accessible by inline functions
    val json = Json { ignoreUnknownKeys = true }
    private val executor = Executors.newSingleThreadExecutor()

    // Changed to internal to be accessible by inline functions
    val _queueState = MutableStateFlow<List<SyncItem>>(emptyList())
    val queueState: StateFlow<List<SyncItem>> = _queueState

    private val _isSyncing = MutableStateFlow(false)
    val isSyncing: StateFlow<Boolean> = _isSyncing

    data class SyncItem(
        val id: String = UUID.randomUUID().toString(),
        val action: String,
        val payloadJson: String,
        var retryCount: Int = 0,
        var lastError: String? = null
    )

    fun init(context: Context) {
        prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        loadQueue()
    }

    fun addToQueue(action: APIAction, payload: Any) {
        // Implementation replaced by generic enqueue below
    }
    
    // Generic Adder
    inline fun <reified T> enqueue(action: APIAction, payload: T) {
        val pString = json.encodeToString(payload)
        val item = SyncItem(
            action = action.name,
            payloadJson = pString
        )
        val current = _queueState.value.toMutableList()
        current.add(item)
        _queueState.value = current
        saveQueue()
        processQueue()
    }

    private fun loadQueue() {
        val str = prefs.getString(QUEUE_KEY, null)
        if (str != null) {
            try {
                // Simple list deserialization skipped for brevity
            } catch (e: Exception) {
                // Reset on error
            }
        }
    }

    // Changed to internal for inline access
    fun saveQueue() {
        // serialize _queueState.value and save to prefs
    }

    fun processQueue() {
        if (_isSyncing.value || _queueState.value.isEmpty()) return

        _isSyncing.value = true
        executor.execute {
            val list = _queueState.value.toMutableList()
            if (list.isNotEmpty()) {
                val item = list[0]
                try {
                    val success = sendNetworkRequest(item)
                    if (success) {
                        list.removeAt(0)
                        _queueState.value = list
                        saveQueue()
                        // Recurse
                        _isSyncing.value = false
                        processQueue()
                    } else {
                        // Retry later
                        _isSyncing.value = false
                    }
                } catch (e: Exception) {
                    item.retryCount++
                    item.lastError = e.message
                    _isSyncing.value = false
                }
            } else {
                _isSyncing.value = false
            }
        }
    }

    private fun sendNetworkRequest(item: SyncItem): Boolean {
        // Manual JSON construction for simplicity
        val bodyJson = """
            {
                "action": "${item.action}",
                "payload": ${item.payloadJson}
            }
        """.trimIndent()
        
        val request = Request.Builder()
            .url(API_URL)
            .post(bodyJson.toRequestBody("application/json".toMediaType()))
            .build()

        client.newCall(request).execute().use { response ->
            return response.isSuccessful
        }
    }

    // Immediate Fetch
    fun fetchBudget(guestId: String, callback: (Int?) -> Unit) {
        executor.execute {
            val bodyJson = """
                {
                    "action": "GET_BUDGET",
                    "payload": { "guestId": "$guestId" }
                }
            """.trimIndent()

            val request = Request.Builder()
                .url(API_URL)
                .post(bodyJson.toRequestBody("application/json".toMediaType()))
                .build()

            try {
                client.newCall(request).execute().use { response ->
                    if (response.isSuccessful) {
                        val respStr = response.body?.string()
                        if (respStr != null) {
                            val match = "\"budget\":(\\d+)".toRegex().find(respStr)
                            val budget = match?.groupValues?.get(1)?.toIntOrNull()
                            callback(budget)
                        } else {
                            callback(null)
                        }
                    } else {
                        callback(null)
                    }
                }
            } catch (e: Exception) {
                callback(null)
            }
        }
    }
}
