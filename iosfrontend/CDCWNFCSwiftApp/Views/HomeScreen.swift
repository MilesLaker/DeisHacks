import SwiftUI

struct HomeScreen: View {
    @EnvironmentObject var nav: NavigationManager
    @EnvironmentObject var sync: SyncService
    @StateObject var guestStore = GuestStore.shared
    @StateObject var nfc = NFCService.shared
    
    @State private var debugId: String = "guest_001"
    
    // Alert State
    @State private var activeAlert: ActiveAlert?
    @State private var currentGuestId: String = ""
    @State private var currentGuestName: String = ""
    
    enum ActiveAlert: Identifiable {
        case guestFound
        case guestNotFound
        
        var id: Int {
            hashValue
        }
    }
    
    var body: some View {
        VStack(spacing: 20) {
            Text("Service Entry")
                .font(.largeTitle)
                .bold()
                .padding(.top, 40)
            
            // Status Circle / Scan Button
            Button(action: {
                nfc.beginScan()
            }) {
                ZStack {
                    Circle()
                        .fill(nfc.isScanning ? Color.green.opacity(0.1) : Color.blue.opacity(0.1))
                        .frame(width: 200, height: 200)
                        .overlay(
                            Circle().stroke(nfc.isScanning ? Color.green : Color.blue, lineWidth: 3)
                        )
                    
                    if nfc.isScanning {
                        Text("Scanning...")
                            .font(.title2)
                            .foregroundColor(.green)
                            .fontWeight(.semibold)
                    } else {
                        VStack(spacing: 10) {
                            Image(systemName: "wave.3.right")
                                .font(.largeTitle)
                            Text("Tap to Scan")
                                .font(.title2)
                                .fontWeight(.semibold)
                        }
                        .foregroundColor(.blue)
                    }
                }
            }
            .padding(.vertical, 30)
            
            Text("Hold card near top of phone")
                .foregroundColor(.gray)
            
            Spacer()
            
            // Anonymous Entry
            Button(action: {
                nav.navigate(to: .serviceEntry(guestId: "anonymous", isNew: false, isAnonymous: true))
            }) {
                Text("Anonymous Entry")
                    .font(.headline)
                    .foregroundColor(.white)
                    .frame(maxWidth: .infinity)
                    .padding()
                    .background(Color.gray)
                    .cornerRadius(10)
            }
            .padding(.horizontal)
            
            // --- DEBUG SECTION ---
            #if DEBUG
            VStack(alignment: .leading, spacing: 10) {
                Text("--- DEV DEBUG & SYNC ---")
                    .font(.caption)
                    .foregroundColor(.gray)
                    .frame(maxWidth: .infinity, alignment: .center)
                
                if let last = nfc.lastScannedId {
                    Text("Last NFC: \(last)").font(.caption).foregroundColor(.green)
                }
                if let err = nfc.lastError {
                    Text("NFC Error: \(err)").font(.caption).foregroundColor(.red)
                }
                
                // Debug Input
                HStack {
                    TextField("Enter Test ID", text: $debugId)
                        .textFieldStyle(RoundedBorderTextFieldStyle())
                    
                    Button("Simulate") {
                        handleScan(id: debugId)
                    }
                    .buttonStyle(.borderedProminent)
                }
                
                // Sync Info
                HStack {
                    Text("Queue: \(sync.queue.count)")
                    Spacer()
                    Text("Syncing: \(sync.isSyncing ? "YES" : "NO")")
                    if let err = sync.queue.first?.lastError {
                        Spacer()
                        Text("Err: \(err.prefix(10))...").foregroundColor(.red).font(.caption)
                    }
                }
                .padding(8)
                .background(Color.gray.opacity(0.1))
                .cornerRadius(5)
                
                HStack {
                    Button("Sync Now") {
                        sync.processQueue()
                    }
                    .buttonStyle(.bordered)
                    
                    Button("Rand Tap") {
                        let r = "guest_" + String(Int.random(in: 100...999))
                        handleScan(id: r)
                    }
                    .buttonStyle(.bordered)
                }
            }
            .padding()
            .background(Color.yellow.opacity(0.1))
            .cornerRadius(10)
            .padding(.horizontal)
            #endif
        }
        .padding(.bottom)
        .onAppear {
            nfc.onScan = { id in
                self.handleScan(id: id)
            }
        }
        .alert(item: $activeAlert) { alertType in
            switch alertType {
            case .guestFound:
                return Alert(
                    title: Text("Guest Found"),
                    message: Text("Name: \(currentGuestName.isEmpty ? "Unknown" : currentGuestName)"),
                    primaryButton: .default(Text("Service Entry")) {
                        nav.navigate(to: .serviceEntry(guestId: currentGuestId, isNew: false, isAnonymous: false))
                    },
                    secondaryButton: .default(Text("Clothing Store")) {
                        nav.navigate(to: .clothing(guestId: currentGuestId))
                    }
                )
            case .guestNotFound:
                return Alert(
                    title: Text("Guest Not Found"),
                    message: Text("ID: \(currentGuestId)"),
                    primaryButton: .default(Text("New User Setup")) {
                        nav.navigate(to: .newUserSetup(guestId: currentGuestId))
                    },
                    secondaryButton: .destructive(Text("Replace Card")) {
                        nav.navigate(to: .searchGuest(newCardId: currentGuestId))
                    }
                )
            }
        }
    }
    
    func handleScan(id: String) {
        print("[NFC] Scanned: \(id)")
        currentGuestId = id
        
        if let guest = guestStore.getGuest(id: id) {
            currentGuestName = guest.name ?? ""
            activeAlert = .guestFound
        } else {
            activeAlert = .guestNotFound
        }
    }
}
