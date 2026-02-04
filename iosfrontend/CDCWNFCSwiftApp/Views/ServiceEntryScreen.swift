import SwiftUI

struct ServiceEntryScreen: View {
    let guestId: String
    let isNew: Bool
    let isAnonymous: Bool // "Anonymous" mode
    
    @EnvironmentObject var nav: NavigationManager
    @EnvironmentObject var sync: SyncService
    @StateObject var guestStore = GuestStore.shared
    
    // Form State
    @State private var shower = false
    @State private var laundry = false
    @State private var meals = 0
    @State private var hygiene = 0
    
    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 20) {
                // Header
                VStack(alignment: .leading) {
                    Text("Guest ID: \(guestId)")
                        .font(.title2)
                        .bold()
                    if isNew {
                        Text("New Guest")
                            .font(.caption)
                            .padding(4)
                            .background(Color.green.opacity(0.2))
                            .foregroundColor(.green)
                            .cornerRadius(4)
                    }
                }
                .padding(.bottom)
                
                // Facilities
                if !isAnonymous {
                    Group {
                        Text("Facilities").font(.headline)
                        ToggleRow(title: "Shower", isOn: $shower)
                        ToggleRow(title: "Laundry", isOn: $laundry)
                    }
                    .padding()
                    .background(Color.white)
                    .cornerRadius(10)
                    .shadow(radius: 1)
                }
                
                // Distributions
                Group {
                    Text("Distributions").font(.headline)
                    CounterRow(title: "Meals", count: $meals)
                    if !isAnonymous {
                         CounterRow(title: "Hygiene Kits", count: $hygiene)
                    }
                }
                .padding()
                .background(Color.white)
                .cornerRadius(10)
                .shadow(radius: 1)
                
                // Submit
                Button(action: submit) {
                    Text("Submit Entry")
                        .font(.headline)
                        .foregroundColor(.white)
                        .frame(maxWidth: .infinity)
                        .padding()
                        .background(Color.green)
                        .cornerRadius(10)
                }
                .padding(.top)
            }
            .padding()
        }
        .background(Color(UIColor.systemGroupedBackground))
        .navigationTitle(isAnonymous ? "Anonymous" : "Service Entry")
        .navigationBarTitleDisplayMode(.inline)
    }
    
    func submit() {
        let ts = Date().iso8601
        
        if isAnonymous {
            let payload = AnonymousEntryPayload(meals: meals, timestamp: ts)
            sync.addToQueue(action: .anonymousEntry, payload: payload)
        } else {
            let services = ServicesDict(shower: shower, laundry: laundry, meals: meals, hygieneKits: hygiene)
            let payload = ServiceEntryPayload(guestId: guestId, services: services, timestamp: ts)
            sync.addToQueue(action: .logService, payload: payload)
            
            // Local Update
            if var g = guestStore.guests[guestId] {
                g.lastVisit = ts
                guestStore.saveGuest(g)
            }
        }
        
        nav.goBack()
    }
}

// MARK: - Components

struct ToggleRow: View {
    let title: String
    @Binding var isOn: Bool
    
    var body: some View {
        HStack {
            Text(title)
            Spacer()
            // Custom Toggle Look
            Button(action: { isOn.toggle() }) {
                Text(isOn ? "YES" : "NO")
                    .bold()
                    .foregroundColor(isOn ? .blue : .gray)
                    .padding(.horizontal, 10)
                    .padding(.vertical, 5)
                    .background(isOn ? Color.blue.opacity(0.1) : Color.clear)
                    .cornerRadius(5)
            }
        }
        .padding(.vertical, 8)
        .overlay(Divider(), alignment: .bottom)
    }
}

struct CounterRow: View {
    let title: String
    @Binding var count: Int
    
    var body: some View {
        HStack {
            Text(title)
            Spacer()
            HStack(spacing: 20) {
                Button(action: { if count > 0 { count -= 1 } }) {
                    Image(systemName: "minus.circle.fill")
                        .foregroundColor(.gray)
                        .font(.title2)
                }
                Text("\(count)")
                    .font(.title3).bold()
                    .frame(minWidth: 20)
                Button(action: { count += 1 }) {
                    Image(systemName: "plus.circle.fill")
                        .foregroundColor(.blue)
                        .font(.title2)
                }
            }
        }
        .padding(.vertical, 10)
        .overlay(Divider(), alignment: .bottom)
    }
}

extension Date {
    var iso8601: String {
        let formatter = ISO8601DateFormatter()
        return formatter.string(from: self)
    }
}
