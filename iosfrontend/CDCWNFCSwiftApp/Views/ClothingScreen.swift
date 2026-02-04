import SwiftUI

struct ClothingScreen: View {
    let guestId: String
    
    @EnvironmentObject var nav: NavigationManager
    @EnvironmentObject var sync: SyncService
    @StateObject var guestStore = GuestStore.shared
    
    @State private var quantity = 0
    @State private var refreshing = false
    @State private var errorMessage: String?
    
    var guest: Guest? {
        guestStore.guests[guestId]
    }
    
    var budget: Int {
        guest?.feltonBucks ?? 0
    }
    
    var body: some View {
        ScrollView {
            VStack(spacing: 20) {
                // Header Card
                VStack(spacing: 10) {
                    Text("Clothing Store")
                        .font(.title2)
                        .bold()
                    
                    Text("Guest: \(guest?.name ?? guestId)")
                        .foregroundColor(.gray)
                    
                    Text("Current Budget")
                        .font(.caption).foregroundColor(.blue)
                        .padding(.top, 10)
                    
                    Text("\(budget) Felton Bucks")
                        .font(.system(size: 32, weight: .bold))
                        .foregroundColor(.blue)
                    
                    Button(action: refreshBudget) {
                        if refreshing {
                            ProgressView()
                        } else {
                            Text("Refresh Budget")
                                .font(.footnote)
                                .underline()
                        }
                    }
                    .padding(.bottom, 10)
                }
                .frame(maxWidth: .infinity)
                .background(Color.blue.opacity(0.1))
                .cornerRadius(10)
                
                if let err = errorMessage {
                    Text(err).foregroundColor(.red).font(.caption)
                }
                
                // Controls
                if budget > 0 {
                    VStack(spacing: 15) {
                        Text("Select Items").font(.headline)
                        
                        HStack(spacing: 30) {
                            Button(action: { if quantity > 0 { quantity -= 1 } }) {
                                Circle().fill(Color.gray.opacity(0.3)).frame(width: 50, height: 50)
                                    .overlay(Image(systemName: "minus").font(.title))
                            }
                            
                            Text("\(quantity)")
                                .font(.system(size: 40, weight: .bold))
                                .frame(minWidth: 50)
                            
                            Button(action: {
                                if quantity < budget { quantity += 1 }
                                else { errorMessage = "Limit reached" }
                            }) {
                                Circle().fill(Color.gray.opacity(0.3)).frame(width: 50, height: 50)
                                    .overlay(Image(systemName: "plus").font(.title))
                            }
                        }
                    }
                    .padding()
                } else {
                    Text("No clothing budget available.")
                        .foregroundColor(.red)
                        .padding()
                }
                
                // Submit
                Button(action: confirmPurchase) {
                    Text("Confirm Purchase (\(quantity))")
                        .font(.headline)
                        .foregroundColor(.white)
                        .frame(maxWidth: .infinity)
                        .padding()
                        .background((quantity > 0 && quantity <= budget) ? Color.green : Color.gray)
                        .cornerRadius(10)
                }
                .disabled(quantity == 0 || quantity > budget)
                .padding(.top, 30)
            }
            .padding()
        }
        .onAppear {
            refreshBudget()
        }
        .navigationTitle("Clothing")
        .navigationBarTitleDisplayMode(.inline)
    }
    
    func refreshBudget() {
        refreshing = true
        errorMessage = nil
        Task {
            do {
                let res = try await NetworkClient.shared.send(action: .getBudget, payload: GetBudgetPayload(guestId: guestId))
                await MainActor.run {
                    refreshing = false
                    if let newB = res.budget {
                        guestStore.updateBudget(id: guestId, newBudget: newB)
                    } else if let msg = res.message {
                        errorMessage = msg
                    }
                }
            } catch {
                await MainActor.run {
                    refreshing = false
                    errorMessage = "Network Error"
                }
            }
        }
    }
    
    func confirmPurchase() {
        guard quantity > 0 else { return }
        let ts = Date().iso8601
        
        let payload = ClothingPurchasePayload(guestId: guestId, quantity: quantity, timestamp: ts)
        
        // Optimistic Update
        let newBalance = budget - quantity
        guestStore.updateBudget(id: guestId, newBudget: newBalance)
        
        // Queue
        sync.addToQueue(action: .clothingPurchase, payload: payload)
        
        // Exit
        nav.goBack()
    }
}
