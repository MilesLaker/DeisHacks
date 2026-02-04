import SwiftUI
import Combine

class NavigationManager: ObservableObject {
    @Published var path = NavigationPath()
    
    func navigate(to route: AppRoute) {
        path.append(route)
    }
    
    func popToRoot() {
        path = NavigationPath()
    }
    
    func goBack() {
        if !path.isEmpty {
            path.removeLast()
        }
    }
    
    // Helper to replace root (simulated by pop then push technically, but NavigationStack is stack based)
    // For "Replace" logic (NewUser -> ServiceEntry), we can manipulate the path directly.
    func replace(with route: AppRoute) {
        if !path.isEmpty {
            path.removeLast()
        }
        path.append(route)
    }
}
