export interface Guest {
    id: string; // Unique ID from NFC
    name?: string; // Optional
    createdAt: string; // ISO Date
    programs: {
        healthcare: boolean;
        seasonalNight: boolean;
        sustainability: boolean;
    };
    feltonBucks?: number; // Budget
    lastVisit: string;
}

export interface ServiceLog {
    guestId: string;
    timestamp: string;
    services: {
        shower: boolean;
        laundry: boolean;
        meals: number;
        hygieneKits: number;
        clothing: number;
    };
}
