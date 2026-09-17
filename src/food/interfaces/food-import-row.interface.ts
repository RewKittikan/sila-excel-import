export type PaymentStatus = 'PAID' | 'EMPTY' | 'NOTE' | 'UNKNOWN';

export interface FoodPayment {
  year: number;
  status: PaymentStatus;
  rawValue: string | null;
}

export type EstablishmentStatus = 'ACTIVE' | 'CLOSED';

export interface FoodImportRow {
  establishment: {
    ownerName: string | null;
    establishmentName: string | null;
    businessType: string | null;
    phone: string | null;
    addressNo: string | null;
    moo: string | null;
    subdistrict: string | null;
    status: EstablishmentStatus;
  };
  license: {
    feeAmount: number | null;
    areaSqm: number | null;
    expiryDay: number | null;
    expiryMonth: number | null;
  };
  payments: FoodPayment[];
}
