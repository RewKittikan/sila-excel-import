export interface HazardousImportRow {
  establishment: {
    ownerName: string | null;
    establishmentName: string | null;
    businessType: string | null;
    phone: string | null;
    addressNo: string | null;
    moo: string | null;
    subdistrict: string | null;
  };

  license: {
    feeAmount: number | null;
    areaSqm: number | null;
    expiryDate: string | null;
  };

  payments: {
    year: number;
    isPaid: boolean;
  }[];
}