import { Worksheet } from 'exceljs';
import { HazardousImportRow } from '../interfaces/hazardous-import-row.interface';

export class EstablishmentParser {
  parse(worksheet: Worksheet): HazardousImportRow[] {
    const results: HazardousImportRow[] = [];

    worksheet.eachRow((row) => {
      const sequence = row.getCell(1).value;

      if (typeof sequence !== 'number') {
        return;
      }

      const ownerName = this.toText(row.getCell(2).value);
      const establishmentName = this.toText(row.getCell(3).value);

      if (!ownerName && !establishmentName) {
        return;
      }

      results.push({
        establishment: {
          ownerName,
          establishmentName,
          businessType: this.toText(row.getCell(4).value),
          addressNo: this.toText(row.getCell(5).value),
          moo: this.toText(row.getCell(6).value),
          subdistrict: this.toText(row.getCell(7).value),
          phone: this.toText(row.getCell(8).value),
        },

        license: {
          feeAmount: this.toNumber(row.getCell(9).value),
          areaSqm: this.toNumber(row.getCell(10).value),
          expiryDate: this.toText(row.getCell(11).value),
        },

        payments: this.parsePayments(row),
      });
    });

    return results;
  }

  private parsePayments(row: any) {
    const payments = [];

    // T = column 20 = 2567
    // W = 23 = 2570
    for (let col = 20, year = 2567; col <= 23; col++, year++) {
      const value = row.getCell(col).value;

      payments.push({
        year,
        isPaid: value === '/' || value === true,
      });
    }

    return payments;
  }

  private toText(value: unknown): string | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    return String(value).trim();
  }

  private toNumber(value: unknown): number | null {
    if (typeof value === 'number') {
      return value;
    }

    if (value === null || value === undefined || value === '') {
      return null;
    }

    const parsed = Number(value);

    return Number.isNaN(parsed) ? null : parsed;
  }
}
