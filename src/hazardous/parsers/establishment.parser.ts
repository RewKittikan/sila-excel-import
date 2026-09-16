import { Row, Worksheet } from 'exceljs';
import {
  EstablishmentStatus,
  HazardousImportRow,
  HazardousPayment,
  PaymentStatus,
} from '../interfaces/hazardous-import-row.interface';

export class EstablishmentParser {
  parse(worksheet: Worksheet): HazardousImportRow[] {
    const results: HazardousImportRow[] = [];

    worksheet.eachRow((row) => {
      const sequence = row.getCell(1).value;

      // รับเฉพาะแถวข้อมูลจริง
      if (typeof sequence !== 'number') {
        return;
      }

      const ownerName = this.toText(row.getCell(2).value);
      const establishmentName = this.toText(row.getCell(3).value);

      // ถ้าไม่มีทั้งผู้ประกอบการและชื่อสถานประกอบการ ให้ข้าม
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
          phone: this.normalizePhone(row.getCell(8).value),
          status: this.getEstablishmentStatus(row),
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

  private parsePayments(row: Row): HazardousPayment[] {
    const payments: HazardousPayment[] = [];

    // L = column 12 = ปี 2559
    // ...
    // W = column 23 = ปี 2570
    for (let column = 12, year = 2559; column <= 23; column++, year++) {
      const rawValue = this.toText(row.getCell(column).value);

      let status: PaymentStatus;

      if (rawValue === '/') {
        status = 'PAID';
      } else if (!rawValue) {
        status = 'EMPTY';
      } else {
        // เช่น "ยกเลิก 66", "เลิกกิจการ63"
        // ตอนนี้ยังไม่ตีความ business logic
        status = 'NOTE';
      }

      payments.push({
        year,
        status,
        rawValue,
      });
    }

    return payments;
  }

  private getEstablishmentStatus(row: Row): EstablishmentStatus {
    return this.isRedRow(row) ? 'CLOSED' : 'ACTIVE';
  }

  private isRedRow(row: Row): boolean {
    let isRed = false;

    row.eachCell({ includeEmpty: true }, (cell) => {
      const fill = cell.fill;

      if (
        fill &&
        fill.type === 'pattern' &&
        fill.fgColor?.argb?.toUpperCase() === 'FFFF9999'
      ) {
        isRed = true;
      }
    });

    return isRed;
  }

  private normalizePhone(value: unknown): string | null {
    const phone = this.toText(value);

    if (!phone) {
      return null;
    }

    return phone.replace(/[\s-]/g, '');
  }

  private toText(value: unknown): string | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    return String(value).trim();
  }

  private toNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    if (typeof value === 'number') {
      return value;
    }

    const parsed = Number(value);

    return Number.isNaN(parsed) ? null : parsed;
  }
}