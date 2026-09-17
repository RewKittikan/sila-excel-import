import { Row, Worksheet } from 'exceljs';
import {
  EstablishmentStatus,
  FoodImportRow,
  FoodPayment,
  PaymentStatus,
} from '../interfaces/food-import-row.interface';

const FIRST_PAYMENT_COLUMN = 12;
const LAST_PAYMENT_COLUMN = 23;
const FIRST_PAYMENT_YEAR = 2559;
const CLOSED_ROW_COLORS = new Set(['FFF7CAAC', 'FFFF9999']);

export class FoodParser {
  parse(worksheet: Worksheet): FoodImportRow[] {
    const results: FoodImportRow[] = [];

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

      const expiry = this.parseExpiryDate(row.getCell(11).value);

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
          expiryDay: expiry.day,
          expiryMonth: expiry.month,
        },
        payments: this.parsePayments(row),
      });
    });

    return results;
  }

  private parsePayments(row: Row): FoodPayment[] {
    const payments: FoodPayment[] = [];

    for (
      let column = FIRST_PAYMENT_COLUMN, year = FIRST_PAYMENT_YEAR;
      column <= LAST_PAYMENT_COLUMN;
      column++, year++
    ) {
      const rawValue = this.toText(row.getCell(column).value);
      let status: PaymentStatus;

      if (rawValue === '/') {
        status = 'PAID';
      } else if (!rawValue) {
        status = 'EMPTY';
      } else {
        status = 'NOTE';
      }

      payments.push({ year, status, rawValue });
    }

    return payments;
  }

  private getEstablishmentStatus(row: Row): EstablishmentStatus {
    return this.isClosedRow(row) ? 'CLOSED' : 'ACTIVE';
  }

  private isClosedRow(row: Row): boolean {
    let isClosed = false;

    row.eachCell({ includeEmpty: true }, (cell) => {
      const fill = cell.fill;

      if (
        fill?.type === 'pattern' &&
        fill.fgColor?.argb &&
        CLOSED_ROW_COLORS.has(fill.fgColor.argb.toUpperCase())
      ) {
        isClosed = true;
      }
    });

    return isClosed;
  }

  private parseExpiryDate(value: unknown): {
    day: number | null;
    month: number | null;
  } {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return { day: value.getDate(), month: value.getMonth() + 1 };
    }

    const raw = this.toText(value);

    if (!raw) {
      return { day: null, month: null };
    }

    const normalized = raw
      .toLowerCase()
      .replace(/[.\-/]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const dayMatch = normalized.match(/^(\d{1,2})/);
    const day = dayMatch ? Number(dayMatch[1]) : null;
    const month = this.findMonth(normalized);

    return {
      day: day !== null && day >= 1 && day <= 31 ? day : null,
      month,
    };
  }

  private findMonth(value: string): number | null {
    const monthAliases: ReadonlyArray<readonly [number, readonly string[]]> = [
      [1, ['มค', 'มกราคม', 'jan', 'january']],
      [2, ['กพ', 'กุมภาพันธ์', 'feb', 'february']],
      [3, ['มีค', 'มีนา', 'มีนาคม', 'mar', 'march']],
      [4, ['เมย', 'เมษายน', 'apr', 'april']],
      [5, ['พค', 'พฤษภาคม', 'may']],
      [6, ['มิย', 'มิถุนายน', 'jun', 'june']],
      [7, ['กค', 'กรกฎาคม', 'jul', 'july']],
      [8, ['สค', 'สิงหาคม', 'aug', 'august']],
      [9, ['กย', 'กันยายน', 'sep', 'sept', 'september']],
      [10, ['ตค', 'ตุลาคม', 'oct', 'october']],
      [11, ['พย', 'พฤศจิกายน', 'nov', 'november']],
      [12, ['ธค', 'ธันวาคม', 'dec', 'december']],
    ];

    const compact = value.replace(/\s/g, '');

    for (const [month, aliases] of monthAliases) {
      if (aliases.some((alias) => compact.includes(alias))) {
        return month;
      }
    }

    const numericMonth = value.match(/^\d{1,2}\s+(\d{1,2})(?:\s|$)/);
    const parsed = numericMonth ? Number(numericMonth[1]) : null;

    return parsed !== null && parsed >= 1 && parsed <= 12 ? parsed : null;
  }

  private normalizePhone(value: unknown): string | null {
    return this.toText(value)?.replace(/[\s-]/g, '') ?? null;
  }

  private toText(value: unknown): string | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    if (
      typeof value === 'string' ||
      typeof value === 'number' ||
      typeof value === 'boolean' ||
      typeof value === 'bigint'
    ) {
      const text = String(value).trim();
      return text || null;
    }

    if (typeof value === 'object') {
      const cellValue = value as {
        text?: unknown;
        result?: unknown;
        richText?: Array<{ text: string }>;
      };

      if (Array.isArray(cellValue.richText)) {
        const text = cellValue.richText
          .map((part) => part.text)
          .join('')
          .trim();
        return text || null;
      }

      if (typeof cellValue.text === 'string') {
        return cellValue.text.trim() || null;
      }

      if (cellValue.result !== undefined) {
        return this.toText(cellValue.result);
      }
    }

    return null;
  }

  private toNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    if (typeof value === 'number') {
      return value;
    }

    const text = this.toText(value);

    if (!text) {
      return null;
    }

    const parsed = Number(text.replace(/,/g, ''));
    return Number.isNaN(parsed) ? null : parsed;
  }
}
