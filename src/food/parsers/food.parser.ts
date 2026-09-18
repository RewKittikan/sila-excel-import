import { Cell, Row, Worksheet } from 'exceljs';
import {
  FoodImportRow,
  FoodPayment,
  PaymentStatus,
  EstablishmentStatus,
} from '../interfaces/food-import-row.interface';

const FIRST_DATA_ROW = 4;

const COL = {
  NO: 1,
  OWNER_NAME: 2,
  ESTABLISHMENT_NAME: 3,
  FOOD_TYPE: 4,

  // E เป็นหัวรวม "ที่อยู่"
  ADDRESS_NO: 6,
  MOO: 7,
  SUBDISTRICT: 8,

  PHONE: 9,
  FEE_AMOUNT: 10,
  AREA_SQM: 11,
  EXPIRY: 12,

  // M = 2559
  PAYMENT_START: 13,
} as const;

const PAYMENT_START_YEAR = 2559;
const PAYMENT_END_YEAR = 2570;

/**
 * สีที่ใช้ไฮไลท์ว่าปิด/เลิกกิจการ
 */
const CLOSED_FILL_COLORS = new Set([
  'C6E0B4',
  'F8CBAD',
  'F4B084',
  'A9D08E',
]);

const THAI_MONTHS: Record<string, number> = {
  'ม.ค.': 1,
  มค: 1,
  มกราคม: 1,

  'ก.พ.': 2,
  กพ: 2,
  กุมภาพันธ์: 2,

  'มี.ค.': 3,
  มีค: 3,
  มีนาคม: 3,

  'เม.ย.': 4,
  เมย: 4,
  เมษายน: 4,

  'พ.ค.': 5,
  พค: 5,
  พฤษภาคม: 5,

  'มิ.ย.': 6,
  มิย: 6,
  มิถุนายน: 6,

  'ก.ค.': 7,
  กค: 7,
  กรกฎาคม: 7,

  'ส.ค.': 8,
  สค: 8,
  สิงหาคม: 8,

  'ก.ย.': 9,
  กย: 9,
  กันยายน: 9,

  'ต.ค.': 10,
  ตค: 10,
  ตุลาคม: 10,

  'พ.ย.': 11,
  พย: 11,
  พฤศจิกายน: 11,

  'ธ.ค.': 12,
  ธค: 12,
  ธันวาคม: 12,
};

const EN_MONTHS: Record<string, number> = {
  jan: 1,
  january: 1,
  feb: 2,
  february: 2,
  mar: 3,
  march: 3,
  apr: 4,
  april: 4,
  may: 5,
  jun: 6,
  june: 6,
  jul: 7,
  july: 7,
  aug: 8,
  august: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  october: 10,
  nov: 11,
  november: 11,
  dec: 12,
  december: 12,
};

export function parseFoodWorksheet(
  worksheet: Worksheet,
): FoodImportRow[] {
  const result: FoodImportRow[] = [];

  for (
    let rowNumber = FIRST_DATA_ROW;
    rowNumber <= worksheet.rowCount;
    rowNumber++
  ) {
    const row = worksheet.getRow(rowNumber);

    if (isEmptyFoodRow(row)) {
      continue;
    }

    const ownerName = textOrNull(
      row.getCell(COL.OWNER_NAME),
    );

    const establishmentName = textOrNull(
      row.getCell(COL.ESTABLISHMENT_NAME),
    );

    /**
     * ถ้าไม่มีทั้งผู้ประกอบการและชื่อสถานประกอบการ
     * ถือว่าไม่ใช่แถวข้อมูลร้าน
     */
    if (!ownerName && !establishmentName) {
      continue;
    }

    const rawFoodType = textOrNull(
      row.getCell(COL.FOOD_TYPE),
    );

    const {
      businessType,
      licenseCat,
    } = parseFoodType(rawFoodType);

    const payments = parsePayments(row);

    const status = parseEstablishmentStatus(
      row,
      payments,
    );

    const expiry = parseExpiryDayMonth(
      row.getCell(COL.EXPIRY),
    );

    result.push({
      establishment: {
        ownerName,
        establishmentName,
        businessType,

        phone: textOrNull(
          row.getCell(COL.PHONE),
        ),

        addressNo: textOrNull(
          row.getCell(COL.ADDRESS_NO),
        ),

        moo: textOrNull(
          row.getCell(COL.MOO),
        ),

        subdistrict: textOrNull(
          row.getCell(COL.SUBDISTRICT),
        ),

        status,
      },

      license: {
        licenseCat,

        feeAmount: numberOrNull(
          row.getCell(COL.FEE_AMOUNT),
        ),

        areaSqm: numberOrNull(
          row.getCell(COL.AREA_SQM),
        ),

        expiryDay: expiry.day,
        expiryMonth: expiry.month,
      },

      payments,
    });
  }

  return result;
}

/**
 * Excel:
 *   จำหน่าย
 *   สะสม
 *
 * DB:
 * establishment.businessType:
 *   สถานที่จำหน่ายอาหาร
 *   สถานที่สะสมอาหาร
 *
 * license.licenseCat:
 *   จำหน่าย
 *   สะสม
 */
function parseFoodType(
  rawValue: string | null,
): {
  businessType: string | null;
  licenseCat: string | null;
} {
  if (!rawValue) {
    return {
      businessType: null,
      licenseCat: null,
    };
  }

  const value = normalizeText(rawValue);

  if (value.includes('สะสม')) {
    return {
      businessType: 'สถานที่สะสมอาหาร',
      licenseCat: 'สะสม',
    };
  }

  if (value.includes('จำหน่าย')) {
    return {
      businessType: 'สถานที่จำหน่ายอาหาร',
      licenseCat: 'จำหน่าย',
    };
  }

  /**
   * ถ้าเจอค่าที่ไม่รู้จัก
   * ไม่เดา licenseCat
   */
  return {
    businessType: rawValue.trim(),
    licenseCat: null,
  };
}

function parsePayments(
  row: Row,
): FoodPayment[] {
  const payments: FoodPayment[] = [];

  for (
    let year = PAYMENT_START_YEAR;
    year <= PAYMENT_END_YEAR;
    year++
  ) {
    const column =
      COL.PAYMENT_START +
      (year - PAYMENT_START_YEAR);

    const cell = row.getCell(column);

    const rawValue = textOrNull(cell);

    payments.push({
      year,
      status: parsePaymentStatus(rawValue),
      rawValue,
    });
  }

  return payments;
}

function parsePaymentStatus(
  rawValue: string | null,
): PaymentStatus {
  if (!rawValue) {
    return 'EMPTY';
  }

  const value = normalizeText(rawValue);

  /**
   * / = ชำระแล้ว
   */
  if (
    value === '/' ||
    value === '✓' ||
    value === '✔'
  ) {
    return 'PAID';
  }

  /**
   * เช่น:
   * เลิก
   * เลิก66
   * เลิก67
   * ยกเลิก68
   * ยกเลิก69
   */
  if (isClosedText(rawValue)) {
    return 'NOTE';
  }

  /**
   * หมายเหตุอื่น ๆ
   */
  if (
    value.includes('เปลี่ยน') ||
    value.includes('สงสัย') ||
    value.includes('หมายเหตุ')
  ) {
    return 'NOTE';
  }

  return 'UNKNOWN';
}

function parseEstablishmentStatus(
  row: Row,
  payments: FoodPayment[],
): EstablishmentStatus {
  /**
   * ถ้ามีคำว่า เลิก / ยกเลิก
   * ใน payment ปีใดปีหนึ่ง
   */
  const hasClosedText = payments.some(
    (payment) =>
      isClosedText(payment.rawValue),
  );

  if (hasClosedText) {
    return 'CLOSED';
  }

  /**
   * หรือมี fill สีที่กำหนดว่าปิด
   */
  if (rowHasClosedColor(row)) {
    return 'CLOSED';
  }

  return 'ACTIVE';
}

function isClosedText(
  value: string | null,
): boolean {
  if (!value) {
    return false;
  }

  const normalized = normalizeText(value);

  return normalized.includes('เลิก');
}

/**
 * เช็กทุก cell ในแถวว่า
 * มีสีที่หมายถึง CLOSED หรือไม่
 */
function rowHasClosedColor(
  row: Row,
): boolean {
  const lastColumn =
    COL.PAYMENT_START +
    (PAYMENT_END_YEAR -
      PAYMENT_START_YEAR);

  for (
    let column = 1;
    column <= lastColumn;
    column++
  ) {
    if (
      isClosedFillColor(
        row.getCell(column),
      )
    ) {
      return true;
    }
  }

  return false;
}

/**
 * คืนสี fill ของ cell เช่น
 * FFC6E0B4
 */
function getCellFillColor(
  cell: Cell,
): string | null {
  const fill = cell.fill;

  if (
    !fill ||
    fill.type !== 'pattern'
  ) {
    return null;
  }

  const argb = fill.fgColor?.argb;

  if (!argb) {
    return null;
  }

  return argb.toUpperCase();
}

/**
 * เช็กว่า fill เป็นหนึ่งในสี CLOSED
 *
 * ExcelJS อาจคืนเป็น:
 * FFC6E0B4
 *
 * แต่เราเก็บ:
 * C6E0B4
 */
function isClosedFillColor(
  cell: Cell,
): boolean {
  const actual =
    getCellFillColor(cell);

  if (!actual) {
    return false;
  }

  const rgb =
    actual.length === 8
      ? actual.slice(2)
      : actual;

  return CLOSED_FILL_COLORS.has(rgb);
}

/**
 * ตอนนี้ยังไม่เชื่อปีใน Excel
 * เลยเอาเฉพาะ day/month
 *
 * รองรับ:
 * 23-มี.ค.
 * 4-ต.ค.
 * 04-Jan
 * 15-Oct
 * 4/1/2569
 * 4/1
 * Date object จาก Excel
 */
function parseExpiryDayMonth(
  cell: Cell,
): {
  day: number | null;
  month: number | null;
} {
  const actualValue =
    unwrapCellValue(cell.value);

  /**
   * ExcelJS อ่านเป็น Date จริง
   * เราสนใจแค่ day/month
   */
  if (actualValue instanceof Date) {
    return {
      day: actualValue.getDate(),
      month:
        actualValue.getMonth() + 1,
    };
  }

  const raw =
    cleanText(cell.text) ??
    cleanText(actualValue);

  if (!raw) {
    return {
      day: null,
      month: null,
    };
  }

  const value = raw
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');

  /**
   * 4/1/2569
   * 04/01/2026
   * 4/1
   */
  const slashMatch = value.match(
    /^(\d{1,2})\/(\d{1,2})(?:\/\d{2,4})?$/,
  );

  if (slashMatch) {
    return validateDayMonth(
      Number(slashMatch[1]),
      Number(slashMatch[2]),
    );
  }

  /**
   * 4-1-2569
   * 4-1
   */
  const numericDashMatch =
    value.match(
      /^(\d{1,2})-(\d{1,2})(?:-\d{2,4})?$/,
    );

  if (numericDashMatch) {
    return validateDayMonth(
      Number(numericDashMatch[1]),
      Number(numericDashMatch[2]),
    );
  }

  /**
   * 23-มี.ค.
   * 23-มีค
   * 23-มีนาคม
   */
  const thaiMonthMatch =
    value.match(
      /^(\d{1,2})-(.+)$/,
    );

  if (thaiMonthMatch) {
    const day =
      Number(thaiMonthMatch[1]);

    const monthText =
      thaiMonthMatch[2]
        .replace(/\.$/, '')
        .trim();

    const month =
      findThaiMonth(monthText);

    if (month !== null) {
      return validateDayMonth(
        day,
        month,
      );
    }
  }

  /**
   * 04-Jan
   * 15-Oct
   */
  const englishMatch =
    value.match(
      /^(\d{1,2})-([a-z]+)\.?$/,
    );

  if (englishMatch) {
    const day =
      Number(englishMatch[1]);

    const month =
      EN_MONTHS[
        englishMatch[2]
      ] ?? null;

    if (month !== null) {
      return validateDayMonth(
        day,
        month,
      );
    }
  }

  return {
    day: null,
    month: null,
  };
}

function findThaiMonth(
  value: string,
): number | null {
  const normalized = value
    .replace(/\./g, '')
    .trim();

  for (
    const [key, month]
    of Object.entries(THAI_MONTHS)
  ) {
    const normalizedKey =
      key
        .replace(/\./g, '')
        .trim();

    if (
      normalized === normalizedKey ||
      normalized.startsWith(
        normalizedKey,
      )
    ) {
      return month;
    }
  }

  return null;
}

function validateDayMonth(
  day: number,
  month: number,
): {
  day: number | null;
  month: number | null;
} {
  if (
    !Number.isInteger(day) ||
    !Number.isInteger(month) ||
    day < 1 ||
    day > 31 ||
    month < 1 ||
    month > 12
  ) {
    return {
      day: null,
      month: null,
    };
  }

  return {
    day,
    month,
  };
}

function numberOrNull(
  cell: Cell,
): number | null {
  const value =
    unwrapCellValue(cell.value);

  if (
    typeof value === 'number' &&
    Number.isFinite(value)
  ) {
    return value;
  }

  const text = cleanText(value);

  if (!text) {
    return null;
  }

  const normalized = text
    .replace(/,/g, '')
    .replace(/[^\d.-]/g, '');

  if (
    !normalized ||
    normalized === '-' ||
    normalized === '.'
  ) {
    return null;
  }

  const parsed =
    Number(normalized);

  return Number.isFinite(parsed)
    ? parsed
    : null;
}

function textOrNull(
  cell: Cell,
): string | null {
  /**
   * cell.text = ค่าที่ Excel แสดง
   */
  const fromText =
    cleanText(cell.text);

  if (fromText) {
    return fromText;
  }

  return cleanText(
    unwrapCellValue(cell.value),
  );
}

function unwrapCellValue(
  value: unknown,
): unknown {
  if (
    value === null ||
    value === undefined
  ) {
    return null;
  }

  if (value instanceof Date) {
    return value;
  }

  if (
    typeof value !== 'object'
  ) {
    return value;
  }

  const data =
    value as Record<
      string,
      unknown
    >;

  /**
   * Formula
   */
  if ('result' in data) {
    return data.result;
  }

  /**
   * Rich Text
   */
  if (
    Array.isArray(
      data.richText,
    )
  ) {
    return data.richText
      .map((item) => {
        if (
          typeof item ===
            'object' &&
          item !== null &&
          'text' in item
        ) {
          return String(
            (
              item as {
                text: unknown;
              }
            ).text,
          );
        }

        return '';
      })
      .join('');
  }

  /**
   * Hyperlink / text object
   */
  if ('text' in data) {
    return data.text;
  }

  return null;
}

function cleanText(
  value: unknown,
): string | null {
  if (
    value === null ||
    value === undefined
  ) {
    return null;
  }

  const text = String(value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (
    text === '' ||
    text === '-' ||
    text.toLowerCase() ===
      'null' ||
    text.toLowerCase() ===
      'undefined'
  ) {
    return null;
  }

  return text;
}

function normalizeText(
  value: string,
): string {
  return value
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function isEmptyFoodRow(
  row: Row,
): boolean {
  const owner =
    textOrNull(
      row.getCell(
        COL.OWNER_NAME,
      ),
    );

  const establishment =
    textOrNull(
      row.getCell(
        COL.ESTABLISHMENT_NAME,
      ),
    );

  const type =
    textOrNull(
      row.getCell(
        COL.FOOD_TYPE,
      ),
    );

  const addressNo =
    textOrNull(
      row.getCell(
        COL.ADDRESS_NO,
      ),
    );

  const moo =
    textOrNull(
      row.getCell(COL.MOO),
    );

  return (
    !owner &&
    !establishment &&
    !type &&
    !addressNo &&
    !moo
  );
}

export class FoodParser {
  parse(worksheet: Worksheet): FoodImportRow[] {
    return parseFoodWorksheet(worksheet);
  }
}