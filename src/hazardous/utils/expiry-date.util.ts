export interface ParsedExpiryDate {
  day: number | null;
  month: number | null;
}

const monthMap: Record<string, number> = {
  // English
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

  // Thai
  มค: 1,
  มกราคม: 1,

  กพ: 2,
  กุมภาพันธ์: 2,

  มีค: 3,
  มีนา: 3,
  มีนาคม: 3,

  เมย: 4,
  เมษายน: 4,

  พค: 5,
  พฤษภาคม: 5,

  มิย: 6,
  มิถุนายน: 6,

  กค: 7,
  กรกฎาคม: 7,

  สค: 8,
  สิงหาคม: 8,

  กย: 9,
  กันยายน: 9,

  ตค: 10,
  ตุลาคม: 10,

  พย: 11,
  พฤศจิกายน: 11,

  ธค: 12,
  ธันวา: 12,
  ธันวาคม: 12,
};

export function parseExpiryDate(value: unknown): ParsedExpiryDate {
  if (value === null || value === undefined || value === '') {
    return {
      day: null,
      month: null,
    };
  }

  // Excel Date จริง
  // ปีที่ติดมา เช่น 1979 / 2017 เราไม่ใช้
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return {
      day: value.getDate(),
      month: value.getMonth() + 1,
    };
  }

  const raw = String(value).trim();

  if (!raw) {
    return {
      day: null,
      month: null,
    };
  }

  const normalized = normalizeDateText(raw);

  const dayMatch = normalized.match(/^(\d{1,2})/);

  if (!dayMatch) {
    return {
      day: null,
      month: null,
    };
  }

  const day = Number(dayMatch[1]);

  if (day < 1 || day > 31) {
    return {
      day: null,
      month: null,
    };
  }

  const month = findMonth(normalized);

  if (!month) {
    return {
      day,
      month: null,
    };
  }

  return {
    day,
    month,
  };
}

function normalizeDateText(value: string): string {
  return value
    .toLowerCase()
    .replace(/\./g, '')
    .replace(/-/g, ' ')
    .replace(/\//g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findMonth(value: string): number | null {
  const parts = value.split(' ');

  for (const part of parts) {
    const month = monthMap[part];

    if (month) {
      return month;
    }
  }

  // รองรับแบบไม่มีเว้นวรรค เช่น 16เมย
  for (const [key, month] of Object.entries(monthMap)) {
    if (value.includes(key)) {
      return month;
    }
  }

  return null;
}