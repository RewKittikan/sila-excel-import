import ExcelJS from 'exceljs';
import { FoodParser } from './food.parser';

describe('FoodParser', () => {
  const parser = new FoodParser();

  it('parses establishment, license, and payment data', () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('มค');
    const row = worksheet.getRow(5);

    row.getCell(1).value = 1;
    row.getCell(2).value = ' ผู้ประกอบการ ';
    row.getCell(3).value = ' ร้านอาหาร ';
    row.getCell(4).value = 'จำหน่ายอาหาร';
    row.getCell(5).value = '99/1';
    row.getCell(6).value = 5;
    row.getCell(7).value = 'ศิลา';
    row.getCell(8).value = '081-234 5678';
    row.getCell(9).value = '1,500';
    row.getCell(10).value = 120;
    row.getCell(11).value = '16 เม.ย. 2569';
    row.getCell(12).value = '/';
    row.getCell(13).value = 'ยกเว้น';

    const [result] = parser.parse(worksheet);

    expect(result.establishment).toEqual({
      ownerName: 'ผู้ประกอบการ',
      establishmentName: 'ร้านอาหาร',
      businessType: 'จำหน่ายอาหาร',
      addressNo: '99/1',
      moo: '5',
      subdistrict: 'ศิลา',
      phone: '0812345678',
      status: 'ACTIVE',
    });
    expect(result.license).toEqual({
      feeAmount: 1500,
      areaSqm: 120,
      expiryDay: 16,
      expiryMonth: 4,
    });
    expect(result.payments).toHaveLength(12);
    expect(result.payments.slice(0, 3)).toEqual([
      { year: 2559, status: 'PAID', rawValue: '/' },
      { year: 2560, status: 'NOTE', rawValue: 'ยกเว้น' },
      { year: 2561, status: 'EMPTY', rawValue: null },
    ]);
  });

  it('marks a row with the food workbook red fill as closed', () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('กพ');
    const row = worksheet.getRow(5);

    row.getCell(1).value = 1;
    row.getCell(2).value = 'ผู้ประกอบการ';
    row.getCell(22).value = 'ยกเลิกกิจการ';
    row.getCell(22).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF7CAAC' },
    };

    const [result] = parser.parse(worksheet);

    expect(result.establishment.status).toBe('CLOSED');
    expect(result.payments[10]).toEqual({
      year: 2569,
      status: 'NOTE',
      rawValue: 'ยกเลิกกิจการ',
    });
  });

  it('ignores headers and rows without an owner or establishment name', () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('มีนา');

    worksheet.getCell('A3').value = 'ลำดับที่';
    worksheet.getCell('A5').value = 1;

    expect(parser.parse(worksheet)).toEqual([]);
  });
});
