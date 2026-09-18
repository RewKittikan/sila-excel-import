import { Test, TestingModule } from '@nestjs/testing';
import ExcelJS from 'exceljs';
import { FoodService } from './food.service';

describe('FoodService', () => {
  let service: FoodService;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [FoodService],
    }).compile();

    service = module.get<FoodService>(FoodService);
  });

  it('should be defined', () => {
    expect(service).toBeDefined();
  });

  it('loads and combines rows from every worksheet', async () => {
    const workbook = new ExcelJS.Workbook();

    for (const sheetName of ['มค', 'กพ']) {
      const worksheet = workbook.addWorksheet(sheetName);
      worksheet.getCell('A5').value = 1;
      worksheet.getCell('B5').value = `owner-${sheetName}`;
    }

    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
    const result = await service.parseExcel({ buffer } as Express.Multer.File);

    expect(result.map((row) => row.establishment.ownerName)).toEqual([
      'owner-มค',
      'owner-กพ',
    ]);
  });
});
