import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { FoodImportRow } from './interfaces/food-import-row.interface';
import { FoodParser } from './parsers/food.parser';

@Injectable()
export class FoodService {
  private readonly parser = new FoodParser();

  async parseExcel(file: Express.Multer.File): Promise<FoodImportRow[]> {
    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.load(file.buffer as any);

    const results: FoodImportRow[] = [];

    for (const worksheet of workbook.worksheets) {
      results.push(...this.parser.parse(worksheet));
    }

    return results;
  }
}
