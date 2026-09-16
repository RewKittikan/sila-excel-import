import { Injectable } from '@nestjs/common';
import { HazardousImportRow } from './interfaces/hazardous-import-row.interface';
import { EstablishmentParser } from './parsers/establishment.parser';
import * as ExcelJS from 'exceljs';

@Injectable()
export class HazardousService {
    private readonly parser = new EstablishmentParser();

    async parseExcel(file: Express.Multer.File): Promise<HazardousImportRow[]> {
        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.load(file.buffer as any);

        const results: HazardousImportRow[] = [];

        for (const worksheet of workbook.worksheets) {
            const parsed = this.parser.parse(worksheet);
            results.push(...parsed);
        }

        return results;
    }
}
