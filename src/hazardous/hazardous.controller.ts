import { Controller, Post, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { HazardousService } from './hazardous.service.js';

@Controller('hazardous')
export class HazardousController {
    constructor(private readonly hazardousService: HazardousService) { }

    @Post('import')
    @UseInterceptors(FileInterceptor('file'))
    async importExcel(@UploadedFile() file: Express.Multer.File) {
        const data = await this.hazardousService.parseExcel(file);

        return {
            total: data.length,
            data,
        };
    }
}
