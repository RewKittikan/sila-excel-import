import { Module } from '@nestjs/common';
import { HazardousController } from './hazardous.controller.js';
import { HazardousService } from './hazardous.service.js';

@Module({
  controllers: [HazardousController],
  providers: [HazardousService]
})
export class HazardousModule {}
