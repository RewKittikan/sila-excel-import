import { Module } from '@nestjs/common';
import { HazardousController } from './hazardous.controller';
import { HazardousService } from './hazardous.service';

@Module({
  controllers: [HazardousController],
  providers: [HazardousService]
})
export class HazardousModule {}
