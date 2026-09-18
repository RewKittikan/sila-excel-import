import { Module } from '@nestjs/common';
import { AppController } from './app.controller.js';
import { AppService } from './app.service.js';
import { HazardousModule } from './hazardous/hazardous.module.js';
import { FoodModule } from './food/food.module.js';

@Module({
  imports: [HazardousModule, FoodModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
