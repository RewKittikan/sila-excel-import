import { Module } from '@nestjs/common';
import { FoodController } from './food.controller.js';
import { FoodService } from './food.service.js';

@Module({
  controllers: [FoodController],
  providers: [FoodService],
})
export class FoodModule {}
