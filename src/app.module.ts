import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { HazardousModule } from './hazardous/hazardous.module';
import { FoodModule } from './food/food.module';

@Module({
  imports: [HazardousModule, FoodModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
