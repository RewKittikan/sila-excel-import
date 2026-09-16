import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { HazardousModule } from './hazardous/hazardous.module';

@Module({
  imports: [HazardousModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
