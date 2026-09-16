import { Test, TestingModule } from '@nestjs/testing';
import { HazardousController } from './hazardous.controller';

describe('HazardousController', () => {
  let controller: HazardousController;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      controllers: [HazardousController],
    }).compile();

    controller = module.get<HazardousController>(HazardousController);
  });

  it('should be defined', () => {
    expect(controller).toBeDefined();
  });
});
