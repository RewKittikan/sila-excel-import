import { Test, TestingModule } from '@nestjs/testing';
import { HazardousService } from './hazardous.service';

describe('HazardousService', () => {
  let service: HazardousService;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [HazardousService],
    }).compile();

    service = module.get<HazardousService>(HazardousService);
  });

  it('should be defined', () => {
    expect(service).toBeDefined();
  });
});
