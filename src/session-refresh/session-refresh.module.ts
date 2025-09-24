import { Module } from '@nestjs/common';
import { SessionRefreshService } from './session-refresh.service';

@Module({
  providers: [SessionRefreshService],
})
export class SessionRefreshModule {}
