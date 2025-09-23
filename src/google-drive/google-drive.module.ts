import { Module } from '@nestjs/common';
import { GoogleDriveService } from './google-drive.service';
import { HttpModule } from '@nestjs/axios';

@Module({
  imports: [HttpModule.register({})],
  providers: [GoogleDriveService],
})
export class GoogleDriveModule {}
