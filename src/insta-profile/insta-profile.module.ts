import { Module } from '@nestjs/common';
import { InstaProfileService } from './insta-profile.service';
import { HttpModule } from '@nestjs/axios';
import { MailService } from 'src/mail/mail.service';
import { GoogleDriveService } from 'src/google-drive/google-drive.service';

@Module({
  imports: [HttpModule.register({})],
  providers: [InstaProfileService, MailService, GoogleDriveService],
})
export class InstaProfileModule {}
