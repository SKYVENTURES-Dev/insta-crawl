import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ScheduleModule } from '@nestjs/schedule';
import { InstaProfileModule } from './insta-profile/insta-profile.module';
import { MailModule } from './mail/mail.module';
import { ConfigModule } from '@nestjs/config';
import { GoogleDriveModule } from './google-drive/google-drive.module';
import { SessionRefreshModule } from './session-refresh/session-refresh.module';

@Module({
  controllers: [AppController],
  providers: [AppService],
  imports: [
    ConfigModule.forRoot({
      isGlobal: true,
      envFilePath:
        process.env.NODE_ENV === 'production' ? '.env.local' : '.env',
    }),
    ScheduleModule.forRoot(),
    InstaProfileModule,
    MailModule,
    SessionRefreshModule,
    GoogleDriveModule,
  ],
})
export class AppModule {}
