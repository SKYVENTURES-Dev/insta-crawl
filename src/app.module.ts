import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
// import { InstaCrawlModule } from './insta-crawl/insta-crawl.module';
import { ScheduleModule } from '@nestjs/schedule';
// import { TiktokModule } from './tiktok/tiktok.module';
import { InstaProfileModule } from './insta-profile/insta-profile.module';
// import { BobaedreamModule } from './bobaedream/bobaedream.module';
// import { FunnyModule } from './funny/funny.module';
// import { InvenModule } from './inven/inven.module';
// import { MlbParkModule } from './mlb-park/mlb-park.module';
// import { NatePannModule } from './nate-pann/nate-pann.module';
// import { TodayHumerModule } from './today-humer/today-humer.module';
// import { YgosuModule } from './ygosu/ygosu.module';
import { MailModule } from './mail/mail.module';
import { ConfigModule } from '@nestjs/config';
import { GoogleDriveModule } from './google-drive/google-drive.module';

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
    // RuriModule,
    // InvenModule,
    // FmModule,
    // MlbParkModule,
    // FunnyModule,
    // InstizModule,
    // TodayHumerModule,
    // ClienModule,
    // YgosuModule,
    // BobaedreamModule,
    // NatePannModule,
    // DcModule,
    // InstaCrawlModule,
    // TiktokModule,
    InstaProfileModule,
    MailModule,
    GoogleDriveModule,
  ],
})
export class AppModule {}
