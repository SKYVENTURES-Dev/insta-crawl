// src/mail/mail.service.ts
import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import nodemailer, { Transporter } from 'nodemailer';
import SMTPTransport from 'nodemailer/lib/smtp-transport';

@Injectable()
export class MailService {
  private transporter: Transporter<SMTPTransport.SentMessageInfo>;

  constructor(private readonly configSevice: ConfigService) {
    // 하드코딩된 SMTP 계정
    const transporterOptions: SMTPTransport.Options = {
      host: this.configSevice.get<string>('SMTP_HOST'), // Gmail SMTP
      port: this.configSevice.get<string>('SMTP_PORT'),
      secure: false, // TLS 사용 여부
      auth: {
        user: this.configSevice.get<string>('SMTP_USER'), // 보내는 계정
        pass: this.configSevice.get<string>('SMTP_PASS'), // 보내는 계정 비밀번호 (앱 비밀번호)
      },
    };

    this.transporter = nodemailer.createTransport(transporterOptions);
  }

  todayDate(): string {
    const now = new Date();
    const year = now.getFullYear().toString().slice(-2); // 마지막 두 자리
    const month = (now.getMonth() + 1).toString().padStart(2, '0'); // 월은 0부터 시작하므로 +1
    const day = now.getDate().toString().padStart(2, '0');

    return `${year}${month}${day}`;
  }

  async sendFileOnlyMail(
    subject: string,
    url: string,
  ): Promise<SMTPTransport.SentMessageInfo> {
    const mailOptions: SMTPTransport.Options = {
      from: `"skyventures_dev" ${this.configSevice.get<string>('SMTP_USER')}`,
      to: 'uej0868@gmail.com',
      // to: 'eslee@hahmpartners.com', // 받는 계정
      // cc: [
      //   'uniqlo_pr@hahmpartners.com',
      //   'ceo@skyventures.co.kr',
      //   'tkdwns27@omtlabs.com',
      // ],
      subject,
      text: `${this.todayDate()}유니클로 데이터 크롤링 파일 전달드립니다. ${url}`, // 간단한 안내 문구
    };

    const info = await this.transporter.sendMail(mailOptions);
    console.log('Email sent:', info.messageId);
    return info;
  }
}
