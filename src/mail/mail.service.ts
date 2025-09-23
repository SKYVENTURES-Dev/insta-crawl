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

  async sendFileOnlyMail(
    subject: string,
    url: string,
  ): Promise<SMTPTransport.SentMessageInfo> {
    const mailOptions: SMTPTransport.Options = {
      from: `"skyventures_dev" ${this.configSevice.get<string>('SMTP_USER')}`,
      to: 'eslee@hahmpartners.com', // 받는 계정
      cc: [
        'uniqlo_pr@hahmpartners.com',
        'ceo@skyventures.co.kr',
        'tkdwns27@omtlabs.com',
      ],
      subject,
      text: `구글 드라이브 링크를 참고해주세요! ${url}`, // 간단한 안내 문구
    };

    const info = await this.transporter.sendMail(mailOptions);
    console.log('Email sent:', info.messageId);
    return info;
  }
}
