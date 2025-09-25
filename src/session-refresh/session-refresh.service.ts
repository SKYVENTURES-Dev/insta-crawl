import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { chromium } from 'playwright';
import * as fs from 'fs/promises';
import { Cron } from '@nestjs/schedule';

@Injectable()
export class SessionRefreshService {
  constructor(private readonly configService: ConfigService) {
    this.sessionRefreshLogin();
  }

  @Cron('58 23 * * *')
  async sessionRefreshLogin() {
    const id = this.configService.get<string>('ID2') || '';
    const password = this.configService.get<string>('PASSWORD2') || '';

    const browser = await chromium.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    const context = await browser.newContext({
      viewport: { width: 1920, height: 1080 },
      userAgent:
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    });

    const page = await context.newPage();
    await page.goto('https://www.instagram.com', {
      waitUntil: 'domcontentloaded',
      timeout: 30000,
    });

    await page.click('input[name="username"]');
    await page.focus('input[name="username"]');

    await page.keyboard.type(id, {
      delay: 100 + Math.random() * 100,
    });

    await page.click('input[type="password"]');
    await page.focus('input[type="password"]');

    await page.keyboard.type(password, {
      delay: 100 + Math.random() * 100,
    });

    await page.click('button[type="submit"]');
    await page.waitForSelector('.x1i10hfl');

    const cookies = await context.cookies();
    fs.writeFile('./cookies.json', JSON.stringify(cookies, null, 2));
    console.log('쿠키 저장됨 ');
  }
}
