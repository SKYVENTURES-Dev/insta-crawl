import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { chromium } from 'playwright';
import * as fs from 'fs/promises';
import { Cron } from '@nestjs/schedule';

@Injectable()
export class SessionRefreshService {
  constructor(private readonly configService: ConfigService) {
    // this.sessionRefreshLogin();
  }

  @Cron('0 6 * * *')
  async sessionRefreshLogin() {
    const id = this.configService.get<string>('ID2') || '';
    const password = this.configService.get<string>('PASSWORD2') || '';

    const browser = await chromium.launch({
      headless: false,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
      ...(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH && {
        executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
      }),
    });

    const context = await browser.newContext({
      viewport: { width: 1920, height: 1080 },
      userAgent:
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    });

    const page = await context.newPage();

    // 기존 쿠키가 있다면 로드
    try {
      const cookiesData = await fs.readFile('./cookies.json', 'utf8');
      const cookies = JSON.parse(cookiesData);
      await context.addCookies(cookies);
    } catch (error) {
      console.log('기존 쿠키 파일이 없습니다.');
    }

    await page.goto('https://www.instagram.com', {
      waitUntil: 'domcontentloaded',
      timeout: 30000,
    });

    // 로그인 필드가 있는지 확인하여 쿠키 상태 체크
    try {
      await page.waitForSelector('input[name="username"]', { timeout: 3000 });
      console.log('쿠키가 만료됐습니다. 다시 로그인합니다.');

      // 로그인 프로세스
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

      // 정보 저장 버튼 처리
      try {
        await page.waitForSelector('button[class*="_asx2"]', {
          timeout: 30000,
        });
        console.log('정보 저장 모달이 나타남');
        await page.click('button[class*="_asx2"]');
      } catch (error) {
        console.log('정보 저장 버튼을 찾을 수 없습니다.');
      }

      // 쿠키 저장
      const cookies = await context.cookies();
      await fs.writeFile('./cookies.json', JSON.stringify(cookies, null, 2));
      console.log('새로운 쿠키가 저장됨');
    } catch (error) {
      console.log('쿠키가 아직 활성화 되어 있습니다.');
    }

    await browser.close();
  }
}
