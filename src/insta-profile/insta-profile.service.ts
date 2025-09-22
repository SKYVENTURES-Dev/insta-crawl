import { Injectable } from '@nestjs/common';
import { HttpService } from '@nestjs/axios';
import { chromium, BrowserContext, Page } from 'playwright';
import * as ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';
import { firstValueFrom } from 'rxjs';
import { Cron } from '@nestjs/schedule';
import { MailService } from 'src/mail/mail.service';

interface InstagramProfile {
  username: string;
  posts: string;
  followers: string;
  following: string;
  latestPost: LatestPostInfo;
  status: 'success' | 'failed';
  error?: string;
}

interface LatestPostInfo {
  postUrl: string;
  thumbnailImage: string;
  likes: string;
  postingDate: string;
  postType: 'feed' | 'reels';
  content: string;
  hashtags: string;
  mentions: string;
}

interface CookieData {
  name: string;
  value: string;
  domain: string;
  path: string;
  expires?: number;
  httpOnly?: boolean;
  secure?: boolean;
  sameSite?: 'Strict' | 'Lax' | 'None';
}

@Injectable()
export class InstaProfileService {
  private readonly cookiesPath =
    '/Users/youyeonjoon/Documents/dev/community_crawl/cookies3.json';

  constructor(
    private readonly httpService: HttpService,
    private readonly mailService: MailService,
  ) {
    // this.executeFullProcess('influencerList1.xlsx');
  }

  // @Cron('0 0 * * *')
  @Cron('48 18 * * *')
  async runDailyInstagramCrawling() {
    console.log('🕛 매일 자정 Instagram 크롤링 시작!');
    const filePath = 'data/instagram_profiles_enhanced_result.xlsx';
    try {
      await this.executeFullProcess('influencerList1.xlsx');
      console.log('✅ 매일 자정 Instagram 크롤링 완료!');

      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
      const info = await this.mailService.sendFileOnlyMail(
        'instagram crawl',
        filePath,
      );
      console.log('✅ 메일 발송 완료:', info.messageId);
    } catch (error) {
      console.error('❌ 매일 자정 Instagram 크롤링 실패:', error);
    }
  }

  async executeFullProcess(fileName: string): Promise<void> {
    // 파일 경로 확인 및 처리
    let filePath: string;

    if (path.isAbsolute(fileName)) {
      filePath = fileName;
    } else if (fileName.includes('/') || fileName.includes('\\')) {
      filePath = path.join(process.cwd(), fileName);
    } else {
      const dataPath = path.join(process.cwd(), 'data', fileName);
      const rootPath = path.join(process.cwd(), fileName);

      if (fs.existsSync(dataPath)) {
        filePath = dataPath;
      } else if (fs.existsSync(rootPath)) {
        filePath = rootPath;
      } else {
        throw new Error(
          `파일을 찾을 수 없습니다. 확인한 경로:\n- ${dataPath}\n- ${rootPath}`,
        );
      }
    }

    if (!fs.existsSync(filePath)) {
      throw new Error(`파일을 찾을 수 없습니다: ${filePath}`);
    }

    console.log(`📂 파일 읽기 시작: ${filePath}`);

    try {
      const influencerUrls = await this.readInfluencerUrls(filePath);
      console.log(
        `총 ${influencerUrls.length}개의 인플루언서 URL을 발견했습니다.`,
      );

      const profiles = await this.crawlMultipleProfiles(influencerUrls);

      const outputFilePath = path.join(
        path.dirname(filePath),
        'instagram_profiles_enhanced_result.xlsx',
      );
      await this.saveToExcel(profiles, outputFilePath);

      console.log(`✅ 크롤링 완료! 결과가 ${outputFilePath}에 저장되었습니다.`);
    } catch (error) {
      console.error('❌ 처리 중 오류 발생:', error);
    }
  }

  private async readInfluencerUrls(filePath: string): Promise<string[]> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);

      const worksheet = workbook.getWorksheet(1);
      const influencerUrls: string[] = [];

      if (!worksheet) {
        throw new Error('워크시트를 찾을 수 없습니다');
      }

      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          if (cell.hyperlink) {
            let url = '';

            if (typeof cell.hyperlink === 'string') {
              url = cell.hyperlink;
            } else if (
              typeof cell.hyperlink === 'object' &&
              cell.hyperlink &&
              'hyperlink' in cell.hyperlink
            ) {
              url = (cell.hyperlink as any).hyperlink;
            }

            if (url && url.includes('instagram.com')) {
              influencerUrls.push(url);
              console.log(`하이퍼링크 발견: ${cell.value} -> ${url}`);
            }
          }
        });
      });

      console.log(
        `총 ${influencerUrls.length}개의 Instagram URL을 발견했습니다.`,
      );

      return [...new Set(influencerUrls)];
    } catch (error) {
      throw new Error(`엑셀 파일 읽기 실패: ${error}`);
    }
  }

  private async loadCookiesToContext(
    context: BrowserContext,
  ): Promise<boolean> {
    try {
      const cookieData = await fs.promises.readFile(this.cookiesPath, 'utf-8');
      const cookies = JSON.parse(cookieData) as CookieData[];

      if (cookies && cookies.length > 0) {
        await context.addCookies(cookies);
        console.log('✅ 쿠키가 컨텍스트에 성공적으로 로드되었습니다.');
        return true;
      } else {
        console.warn('쿠키 파일이 비어있습니다.');
        return false;
      }
    } catch (error) {
      console.error('❌ 쿠키 로드 실패:', error);
      return false;
    }
  }

  private async crawlMultipleProfiles(
    influencerUrls: string[],
  ): Promise<InstagramProfile[]> {
    const browser = await chromium.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    const context = await browser.newContext({
      viewport: { width: 1920, height: 1080 },
      userAgent:
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    });

    await this.loadCookiesToContext(context);

    const profiles: InstagramProfile[] = [];
    const batchSize = 3; // 동시 처리할 프로필 수

    console.log(
      `총 ${influencerUrls.length}개 프로필을 ${batchSize}개씩 병렬 처리합니다.`,
    );

    for (let i = 0; i < influencerUrls.length; i += batchSize) {
      const batch = influencerUrls.slice(i, i + batchSize);
      const batchNumber = Math.floor(i / batchSize) + 1;
      const totalBatches = Math.ceil(influencerUrls.length / batchSize);

      console.log(
        `\n=== 배치 ${batchNumber}/${totalBatches} 시작 (${batch.length}개 프로필) ===`,
      );

      // 배치 내 프로필들을 병렬로 처리
      const batchPromises = batch.map(async (url, batchIndex) => {
        const username = this.extractUsernameFromUrl(url);
        const globalIndex = i + batchIndex + 1;

        console.log(
          `[${globalIndex}/${influencerUrls.length}] ${username} 처리 시작`,
        );

        let profile: InstagramProfile | null = null;
        let lastError: Error | null = null;

        // 최대 2번 시도
        for (let attempt = 1; attempt <= 2; attempt++) {
          try {
            console.log(`[${globalIndex}] ${username} - ${attempt}번째 시도`);
            profile = await this.crawlSingleProfile(context, url, username);
            console.log(`[${globalIndex}] ${username} - 성공!`);
            break;
          } catch (error) {
            lastError = error as Error;
            console.error(
              `[${globalIndex}] ${username} ${attempt}번째 시도 실패:`,
              error.message,
            );

            if (attempt < 2) {
              await this.delay(1500 + Math.random() * 1000);
            }
          }
        }

        if (profile) {
          return profile;
        } else {
          console.error(`[${globalIndex}] ${username} - 최종 실패`);
          return {
            username,
            posts: '',
            followers: '',
            following: '',
            latestPost: {
              postUrl: '',
              thumbnailImage: '',
              likes: '',
              postingDate: '',
              postType: 'feed' as const,
              content: '',
              hashtags: '',
              mentions: '',
            },
            status: 'failed' as const,
            error: lastError?.message || '알 수 없는 오류',
          };
        }
      });

      // 배치 결과 대기
      const batchResults = await Promise.allSettled(batchPromises);

      batchResults.forEach((result, batchIndex) => {
        const globalIndex = i + batchIndex + 1;
        if (result.status === 'fulfilled') {
          profiles.push(result.value);
          console.log(
            `[${globalIndex}] 배치 처리 완료: ${result.value.username} - ${result.value.status}`,
          );
        } else {
          console.error(`[${globalIndex}] 배치 처리 실패:`, result.reason);
          // 실패한 경우 기본 프로필 추가
          const username = this.extractUsernameFromUrl(batch[batchIndex]);
          profiles.push({
            username,
            posts: '',
            followers: '',
            following: '',
            latestPost: {
              postUrl: '',
              thumbnailImage: '',
              likes: '',
              postingDate: '',
              postType: 'feed' as const,
              content: '',
              hashtags: '',
              mentions: '',
            },
            status: 'failed' as const,
            error: '배치 처리 중 예외 발생',
          });
        }
      });

      console.log(
        `배치 ${batchNumber} 완료: ${batchResults.length}개 처리됨 (총 ${profiles.length}개 완료)`,
      );

      // 배치 간 딜레이 (Instagram rate limiting 방지)
      if (i + batchSize < influencerUrls.length) {
        console.log('다음 배치 전 대기...');
        await this.delay(3000 + Math.random() * 2000);
      }
    }

    await browser.close();
    console.log(`\n=== 전체 크롤링 완료: ${profiles.length}개 프로필 처리 ===`);
    return profiles;
  }

  private extractUsernameFromUrl(url: string): string {
    try {
      const urlObj = new URL(url);
      const pathname = urlObj.pathname;
      const username = pathname.split('/')[1] || pathname.replace('/', '');
      return username;
    } catch (error) {
      const parts = url.split('/');
      return parts[parts.length - 1] || url;
    }
  }

  private async crawlSingleProfile(
    context: BrowserContext,
    url: string,
    username: string,
  ): Promise<InstagramProfile> {
    const startTime = Date.now();
    const page = await context.newPage();

    try {
      console.log(`📍 ${username} - 페이지 로딩 시작`);
      await page.goto(url, {
        waitUntil: 'domcontentloaded',
        timeout: 30000,
      });

      await page.waitForSelector('header', { timeout: 10000 });
      console.log(
        `📍 ${username} - 페이지 로딩 완료 (${Date.now() - startTime}ms)`,
      );

      // 1. 기본 프로필 정보 추출
      const actualUsername = await page
        .locator('header h2')
        .innerText()
        .catch(() => username);

      const stats = await this.extractDetailedStats(page);

      // 2. 타임스탬프 기반으로 최신 포스트 찾기 및 상세정보 추출
      const latestPostInfo = await this.extractLatestPostInfo(page, username);

      await page.close();

      const totalTime = Date.now() - startTime;
      console.log(
        `✅ ${username} - 전체 크롤링 완료! 총 소요시간: ${totalTime}ms`,
      );

      return {
        username: actualUsername || username,
        posts: stats.posts,
        followers: stats.followers,
        following: stats.following,
        latestPost: latestPostInfo,
        status: 'success',
      };
    } catch (error) {
      await page.close();
      console.log(
        `❌ ${username} - 크롤링 실패! 총 소요시간: ${Date.now() - startTime}ms`,
      );
      throw error;
    }
  }

  private async extractLatestPostInfo(
    page: Page,
    username: string,
  ): Promise<LatestPostInfo> {
    try {
      console.log(`${username} - 최신 포스트 정보 추출 시작`);

      // 포스트 컨테이너 대기
      await page.waitForSelector('div._aagu', { timeout: 10000 });

      // 모든 포스트 컨테이너에서 기본 정보와 타임스탬프 추출
      const postsWithTimestamp = await page.evaluate(() => {
        const postContainers = document.querySelectorAll('div._aagu');
        const posts: Array<{
          postUrl: string;
          thumbnailImage: string;
          isPinned: boolean;
          index: number;
        }> = [];

        postContainers.forEach((container, index) => {
          // 고정 게시물 확인
          const pinnedIcon = container.querySelector(
            'svg[aria-label="고정 게시물"]',
          );
          const isPinned = !!pinnedIcon;

          // 포스트 링크와 썸네일 찾기
          const linkElement = container.closest('a');
          const imgElement = container.querySelector('img');

          if (linkElement && linkElement.href) {
            posts.push({
              postUrl: linkElement.href,
              thumbnailImage: imgElement ? imgElement.src : '',
              isPinned,
              index,
            });
          }
        });

        return posts;
      });

      if (postsWithTimestamp.length === 0) {
        console.log(`${username} - 포스트를 찾을 수 없습니다`);
        return {
          postUrl: '',
          thumbnailImage: '',
          likes: '',
          postingDate: '',
          postType: 'feed',
          content: '',
          hashtags: '',
          mentions: '',
        };
      }

      console.log(
        `${username} - 총 ${postsWithTimestamp.length}개 포스트 발견`,
      );

      // 고정 게시물을 제외한 포스트들 필터링
      const nonPinnedPosts = postsWithTimestamp.filter(
        (post) => !post.isPinned,
      );

      if (nonPinnedPosts.length === 0) {
        console.log(`${username} - 고정되지 않은 포스트를 찾을 수 없습니다`);
        return {
          postUrl: '',
          thumbnailImage: '',
          likes: '',
          postingDate: '',
          postType: 'feed',
          content: '',
          hashtags: '',
          mentions: '',
        };
      }

      console.log(
        `${username} - 고정되지 않은 포스트 ${nonPinnedPosts.length}개 중에서 최신 포스트 찾기 (상위 3개만 확인)`,
      );

      // 각 포스트의 타임스탬프 추출하여 최신 포스트 찾기 - 상위 3개만 확인
      let latestPost = nonPinnedPosts[0];
      let latestTimestamp: Date | null = null;

      // 처음 3개만 체크하여 Instagram rate limiting 방지 및 성능 최적화
      const postsToCheck = nonPinnedPosts.slice(0, 5);
      console.log(
        `${username} - 상위 ${postsToCheck.length}개 포스트의 타임스탬프를 확인합니다`,
      );

      for (const post of postsToCheck) {
        try {
          const timestamp = await this.getPostTimestamp(
            page,
            post.postUrl,
            username,
          );
          if (timestamp) {
            const timestampDate = new Date(timestamp);
            if (!latestTimestamp || timestampDate > latestTimestamp) {
              latestTimestamp = timestampDate;
              latestPost = post;
            }
            console.log(
              `${username} - 포스트 ${post.index} 타임스탬프: ${timestamp}`,
            );
          }
        } catch (error) {
          console.error(
            `${username} - 포스트 ${post.index} 타임스탬프 추출 실패:`,
            error,
          );
        }
      }

      console.log(`${username} - 최신 포스트 선택: ${latestPost.postUrl}`);
      if (latestTimestamp) {
        console.log(
          `${username} - 선택된 포스트 날짜: ${latestTimestamp.toISOString()}`,
        );
      }

      // 선택된 최신 포스트의 상세 정보 추출
      const detailInfo = await this.extractPostDetailInNewTab(
        page,
        latestPost.postUrl,
        username,
      );

      return {
        postUrl: latestPost.postUrl,
        thumbnailImage: latestPost.thumbnailImage,
        ...detailInfo,
      };
    } catch (error) {
      console.error(`${username} - 최신 포스트 정보 추출 실패:`, error);
      return {
        postUrl: '',
        thumbnailImage: '',
        likes: '',
        postingDate: '',
        postType: 'feed',
        content: '',
        hashtags: '',
        mentions: '',
      };
    }
  }

  private async getPostTimestamp(
    mainPage: Page,
    postUrl: string,
    username: string,
  ): Promise<string | null> {
    let newPage: Page | null = null;

    try {
      const context = mainPage.context();
      newPage = await context.newPage();

      await newPage.goto(postUrl, {
        waitUntil: 'domcontentloaded',
        timeout: 15000,
      });
      await this.delay(1000);

      const timestamp = await newPage.evaluate(() => {
        // title 속성에서 타임스탬프 추출 우선
        const timeWithTitle = document.querySelector(
          'time[title]',
        ) as HTMLTimeElement;
        if (timeWithTitle && timeWithTitle.title) {
          return timeWithTitle.title;
        }

        // datetime 속성에서 추출
        const timeElement = document.querySelector(
          'time[datetime]',
        ) as HTMLTimeElement;
        if (timeElement && timeElement.dateTime) {
          return timeElement.dateTime;
        }

        return null;
      });

      return timestamp;
    } catch (error) {
      console.error(`${username} - 타임스탬프 추출 실패 (${postUrl}):`, error);
      return null;
    } finally {
      if (newPage) {
        try {
          await newPage.close();
        } catch (closeError) {
          console.error(`${username} - 타임스탬프 탭 닫기 실패:`, closeError);
        }
      }
    }
  }

  private async extractPostDetailInNewTab(
    mainPage: Page,
    postUrl: string,
    username: string,
  ): Promise<{
    likes: string;
    postingDate: string;
    postType: 'feed' | 'reels';
    content: string;
    hashtags: string;
    mentions: string;
  }> {
    let newPage: Page | null = null;

    try {
      console.log(`${username} - 새 탭에서 포스트 상세 정보 추출: ${postUrl}`);

      const context = mainPage.context();
      newPage = await context.newPage();

      await newPage.goto(postUrl, {
        waitUntil: 'domcontentloaded',
        timeout: 30000,
      });
      await this.delay(2000);

      // 릴스인 경우 추가 대기 및 상호작용
      if (postUrl.includes('/reel/')) {
        console.log(`${username} - 릴스 감지, 추가 로딩 대기 중...`);

        // 비디오 영역 클릭하여 캡션 로딩 유도
        try {
          await newPage.click('video', { timeout: 5000 });
        } catch (e) {
          // 비디오 클릭 실패해도 계속 진행
        }

        // 추가 대기 시간
        await this.delay(3000);

        // 스크롤을 통한 추가 로딩 유도
        await newPage.evaluate(() => {
          window.scrollBy(0, 100);
        });
        await this.delay(1000);
      }

      const postDetail = await newPage.evaluate(() => {
        const result = {
          likes: '',
          postingDate: '',
          postType: 'feed' as 'feed' | 'reels',
          content: '',
          hashtags: '',
          mentions: '',
        };

        // 좋아요 수 추출
        let likesCount = '';
        const likesSection = document.querySelector(
          'section div span[dir="auto"]',
        );
        if (likesSection && likesSection.textContent?.includes('좋아요')) {
          const likesText = likesSection.textContent;
          const numberMatch = likesText.match(/(\d+)/);
          if (numberMatch) {
            likesCount = numberMatch[1];
          }
        }

        if (!likesCount) {
          const likesLink = document.querySelector('a[href*="/liked_by/"]');
          if (likesLink) {
            const likesText = likesLink.textContent || '';
            const numberMatch = likesText.match(/(\d+)/);
            if (numberMatch) {
              likesCount = numberMatch[1];
            }
          }
        }
        result.likes = likesCount;

        // 포스팅 날짜 추출 (title 속성 우선)
        const timeWithTitle = document.querySelector(
          'time[title]',
        ) as HTMLTimeElement;
        if (timeWithTitle && timeWithTitle.title) {
          result.postingDate = timeWithTitle.title;
        } else {
          const timeElement = document.querySelector(
            'time[datetime]',
          ) as HTMLTimeElement;
          if (timeElement && timeElement.dateTime) {
            result.postingDate = timeElement.dateTime;
          }
        }

        // 포스트 타입 확인 (릴스인지 피드인지)
        if (window.location.href.includes('/reel/')) {
          result.postType = 'reels';
        }

        // 내용 추출 - 다양한 선택자로 시도
        let content = '';

        // 1차 시도: 포스트 타입별 선택자
        const isReels = window.location.href.includes('/reel/');
        let contentSelectors: string[] = [];

        if (isReels) {
          // 릴스용 선택자
          contentSelectors = [
            'h1._ap3a',
            'h1[dir="auto"]',
            '._ap3a._aaco._aacu._aacx._aad7._aade',
            '[data-testid="reels-caption"] span',
            'article section div span[dir="auto"]',
            'div[role="button"] span[dir="auto"]',
            'span[dir="auto"]',
            'div[style*="line-height"] span',
          ];
        } else {
          // 일반 피드용 선택자
          contentSelectors = [
            '[data-testid="post-content"] span',
            'article span',
            'h1._ap3a',
            'h1[dir="auto"]',
            '._ap3a._aaco._aacu._aacx._aad7._aade',
            'article div span[dir="auto"]',
            'section div span[dir="auto"]',
          ];
        }

        for (const selector of contentSelectors) {
          const elements = document.querySelectorAll(selector);
          for (const element of elements) {
            const text = element.textContent?.trim();
            // 에러 메시지 및 불필요한 텍스트 제외
            if (
              text &&
              text.length > 20 &&
              !text.includes('좋아요') &&
              !text.includes('댓글') &&
              !text.includes('시간') &&
              !text.includes("Sorry, we're having trouble") &&
              !text.includes('This video is unavailable') &&
              !text.includes('Video unavailable')
            ) {
              content = text;
              break;
            }
          }
          if (content) break;
        }

        // 2차 시도: 모든 텍스트 요소에서 긴 텍스트 찾기
        if (!content) {
          const allElements = document.querySelectorAll('span, h1, div, p');
          const candidateTexts: string[] = [];

          for (const element of allElements) {
            const text = element.textContent?.trim();
            if (
              text &&
              text.length > 30 &&
              !text.includes("Sorry, we're having trouble") &&
              !text.includes('좋아요') &&
              !text.includes('팔로우') &&
              !text.includes('시간') &&
              element.children.length === 0
            ) {
              // 자식 요소가 없는 leaf 노드만
              candidateTexts.push(text);
            }
          }

          // 가장 긴 텍스트를 콘텐츠로 선택
          if (candidateTexts.length > 0) {
            content = candidateTexts.reduce((longest, current) =>
              current.length > longest.length ? current : longest,
            );
          }
        }

        result.content = content;

        // 해시태그와 멘션 분류
        if (content) {
          const hashtags = content.match(/#[\w가-힣]+/g) || [];
          const mentions = content.match(/@[\w가-힣.]+/g) || [];

          result.hashtags = hashtags.join(', ');
          result.mentions = mentions.join(', ');
        }

        return result;
      });

      console.log(`${username} - 포스트 상세 정보 추출 완료:`, postDetail);
      return postDetail;
    } catch (error) {
      console.error(`${username} - 포스트 상세 정보 추출 실패:`, error);
      return {
        likes: '',
        postingDate: '',
        postType: 'feed',
        content: '',
        hashtags: '',
        mentions: '',
      };
    } finally {
      if (newPage) {
        try {
          await newPage.close();
          console.log(`${username} - 포스트 상세 정보 탭 정리 완료`);
        } catch (closeError) {
          console.error(`${username} - 탭 닫기 실패:`, closeError);
        }
      }
    }
  }

  private cleanStatText(text: string): string {
    return text
      .replace(/posts?/gi, '')
      .replace(/followers?/gi, '')
      .replace(/following/gi, '')
      .replace(/게시물/g, '')
      .replace(/팔로워/g, '')
      .replace(/팔로잉/g, '')
      .replace(/팔로우/g, '')
      .trim();
  }

  private async saveToExcel(
    profiles: InstagramProfile[],
    filePath: string,
  ): Promise<void> {
    try {
      console.log('\n=== 엑셀 파일 저장 시작 ===');

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Instagram Profiles Enhanced');

      worksheet.columns = [
        { header: '사용자명', key: 'username', width: 20 },
        { header: '게시물 수', key: 'posts', width: 15 },
        { header: '팔로워 수', key: 'followers', width: 15 },
        { header: '팔로잉 수', key: 'following', width: 15 },
        { header: '포스트 URL', key: 'postUrl', width: 60 },
        { header: '썸네일 이미지', key: 'thumbnailImage', width: 25 },
        { header: '좋아요 수', key: 'likes', width: 15 },
        { header: '포스팅 날짜', key: 'postingDate', width: 20 },
        { header: '포스트 형식', key: 'postType', width: 15 },
        { header: '내용', key: 'content', width: 80 },
        { header: '해시태그', key: 'hashtags', width: 50 },
        { header: '멘션', key: 'mentions', width: 30 },
        { header: '상태', key: 'status', width: 10 },
        { header: '오류 메시지', key: 'error', width: 30 },
      ];

      // 헤더 행 스타일 적용
      const headerRow = worksheet.getRow(1);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '366092' },
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });

      // 데이터 행 추가 - 이미지 다운로드를 순차 처리
      for (let index = 0; index < profiles.length; index++) {
        const profile = profiles[index];

        const row = worksheet.addRow({
          username: profile.username,
          posts: this.cleanStatText(profile.posts),
          followers: this.cleanStatText(profile.followers),
          following: this.cleanStatText(profile.following),
          postUrl: profile.latestPost.postUrl,
          thumbnailImage: profile.latestPost.thumbnailImage, // 기본값으로 URL 설정
          likes: profile.latestPost.likes,
          postingDate: profile.latestPost.postingDate,
          postType: profile.latestPost.postType,
          content: profile.latestPost.content,
          hashtags: profile.latestPost.hashtags,
          mentions: profile.latestPost.mentions,
          status: profile.status,
          error: profile.error || '',
        });

        // HttpService를 사용한 이미지 다운로드 및 삽입
        if (profile.latestPost.thumbnailImage) {
          try {
            console.log(
              `이미지 다운로드 중 (${index + 1}/${profiles.length}): ${profile.latestPost.thumbnailImage}`,
            );

            // HttpService로 이미지 다운로드
            const response = await firstValueFrom(
              this.httpService.get(profile.latestPost.thumbnailImage, {
                responseType: 'arraybuffer',
                timeout: 10000,
                headers: {
                  'User-Agent':
                    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                },
              }),
            );

            const buffer = Buffer.from(response.data);

            // 이미지 타입 확인 및 extension 설정
            let extension: 'jpeg' | 'png' | 'gif' = 'jpeg';
            const contentType = response.headers['content-type'];
            if (contentType?.includes('png')) {
              extension = 'png';
            } else if (contentType?.includes('gif')) {
              extension = 'gif';
            }

            const imageId = workbook.addImage({
              buffer: buffer,
              extension: extension,
            });

            worksheet.addImage(imageId, {
              tl: { col: 5, row: row.number - 1 }, // F열 (썸네일 이미지 컬럼, 0-based)
              ext: { width: 100, height: 100 },
              editAs: 'oneCell',
            });

            // 이미지 셀은 빈 값으로 설정
            const imageCell = worksheet.getCell(row.number, 6);
            imageCell.value = '';

            console.log(
              `이미지 삽입 성공: ${profile.latestPost.thumbnailImage}`,
            );
          } catch (imageError) {
            console.error(
              `이미지 삽입 실패 (${profile.latestPost.thumbnailImage}):`,
              imageError.message,
            );
            // 실패시 URL만 텍스트로 삽입
            const imageCell = worksheet.getCell(row.number, 6);
            imageCell.value = profile.latestPost.thumbnailImage;
          }
        }

        // 데이터 행 스타일 적용
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
          cell.alignment = { vertical: 'middle', wrapText: true };
        });

        // 홀수/짝수 행 배경색 구분
        if (index % 2 === 1) {
          row.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'F8F9FA' },
            };
          });
        }

        // 이미지 표시를 위한 행 높이 조정
        worksheet.getRow(row.number).height = 60;
      }

      // 숫자 컬럼 정렬 (게시물수, 팔로워수, 팔로잉수, 좋아요수)
      const numberColumns = [2, 3, 4, 7]; // B, C, D, G 컬럼
      numberColumns.forEach((colNum) => {
        worksheet.getColumn(colNum).alignment = {
          horizontal: 'right',
          vertical: 'middle',
        };
      });

      await workbook.xlsx.writeFile(filePath);

      console.log(`엑셀 파일 저장 완료: ${filePath}`);
      console.log(`총 ${profiles.length}개의 프로필 데이터가 저장되었습니다.`);

      // 저장된 파일 정보 출력
      const fileStats = await fs.promises.stat(filePath);
      console.log(`파일 크기: ${(fileStats.size / 1024).toFixed(2)} KB`);
    } catch (error) {
      console.error('엑셀 파일 저장 중 오류:', error);
      throw error;
    }
  }

  private async extractDetailedStats(
    page: Page,
  ): Promise<{ posts: string; followers: string; following: string }> {
    try {
      const allStats = await page
        .locator('header ul li span')
        .allInnerTexts()
        .catch(() => ['', '', '']);

      let followersText = '';
      try {
        const followersWithTitle = await page
          .locator('a[href*="/followers/"] span[title]')
          .getAttribute('title');
        if (followersWithTitle) {
          followersText = followersWithTitle;
          console.log(
            `📊 정확한 팔로워 수 발견 (title 속성): ${followersText}`,
          );
        } else {
          followersText =
            allStats.find(
              (text) => text.includes('팔로워') || text.includes('followers'),
            ) || '';
        }
      } catch (error) {
        followersText =
          allStats.find(
            (text) => text.includes('팔로워') || text.includes('followers'),
          ) || '';
      }

      const result = {
        posts:
          allStats.find(
            (text) => text.includes('게시물') || text.includes('posts'),
          ) || '',
        followers: followersText,
        following:
          allStats.find(
            (text) => text.includes('팔로우') || text.includes('following'),
          ) || '',
      };

      console.log('✅ 추출된 통계:', result);
      return result;
    } catch (error) {
      console.error('통계 추출 중 오류:', error);
      return { posts: '', followers: '', following: '' };
    }
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
