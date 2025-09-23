import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { HttpService } from '@nestjs/axios';
import { firstValueFrom } from 'rxjs';
import * as fs from 'fs/promises';
import * as path from 'path';
import FormData from 'form-data';

interface GoogleToken {
  access_token: string;
  expires_in: number;
  token_type: string;
  scope: string;
}

interface UploadResult {
  id: string;
  name: string;
  mimeType: string;
  webViewLink: string;
  webContentLink: string;
}

@Injectable()
export class GoogleDriveService {
  private readonly logger = new Logger(GoogleDriveService.name);
  private readonly DRIVE_UPLOAD_URL =
    'https://www.googleapis.com/upload/drive/v3/files';
  private readonly DRIVE_API_URL = 'https://www.googleapis.com/drive/v3';

  private accessToken: string | null = null;
  private tokenExpiresAt: number = 0;
  private targetFolderId: string | null = null;
  private readonly TARGET_FOLDER_NAME = 'Instagram Crawl Data'; // 원하는 폴더명으로 변경 가능

  constructor(
    private readonly configService: ConfigService,
    private readonly httpService: HttpService,
  ) {}

  /**
   * 타겟 폴더 확보 (없으면 생성, 있으면 찾기)
   */
  private async ensureTargetFolder(): Promise<string> {
    if (this.targetFolderId) {
      return this.targetFolderId;
    }

    try {
      // 1. 기존 폴더 검색
      const existingFolder = await this.findFolderByName(
        this.TARGET_FOLDER_NAME,
      );

      if (existingFolder) {
        this.targetFolderId = existingFolder.id;
        this.logger.log(
          `기존 폴더 사용: ${this.TARGET_FOLDER_NAME} (ID: ${this.targetFolderId})`,
        );
        return this.targetFolderId;
      }

      // 2. 폴더가 없으면 새로 생성
      this.logger.log(`폴더 생성 중: ${this.TARGET_FOLDER_NAME}`);
      const newFolder = await this.createFolder(this.TARGET_FOLDER_NAME);
      this.targetFolderId = newFolder.folderId;

      this.logger.log(
        `새 폴더 생성 완료: ${this.TARGET_FOLDER_NAME} (ID: ${this.targetFolderId})`,
      );
      this.logger.log(`폴더 URL: ${newFolder.folderUrl}`);

      return this.targetFolderId;
    } catch (error) {
      this.logger.error('타겟 폴더 설정 실패:', error);
      throw new Error(`타겟 폴더 설정 실패: ${error.message}`);
    }
  }

  /**
   * 이름으로 폴더 검색
   */
  private async findFolderByName(
    folderName: string,
  ): Promise<{ id: string; name: string } | null> {
    try {
      const accessToken = await this.getAccessToken();

      const response = await firstValueFrom(
        this.httpService.get(
          `${this.DRIVE_API_URL}/files?q=name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false&fields=files(id,name)&supportsAllDrives=true`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
            },
          },
        ),
      );

      const folders = response.data.files;
      return folders && folders.length > 0 ? folders[0] : null;
    } catch (error) {
      this.logger.error('폴더 검색 실패:', error);
      return null;
    }
  }

  async getAccessToken(): Promise<string> {
    // 토큰이 유효하면 기존 토큰 반환
    if (this.accessToken && Date.now() < this.tokenExpiresAt) {
      return this.accessToken;
    }

    this.logger.log('새로운 액세스 토큰 요청 중...');

    const url = 'https://oauth2.googleapis.com/token';
    const requestBody = new URLSearchParams({
      grant_type: 'refresh_token',
      client_id: this.configService.get<string>('CLIENT_ID'),
      client_secret: this.configService.get<string>('CLIENT_SECRET'),
      refresh_token: this.configService.get<string>('REFRESH_TOKEN'),
    } as Record<string, string>);

    try {
      const { data } = await firstValueFrom(
        this.httpService.post<GoogleToken>(url, requestBody, {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        }),
      );

      this.accessToken = data.access_token;
      this.tokenExpiresAt = Date.now() + (data.expires_in - 300) * 1000; // 5분 여유

      this.logger.log('액세스 토큰 갱신 완료');
      return data.access_token;
    } catch (error) {
      this.logger.error('토큰 갱신 실패:', error.response?.data);
      throw new Error(
        `토큰 갱신 실패: ${error.response?.data?.error_description || error.message}`,
      );
    }
  }

  /**
   * 로컬 엑셀 파일을 타겟 폴더에 업로드
   */
  async uploadExcelFile(
    filePath: string = 'data/instagram_profiles_result.xlsx',
  ): Promise<{
    fileId: string;
    shareableUrl: string;
    fileName: string;
    folderUrl: string;
  }> {
    try {
      this.logger.log(`엑셀 파일 업로드 시작: ${filePath}`);

      // 파일 존재 확인
      await fs.access(filePath);

      // 타겟 폴더 확보
      const targetFolderId = await this.ensureTargetFolder();

      // 파일 정보
      const originalFileName = path.basename(filePath);
      const timestamp = new Date()
        .toISOString()
        .replace(/[:.]/g, '-')
        .slice(0, -5); // YYYY-MM-DDTHH-MM-SS
      const fileName = `${timestamp}_${originalFileName}`; // 타임스탬프 추가로 중복 방지
      const fileBuffer = await fs.readFile(filePath);
      const mimeType =
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

      // 액세스 토큰 획득
      const accessToken = await this.getAccessToken();

      // 파일 메타데이터
      const metadata = {
        name: fileName,
        mimeType,
        parents: [targetFolderId],
      };

      // FormData 생성 (multipart upload)
      const formData = new FormData();
      formData.append('metadata', JSON.stringify(metadata), {
        contentType: 'application/json; charset=UTF-8',
      });
      formData.append('file', fileBuffer, {
        filename: fileName,
        contentType: mimeType,
      });

      // 파일 업로드
      const uploadResponse = await firstValueFrom(
        this.httpService.post<UploadResult>(
          `${this.DRIVE_UPLOAD_URL}?uploadType=multipart&supportsAllDrives=true`,
          formData,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              ...formData.getHeaders(),
            },
          },
        ),
      );

      const fileId = uploadResponse.data.id;
      this.logger.log(`파일 업로드 완료. File ID: ${fileId}`);

      // 파일 권한 설정 (URL을 가진 모든 사용자가 읽을 수 있도록)
      await this.setFilePublicPermission(fileId);

      // URL 생성
      const shareableUrl = `https://drive.google.com/file/d/${fileId}/view`;
      const folderUrl = `https://drive.google.com/drive/folders/${targetFolderId}`;

      this.logger.log(`파일 공유 URL: ${shareableUrl}`);
      this.logger.log(`폴더 URL: ${folderUrl}`);

      return {
        fileId,
        shareableUrl,
        fileName,
        folderUrl,
      };
    } catch (error) {
      this.logger.error('엑셀 파일 업로드 실패:', error);

      if (error.code === 'ENOENT') {
        throw new Error(`파일을 찾을 수 없습니다: ${filePath}`);
      }

      throw new Error(
        `파일 업로드 실패: ${error.response?.data?.error?.message || error.message}`,
      );
    }
  }

  /**
   * 파일을 public으로 설정 (URL을 가진 모든 사용자가 읽기 가능)
   */
  private async setFilePublicPermission(fileId: string): Promise<void> {
    try {
      const accessToken = await this.getAccessToken();

      await firstValueFrom(
        this.httpService.post(
          `${this.DRIVE_API_URL}/files/${fileId}/permissions`,
          {
            role: 'reader',
            type: 'anyone',
          },
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
          },
        ),
      );

      this.logger.log('파일 공개 권한 설정 완료');
    } catch (error) {
      this.logger.error('권한 설정 실패:', error.response?.data);
      throw new Error(
        `권한 설정 실패: ${error.response?.data?.error?.message || error.message}`,
      );
    }
  }

  /**
   * 폴더 생성
   */
  async createFolder(
    folderName: string,
    parentFolderId?: string,
  ): Promise<{ folderId: string; folderUrl: string }> {
    try {
      const accessToken = await this.getAccessToken();

      const metadata = {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
        ...(parentFolderId && { parents: [parentFolderId] }),
      };

      const response = await firstValueFrom(
        this.httpService.post<UploadResult>(
          `${this.DRIVE_API_URL}/files?supportsAllDrives=true`,
          metadata,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
          },
        ),
      );

      const folderId = response.data.id;

      // 폴더도 공개 설정
      await this.setFilePublicPermission(folderId);

      const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

      this.logger.log(`폴더 생성 완료: ${folderName}`);

      return {
        folderId,
        folderUrl,
      };
    } catch (error) {
      this.logger.error('폴더 생성 실패:', error);
      throw new Error(
        `폴더 생성 실패: ${error.response?.data?.error?.message || error.message}`,
      );
    }
  }

  /**
   * 파일 정보 조회
   */
  async getFileInfo(fileId: string): Promise<UploadResult> {
    try {
      const accessToken = await this.getAccessToken();

      const response = await firstValueFrom(
        this.httpService.get<UploadResult>(
          `${this.DRIVE_API_URL}/files/${fileId}?fields=id,name,mimeType,webViewLink,webContentLink&supportsAllDrives=true`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
            },
          },
        ),
      );

      return response.data;
    } catch (error) {
      this.logger.error('파일 정보 조회 실패:', error);
      throw new Error(
        `파일 정보 조회 실패: ${error.response?.data?.error?.message || error.message}`,
      );
    }
  }

  /**
   * 한번에 실행하는 메인 함수
   */
  async uploadInstagramProfilesFile(): Promise<{
    success: boolean;
    fileId?: string;
    shareableUrl?: string;
    fileName?: string;
    folderUrl?: string;
    error?: string;
  }> {
    try {
      this.logger.log('=== Instagram 프로필 파일 업로드 시작 ===');

      const result = await this.uploadExcelFile();

      this.logger.log('✅ 업로드 성공!');
      this.logger.log(`파일명: ${result.fileName}`);
      this.logger.log(`파일 공유 URL: ${result.shareableUrl}`);
      this.logger.log(`폴더 URL: ${result.folderUrl}`);
      this.logger.log('=== 업로드 완료 ===');

      return {
        success: true,
        fileId: result.fileId,
        shareableUrl: result.shareableUrl,
        fileName: result.fileName,
        folderUrl: result.folderUrl,
      };
    } catch (error) {
      this.logger.error('❌ 업로드 실패:', error.message);

      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 타겟 폴더 정보 조회
   */
  async getTargetFolderInfo(): Promise<{
    folderId: string;
    folderUrl: string;
    folderName: string;
  }> {
    const folderId = await this.ensureTargetFolder();
    return {
      folderId,
      folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
      folderName: this.TARGET_FOLDER_NAME,
    };
  }

  /**
   * 폴더 내 파일 목록 조회
   */
  async listFilesInTargetFolder(): Promise<
    Array<{
      id: string;
      name: string;
      webViewLink: string;
      createdTime: string;
    }>
  > {
    try {
      const folderId = await this.ensureTargetFolder();
      const accessToken = await this.getAccessToken();

      const response = await firstValueFrom(
        this.httpService.get(
          `${this.DRIVE_API_URL}/files?q='${folderId}' in parents and trashed=false&fields=files(id,name,webViewLink,createdTime)&orderBy=createdTime desc&supportsAllDrives=true`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
            },
          },
        ),
      );

      return response.data.files || [];
    } catch (error) {
      this.logger.error('파일 목록 조회 실패:', error);
      throw new Error(
        `파일 목록 조회 실패: ${error.response?.data?.error?.message || error.message}`,
      );
    }
  }
}
