import { publicProcedure, router } from "../_core/trpc";
import { z } from "zod";
import fs from "fs";
import path from "path";
import os from "os";
import { createReadStream, createWriteStream } from "fs";
import { pipeline } from "stream/promises";
import archiver from "archiver";

/**
 * 臨時檔案管理器
 * 負責創建、追蹤和清理臨時檔案
 */
class TempFileManager {
  private tempDir: string;
  private trackedFiles: Set<string> = new Set();

  constructor() {
    this.tempDir = path.join(os.tmpdir(), "ticketsplitter");
    // 確保臨時目錄存在
    if (!fs.existsSync(this.tempDir)) {
      fs.mkdirSync(this.tempDir, { recursive: true });
    }
  }

  /**
   * 創建臨時檔案路徑
   */
  createTempFilePath(originalFilename: string): string {
    const timestamp = Date.now();
    const random = Math.random().toString(36).substring(7);
    const ext = path.extname(originalFilename);
    const filename = `${timestamp}-${random}${ext}`;
    const filepath = path.join(this.tempDir, filename);
    this.trackedFiles.add(filepath);
    return filepath;
  }

  /**
   * 創建臨時目錄用於存放分割後的檔案
   */
  createTempDir(): string {
    const timestamp = Date.now();
    const random = Math.random().toString(36).substring(7);
    const dirname = `split-${timestamp}-${random}`;
    const dirpath = path.join(this.tempDir, dirname);
    fs.mkdirSync(dirpath, { recursive: true });
    this.trackedFiles.add(dirpath);
    return dirpath;
  }

  /**
   * 清理所有追蹤的臨時檔案和目錄
   */
  async cleanup(): Promise<void> {
    for (const filepath of Array.from(this.trackedFiles)) {
      try {
        const stats = fs.statSync(filepath);
        if (stats.isDirectory()) {
          fs.rmSync(filepath, { recursive: true, force: true });
        } else {
          fs.unlinkSync(filepath);
        }
      } catch (error) {
        console.error(`[TempFileManager] Failed to cleanup ${filepath}:`, error);
      }
    }
    this.trackedFiles.clear();
  }

  /**
   * 獲取臨時目錄路徑
   */
  getTempDir(): string {
    return this.tempDir;
  }
}

/**
 * 檔案分割路由
 * 
 * 流程：
 * 1. 接收上傳的檔案和航空公司選擇
 * 2. 保存到臨時位置
 * 3. 調用對應的 Python 腳本進行分割
 * 4. 將分割結果打包成 ZIP
 * 5. 返回 ZIP 檔案供下載
 * 6. 自動清理所有臨時檔案
 */
export const splitRouter = router({
  /**
   * 上傳並分割檔案
   * 
   * 預期的 Python 腳本函數簽名：
   * - split_lion_air(input_file_path: str, output_dir: str) -> List[str]
   * - split_tiger_air(input_file_path: str, output_dir: str) -> List[str]
   * - split_scoot_air(input_file_path: str, output_dir: str) -> List[str]
   * 
   * 每個函數應返回輸出檔案路徑列表
   */
  split: publicProcedure
    .input(
      z.object({
        airline: z.enum(["SL", "IT", "BZ"]), // SL: 泰獅航, IT: 台灣虎航, BZ: 酷航
        filename: z.string().min(1),
        fileBuffer: z.instanceof(Buffer),
      })
    )
    .mutation(async ({ input }) => {
      const tempManager = new TempFileManager();

      try {
        // 1. 保存上傳的檔案到臨時位置
        const uploadedFilePath = tempManager.createTempFilePath(input.filename);
        fs.writeFileSync(uploadedFilePath, input.fileBuffer);

        // 2. 創建輸出目錄
        const outputDir = tempManager.createTempDir();

        // 3. 調用對應的 Python 腳本
        // TODO: 整合 Python 腳本
        // 根據航空公司選擇調用對應的分割函數
        // const splitFunction = getSplitFunction(input.airline);
        // const outputFiles = await splitFunction(uploadedFilePath, outputDir);

        // 暫時使用模擬輸出（待 Python 腳本集成）
        const outputFiles: string[] = [];

        // 4. 創建 ZIP 檔案
        const zipPath = tempManager.createTempFilePath("split-result.zip");
        await createZipArchive(outputDir, zipPath, outputFiles);

        // 5. 讀取 ZIP 檔案內容
        const zipBuffer = fs.readFileSync(zipPath);

        // 6. 返回 ZIP 檔案內容和元數據
        return {
          success: true,
          zipBuffer: zipBuffer.toString("base64"), // 轉換為 base64 以便傳輸
          filename: `ticket-split-${input.airline}-${Date.now()}.zip`,
          size: zipBuffer.length,
        };
      } catch (error) {
        console.error("[SplitRouter] Error during file split:", error);
        throw new Error(
          `檔案分割失敗: ${error instanceof Error ? error.message : "未知錯誤"}`
        );
      } finally {
        // 7. 清理所有臨時檔案（無論成功或失敗）
        await tempManager.cleanup();
      }
    }),

  /**
   * 健康檢查端點
   */
  health: publicProcedure.query(() => ({
    status: "ok",
    tempDir: path.join(os.tmpdir(), "ticketsplitter"),
  })),
});

/**
 * 創建 ZIP 檔案
 */
async function createZipArchive(
  sourceDir: string,
  zipPath: string,
  files: string[]
): Promise<void> {
  return new Promise((resolve, reject) => {
    const output = createWriteStream(zipPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => resolve());
    archive.on("error", reject);

    archive.pipe(output);

    // 如果有指定的檔案列表，只添加這些檔案
    if (files.length > 0) {
      for (const file of files) {
        if (fs.existsSync(file)) {
          const filename = path.basename(file);
          archive.file(file, { name: filename });
        }
      }
    } else {
      // 否則添加整個目錄
      archive.directory(sourceDir, false);
    }

    archive.finalize();
  });
}

/**
 * 獲取對應航空公司的分割函數
 * 
 * 待實現：集成 Python 腳本
 * 
 * 預期的 Python 模組結構：
 * ```
 * server/scripts/
 *   ├── lion_air_splitter.py
 *   ├── tiger_air_splitter.py
 *   └── scoot_air_splitter.py
 * ```
 */
function getSplitFunction(airline: string) {
  // TODO: 實現 Python 腳本集成
  // 使用 child_process 或 Python 子進程調用
  const splitFunctions: Record<string, Function> = {
    SL: async (input: string, output: string) => {
      // 調用 lion_air_splitter.split_lion_air(input, output)
      console.log(`[SplitRouter] Splitting Lion Air file: ${input}`);
      return [];
    },
    IT: async (input: string, output: string) => {
      // 調用 tiger_air_splitter.split_tiger_air(input, output)
      console.log(`[SplitRouter] Splitting Tiger Air file: ${input}`);
      return [];
    },
    BZ: async (input: string, output: string) => {
      // 調用 scoot_air_splitter.split_scoot_air(input, output)
      console.log(`[SplitRouter] Splitting Scoot Air file: ${input}`);
      return [];
    },
  };

  return splitFunctions[airline] || splitFunctions.SL;
}
