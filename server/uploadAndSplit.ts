import type { Express } from "express";
import fs from "fs";
import path from "path";
import os from "os";
import { spawn } from "child_process";
import archiver from "archiver";
import multer from "multer";
import { nanoid } from "nanoid";

/** 單次任務上下文：打包 ZIP 時依各航空策略抓取檔案 */
export type OutputGrabContext = {
  jobDir: string;
  outputDir: string;
};

export type CustomOutputGrab = (ctx: OutputGrabContext) => string[];

export type AirlineScriptConfig = {
  scriptName: string;
  requiresOutputDir: boolean;
  customOutputGrab: CustomOutputGrab;
};

function collectFilesRecursive(dir: string): string[] {
  const results: string[] = [];
  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return results;
  }
  for (const ent of entries) {
    const full = path.join(dir, ent.name);
    if (ent.isDirectory()) {
      results.push(...collectFilesRecursive(full));
    } else if (ent.isFile()) {
      results.push(full);
    }
  }
  return results;
}

/** 僅 .pdf，且檔名（不含副檔名）不以 input 開頭（不區分大小寫） */
function grabPdfsIgnoreInputPrefix(ctx: OutputGrabContext): string[] {
  const { outputDir } = ctx;
  const all = collectFilesRecursive(outputDir);
  return all.filter(p => {
    const ext = path.extname(p).toLowerCase();
    if (ext !== ".pdf") return false;
    const base = path.basename(p, ext);
    if (/^input/i.test(base)) return false;
    return true;
  });
}

/** 虎航：抓取 outputDir/input/（或 jobDir/input/）底下所有檔案 */
function grabTigerInputFolder(ctx: OutputGrabContext): string[] {
  const underOut = path.join(ctx.outputDir, "input");
  if (fs.existsSync(underOut)) {
    return collectFilesRecursive(underOut);
  }
  const underJob = path.join(ctx.jobDir, "input");
  if (fs.existsSync(underJob)) {
    return collectFilesRecursive(underJob);
  }
  return [];
}

/**
 * 各航空 Python 腳本註冊表（可無限擴充）。
 * - requiresOutputDir: true → `python script.py <input> -o <outputDir>`
 * - requiresOutputDir: false → `python script.py <input>`，cwd = outputDir
 */
export const ScriptRegistry: Record<string, AirlineScriptConfig> = {
  thailionair: {
    scriptName: "split_thailionair.py",
    requiresOutputDir: true,
    customOutputGrab: grabPdfsIgnoreInputPrefix,
  },
  tigerair: {
    scriptName: "split_tigerair_docx.py",
    requiresOutputDir: false,
    customOutputGrab: grabTigerInputFolder,
  },
  scoot: {
    scriptName: "split_ticket_by_name.py",
    requiresOutputDir: true,
    customOutputGrab: grabPdfsIgnoreInputPrefix,
  },
  airasia: {
    scriptName: "split_airasia_ticket.py",
    requiresOutputDir: true,
    customOutputGrab: grabPdfsIgnoreInputPrefix,
  },
};

function scriptAbsolutePath(scriptName: string): string {
  if (process.env.TICKETSPLITTER_SCRIPTS_DIR) {
    return path.join(path.resolve(process.env.TICKETSPLITTER_SCRIPTS_DIR), scriptName);
  }
  return path.join(process.cwd(), "server", "scripts", scriptName);
}

const ALLOWED_EXT = new Set([".pdf", ".doc", ".docx", ".jpg", ".jpeg", ".png"]);
const ALLOWED_MIME = new Set([
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "image/jpeg",
  "image/png",
]);

function safeDiskFilename(originalname: string): string {
  const ext = path.extname(path.basename(originalname)).toLowerCase();
  const suffix = ALLOWED_EXT.has(ext) ? ext : "";
  return `${nanoid(18)}${suffix || ".upload"}`;
}

function getPythonExecutable(): string {
  if (process.env.PYTHON_PATH) return process.env.PYTHON_PATH;
  return process.platform === "win32" ? "python" : "python3";
}

function runPythonScript(
  scriptPath: string,
  args: string[],
  opts?: { cwd?: string }
): Promise<void> {
  const py = getPythonExecutable();
  return new Promise((resolve, reject) => {
    const child = spawn(py, [scriptPath, ...args], {
      stdio: ["ignore", "pipe", "pipe"],
      cwd: opts?.cwd,
      env: { ...process.env, PYTHONUNBUFFERED: "1" },
    });
    let stderr = "";
    child.stderr?.setEncoding("utf8");
    child.stderr?.on("data", (chunk: string) => {
      stderr += chunk;
      if (stderr.length > 32_000) stderr = stderr.slice(-16_000);
    });
    child.stdout?.resume();
    child.on("error", reject);
    child.on("close", code => {
      if (code === 0) resolve();
      else reject(new Error(stderr.trim() || `Python 腳本結束碼 ${code}`));
    });
  });
}

function rmJobDir(jobDir: string): void {
  try {
    if (jobDir && fs.existsSync(jobDir)) {
      fs.rmSync(jobDir, { recursive: true, force: true });
    }
  } catch (e) {
    console.error("[upload-and-split] Failed to remove temp dir:", jobDir, e);
  }
}

function addFilesToArchive(
  archive: ReturnType<typeof archiver>,
  files: string[],
  zipRoot: string
): void {
  for (const absPath of files) {
    const rel = path.relative(zipRoot, absPath);
    const name = rel.startsWith("..") ? path.basename(absPath) : rel;
    archive.file(absPath, { name: name.split(path.sep).join("/") });
  }
}

/**
 * POST /api/upload-and-split
 * multipart: field `file`, form field `airline`（ScriptRegistry 的 key，如 thailionair、tigerair、scoot、airasia）
 */
export function registerUploadAndSplitRoute(app: Express): void {
  app.post("/api/upload-and-split", (req, res) => {
    let jobDir: string;
    try {
      jobDir = fs.mkdtempSync(path.join(os.tmpdir(), "ticketsplitter-"));
    } catch (e) {
      const msg = e instanceof Error ? e.message : "無法建立暫存目錄";
      res.status(500).json({ error: msg });
      return;
    }

    const destroyJob = () => rmJobDir(jobDir);

    const uploadMw = multer({
      storage: multer.diskStorage({
        destination: (_req, _file, cb) => cb(null, jobDir),
        filename: (_req, file, cb) => cb(null, safeDiskFilename(file.originalname)),
      }),
      limits: { fileSize: 50 * 1024 * 1024 },
      fileFilter: (_req, file, cb) => {
        const ext = path.extname(path.basename(file.originalname)).toLowerCase();
        const mimeOk = ALLOWED_MIME.has(file.mimetype);
        const extOk = ALLOWED_EXT.has(ext);
        if (mimeOk || extOk) return cb(null, true);
        cb(new Error("不支援的檔案格式。請上傳 PDF、Word 或圖片檔案。"));
      },
    }).single("file");

    uploadMw(req, res, (err: unknown) => {
      void (async () => {
        if (err) {
          destroyJob();
          const message =
            err instanceof multer.MulterError
              ? err.code === "LIMIT_FILE_SIZE"
                ? "檔案過大。請上傳小於 50MB 的檔案。"
                : err.message
              : err instanceof Error
                ? err.message
                : "上傳失敗";
          if (!res.headersSent) {
            res.status(400).json({ error: message });
          }
          return;
        }

        try {
          const file = req.file;
          const airlineRaw = req.body?.airline;
          const airlineKey =
            typeof airlineRaw === "string" ? airlineRaw.trim().toLowerCase() : "";

          if (!file) {
            destroyJob();
            if (!res.headersSent) res.status(400).json({ error: "缺少檔案" });
            return;
          }

          const scriptConfig = ScriptRegistry[airlineKey];
          if (!scriptConfig) {
            destroyJob();
            if (!res.headersSent) {
              const valid = Object.keys(ScriptRegistry).join("、");
              res.status(400).json({ error: `無效的航空公司代碼。請使用：${valid}` });
            }
            return;
          }

          const scriptPath = scriptAbsolutePath(scriptConfig.scriptName);
          if (!fs.existsSync(scriptPath)) {
            destroyJob();
            if (!res.headersSent) {
              res.status(500).json({ error: `找不到分割腳本：${scriptConfig.scriptName}` });
            }
            return;
          }

          const inputPath = path.join(jobDir, file.filename);
          const outputDir = path.join(jobDir, "split-out");
          fs.mkdirSync(outputDir, { recursive: true });

          if (scriptConfig.requiresOutputDir) {
            await runPythonScript(scriptPath, [inputPath, "-o", outputDir]);
          } else {
            await runPythonScript(scriptPath, [inputPath], { cwd: outputDir });
          }

          const grabCtx: OutputGrabContext = { jobDir, outputDir };
          const outFiles = scriptConfig.customOutputGrab(grabCtx);
          if (outFiles.length === 0) {
            destroyJob();
            if (!res.headersSent) {
              res.status(500).json({ error: "處理後未產生任何檔案" });
            }
            return;
          }

          let cleanedAfterStream = false;
          const cleanupOnceAfterStream = () => {
            if (cleanedAfterStream) return;
            cleanedAfterStream = true;
            destroyJob();
          };
          res.once("finish", cleanupOnceAfterStream);
          res.once("close", cleanupOnceAfterStream);

          const zipName = `ticket-split-${airlineKey}-${Date.now()}.zip`;
          res.setHeader("Content-Type", "application/zip");
          res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);

          const archive = archiver("zip", { zlib: { level: 9 } });
          archive.on("error", archiveErr => {
            console.error("[upload-and-split] archiver:", archiveErr);
            cleanupOnceAfterStream();
            if (!res.headersSent) {
              res.status(500).json({ error: "打包失敗" });
            } else {
              res.destroy();
            }
          });

          archive.pipe(res);

          addFilesToArchive(archive, outFiles, outputDir);

          await archive.finalize();
        } catch (e) {
          destroyJob();
          console.error("[upload-and-split]", e);
          if (!res.headersSent) {
            res.status(500).json({
              error: e instanceof Error ? e.message : "處理失敗",
            });
          }
        }
      })();
    });
  });
}
