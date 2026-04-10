import { useState, useRef } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { toast } from "sonner";
import { Spinner } from "@/components/ui/spinner";

interface UploadCardProps {
  airline: string;
  /** POST `airline`，須與後端 ScriptRegistry key 一致 */
  airlineKey: string;
  /** 卡片左上角小標，例如 IATA 或品牌代碼 */
  badgeCode: string;
  description: string;
}

type UploadState = "idle" | "dragover" | "uploading" | "success" | "error";

function parseContentDispositionFilename(header: string | null): string | null {
  if (!header) return null;
  const star = /filename\*\s*=\s*UTF-8''([^;\s]+)/i.exec(header);
  if (star) {
    try {
      return decodeURIComponent(star[1].replace(/"/g, "").trim());
    } catch {
      return star[1].trim();
    }
  }
  const plain = /filename\s*=\s*("?)([^";\n]+)\1/i.exec(header);
  return plain ? plain[2].trim() : null;
}

function triggerBlobDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export default function UploadCard({ airline, airlineKey, badgeCode, description }: UploadCardProps) {
  const [state, setState] = useState<UploadState>("idle");
  const [progress, setProgress] = useState(0);
  const [errorMessage, setErrorMessage] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dragCounterRef = useRef(0);
  const progressIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const clearProgressInterval = () => {
    if (progressIntervalRef.current) {
      clearInterval(progressIntervalRef.current);
      progressIntervalRef.current = null;
    }
  };

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current++;
    if (state !== "uploading") {
      setState("dragover");
    }
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current--;
    if (dragCounterRef.current === 0 && state !== "uploading") {
      setState("idle");
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current = 0;
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      void handleFiles(files);
    }
  };

  const handleFiles = async (files: FileList) => {
    const file = files[0];
    if (!file) return;

    const validTypes = [
      "application/pdf",
      "application/msword",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "image/jpeg",
      "image/png",
    ];
    if (!validTypes.includes(file.type)) {
      setErrorMessage("不支援的檔案格式。請上傳 PDF、Word 或圖片檔案。");
      setState("error");
      toast.error("不支援的檔案格式");
      setTimeout(() => {
        setState("idle");
        setErrorMessage("");
      }, 3000);
      return;
    }

    const maxSize = 50 * 1024 * 1024;
    if (file.size > maxSize) {
      setErrorMessage("檔案過大。請上傳小於 50MB 的檔案。");
      setState("error");
      toast.error("檔案過大");
      setTimeout(() => {
        setState("idle");
        setErrorMessage("");
      }, 3000);
      return;
    }

    setState("uploading");
    setProgress(0);
    clearProgressInterval();
    progressIntervalRef.current = setInterval(() => {
      setProgress(prev => {
        if (prev >= 88) return 88;
        return prev + Math.random() * 12 + 4;
      });
    }, 280);

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("airline", airlineKey);

      const res = await fetch("/api/upload-and-split", {
        method: "POST",
        body: formData,
        credentials: "include",
      });

      clearProgressInterval();

      if (!res.ok) {
        let msg = `請求失敗（${res.status}）`;
        const text = await res.text();
        try {
          const data = JSON.parse(text) as { error?: string };
          if (data?.error) msg = data.error;
        } catch {
          if (text) msg = text.slice(0, 200);
        }
        throw new Error(msg);
      }

      const blob = await res.blob();
      const fromHeader = parseContentDispositionFilename(res.headers.get("Content-Disposition"));
      const filename = fromHeader || `ticket-split-${airlineKey}-${Date.now()}.zip`;

      triggerBlobDownload(blob, filename);
      setProgress(100);
      setState("success");
      toast.success(`已完成拆分並下載 ${filename}`);

      setTimeout(() => {
        setState("idle");
        setProgress(0);
      }, 2800);
    } catch (error) {
      clearProgressInterval();
      const errorMsg = error instanceof Error ? error.message : "檔案分割失敗，請重試";
      setErrorMessage(errorMsg);
      setState("error");
      toast.error(errorMsg);
      setProgress(0);

      setTimeout(() => {
        setState("idle");
        setErrorMessage("");
      }, 3200);
    }
  };

  const handleCardClick = () => {
    if (state !== "uploading") {
      fileInputRef.current?.click();
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.currentTarget.files;
    if (files) {
      void handleFiles(files);
    }
    e.currentTarget.value = "";
  };

  const isBusy = state === "uploading";
  const showHoverMotion = state === "idle";

  return (
    <motion.div
      className="relative"
      animate={
        state === "dragover"
          ? {
              scale: 1.045,
              y: -10,
              transition: { type: "spring", stiffness: 420, damping: 22 },
            }
          : { scale: 1, y: 0, transition: { type: "spring", stiffness: 380, damping: 28 } }
      }
      whileHover={
        showHoverMotion
          ? {
              y: -12,
              scale: 1.02,
              transition: { type: "spring", stiffness: 400, damping: 26 },
            }
          : undefined
      }
      style={{ transformOrigin: "center center" }}
    >
      <motion.div
        role="button"
        tabIndex={0}
        onKeyDown={e => {
          if ((e.key === "Enter" || e.key === " ") && !isBusy) {
            e.preventDefault();
            handleCardClick();
          }
        }}
        className={`card upload-card-motion ${state === "dragover" ? "drag-over" : ""}`}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        onClick={handleCardClick}
        layout
        transition={{ layout: { duration: 0.35, ease: [0.4, 0, 0.2, 1] } }}
        style={{
          borderColor: state === "dragover" ? "#0066ff" : undefined,
        }}
      >
        <input
          ref={fileInputRef}
          type="file"
          className="hidden"
          onChange={handleFileInputChange}
          accept=".pdf,.doc,.docx,.jpg,.jpeg,.png"
          disabled={isBusy}
          aria-label={`上傳 ${airline} 機票檔案`}
        />

        <AnimatePresence mode="wait">
          {state === "idle" && (
            <motion.div
              key="idle"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.3 }}
              className="w-full"
            >
              <div className="card-code">{badgeCode}</div>
              <div className="card-title">{airline}</div>
              <div className="card-description">{description}</div>
              <div className="upload-zone">
                <div className="upload-icon">↑</div>
                <div className="upload-text">點擊或拖曳檔案</div>
                <div className="upload-hint">支援 PDF、Word、JPG、PNG</div>
              </div>
            </motion.div>
          )}

          {state === "dragover" && (
            <motion.div
              key="dragover"
              initial={{ opacity: 0, scale: 0.96 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.96 }}
              transition={{ duration: 0.22, ease: [0.4, 0, 0.2, 1] }}
              className="w-full text-center"
            >
              <div className="text-lg font-semibold text-blue-600 tracking-tight">放開以上傳</div>
              <p className="text-xs text-slate-500 mt-2">將為您安全處理並打包下載</p>
            </motion.div>
          )}

          {state === "uploading" && (
            <motion.div
              key="uploading"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.3 }}
              className="w-full flex flex-col items-center"
            >
              <div className="relative mb-5">
                <div className="absolute inset-0 rounded-full bg-blue-500/15 animate-ping" style={{ animationDuration: "1.8s" }} />
                <Spinner className="size-11 text-[var(--accent-primary)] relative z-10" />
              </div>
              <div className="text-sm font-medium text-slate-800 mb-3 tracking-tight">正在上傳並拆分…</div>
              <div className="w-full max-w-[220px] h-1.5 rounded-full bg-slate-200/80 overflow-hidden">
                <motion.div
                  className="h-full rounded-full bg-gradient-to-r from-[var(--gradient-start)] to-[var(--gradient-end)]"
                  initial={{ width: "0%" }}
                  animate={{ width: `${Math.min(progress, 100)}%` }}
                  transition={{ duration: 0.35, ease: "easeOut" }}
                />
              </div>
              <div className="text-xs text-slate-500 mt-2 tabular-nums">{Math.round(Math.min(progress, 100))}%</div>
            </motion.div>
          )}

          {state === "success" && (
            <motion.div
              key="success"
              initial={{ opacity: 0, scale: 0.92 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.92 }}
              transition={{ duration: 0.35, ease: [0.4, 0, 0.2, 1] }}
              className="w-full text-center"
            >
              <motion.div
                className="mx-auto mb-3 flex size-14 items-center justify-center rounded-full bg-emerald-500/10 text-emerald-600 text-2xl font-semibold"
                animate={{ scale: [1, 1.06, 1] }}
                transition={{ duration: 0.55, ease: "easeOut" }}
              >
                ✓
              </motion.div>
              <div className="status-text status-success">拆分完成</div>
              <div className="text-xs text-slate-500 mt-2">ZIP 已開始下載</div>
            </motion.div>
          )}

          {state === "error" && (
            <motion.div
              key="error"
              initial={{ opacity: 0, scale: 0.92 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.92 }}
              transition={{ duration: 0.3 }}
              className="w-full text-center"
            >
              <div className="text-4xl mb-3">⚠</div>
              <div className="status-text status-error">{errorMessage || "上傳失敗"}</div>
              <div className="text-xs text-slate-500 mt-2">請檢查檔案格式後重試</div>
            </motion.div>
          )}
        </AnimatePresence>
      </motion.div>
    </motion.div>
  );
}
