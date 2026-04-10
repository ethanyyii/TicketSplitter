# Python 腳本集成指南

## 概述

TicketSplitter 後端已準備好集成您的 Python 檔案分割腳本。本指南說明如何整合三個航空公司的分割函數。

## 當前架構

### 後端結構

```
server/
├── routers/
│   └── split.ts              # 檔案分割路由（已實現）
├── scripts/                  # Python 腳本目錄（待創建）
│   ├── lion_air_splitter.py  # 泰國獅子航空分割腳本
│   ├── tiger_air_splitter.py # 台灣虎航分割腳本
│   └── scoot_air_splitter.py # 酷航分割腳本
└── routers.ts               # 主路由（已集成 splitRouter）
```

### 流程

```
前端上傳檔案
    ↓
後端接收 (split.split mutation)
    ↓
保存到臨時位置 (/tmp/ticketsplitter/)
    ↓
調用對應 Python 腳本
    ↓
Python 腳本返回分割後的檔案列表
    ↓
打包成 ZIP 檔案
    ↓
返回 ZIP 給前端下載
    ↓
自動清理所有臨時檔案
```

## 集成步驟

### 1. 準備 Python 腳本

每個 Python 腳本應該導出一個主要函數，簽名如下：

```python
def split_lion_air(input_file_path: str, output_dir: str) -> List[str]:
    """
    分割泰國獅子航空的機票檔案
    
    Args:
        input_file_path: 輸入檔案的完整路徑
        output_dir: 輸出目錄的完整路徑
    
    Returns:
        分割後檔案的完整路徑列表
        例如: ['/tmp/ticketsplitter/xxx/page1.pdf', '/tmp/ticketsplitter/xxx/page2.pdf']
    """
    # 您的分割邏輯
    output_files = []
    # ... 處理檔案 ...
    return output_files
```

類似地實現：
- `split_tiger_air(input_file_path: str, output_dir: str) -> List[str]`
- `split_scoot_air(input_file_path: str, output_dir: str) -> List[str]`

### 2. 修改 split.ts 中的 getSplitFunction()

在 `server/routers/split.ts` 中找到 `getSplitFunction()` 函數，修改為調用您的 Python 腳本：

```typescript
import { spawn } from 'child_process';

async function getSplitFunction(airline: string) {
  const splitFunctions: Record<string, Function> = {
    SL: async (input: string, output: string) => {
      return await callPythonScript('lion_air_splitter', 'split_lion_air', input, output);
    },
    IT: async (input: string, output: string) => {
      return await callPythonScript('tiger_air_splitter', 'split_tiger_air', input, output);
    },
    BZ: async (input: string, output: string) => {
      return await callPythonScript('scoot_air_splitter', 'split_scoot_air', input, output);
    },
  };

  return splitFunctions[airline] || splitFunctions.SL;
}

async function callPythonScript(
  moduleName: string,
  functionName: string,
  inputPath: string,
  outputDir: string
): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const pythonProcess = spawn('python3', [
      '-c',
      `
import sys
sys.path.insert(0, '${path.join(__dirname, '../scripts')}')
from ${moduleName} import ${functionName}
import json

try:
    result = ${functionName}('${inputPath}', '${outputDir}')
    print(json.dumps(result))
except Exception as e:
    print(json.dumps({'error': str(e)}), file=sys.stderr)
    sys.exit(1)
      `,
    ]);

    let stdout = '';
    let stderr = '';

    pythonProcess.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    pythonProcess.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    pythonProcess.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`Python script failed: ${stderr}`));
      } else {
        try {
          const result = JSON.parse(stdout);
          if (result.error) {
            reject(new Error(result.error));
          } else {
            resolve(result);
          }
        } catch (error) {
          reject(new Error(`Failed to parse Python output: ${stdout}`));
        }
      }
    });
  });
}
```

### 3. 安裝 Python 依賴

確保您的 Python 環境中已安裝必要的依賴：

```bash
pip install PyMuPDF python-docx
```

### 4. 測試集成

使用提供的測試檔案驗證集成：

```bash
pnpm test
```

## 隱私和安全考量

### 檔案清理保證

- ✓ 所有臨時檔案在 `finally` 塊中被清理
- ✓ 即使發生錯誤也會清理
- ✓ 使用 `fs.rmSync({ recursive: true, force: true })` 確保完全刪除
- ✓ 沒有任何檔案被永久存儲

### 檔案驗證

前端已實現以下驗證：
- 檔案類型檢查（PDF、Word、JPG、PNG）
- 檔案大小限制（最大 50MB）
- 後端應額外驗證檔案內容

## 錯誤處理

後端已實現完整的錯誤處理：

```typescript
try {
  // 上傳和處理
} catch (error) {
  // 錯誤自動被捕獲並返回給前端
  // 前端會顯示友好的錯誤訊息
} finally {
  // 無論成功或失敗，都會清理臨時檔案
  await tempManager.cleanup();
}
```

## 前端 API 調用

前端通過 tRPC 調用後端：

```typescript
const splitMutation = trpc.split.split.useMutation({
  onSuccess: (data) => {
    // data.zipBuffer: base64 編碼的 ZIP 檔案
    // data.filename: 建議的下載檔名
    // data.size: ZIP 檔案大小（字節）
  },
  onError: (error) => {
    // 錯誤訊息已自動顯示給用戶
  },
});

// 調用
await splitMutation.mutateAsync({
  airline: 'SL', // 'SL' | 'IT' | 'BZ'
  filename: 'ticket.pdf',
  fileBuffer: Buffer.from(fileContent),
});
```

## 性能優化建議

1. **非同步處理**：如果分割耗時較長，考慮實現隊列系統
2. **進度報告**：可以通過 WebSocket 實時報告進度
3. **快取**：如果相同檔案被多次上傳，考慮臨時快取
4. **並發限制**：限制同時處理的檔案數量

## 故障排除

### Python 腳本找不到

確保：
1. Python 腳本在 `server/scripts/` 目錄中
2. 檔名和函數名正確
3. Python 路徑正確配置

### 檔案未被清理

檢查：
1. 臨時目錄權限
2. 檔案是否被其他進程鎖定
3. 磁盤空間是否充足

### 前端無法下載

確保：
1. ZIP 檔案正確生成
2. Base64 編碼正確
3. 瀏覽器允許下載

## 聯繫和支援

如有任何問題，請檢查：
1. 伺服器日誌：`server/_core/index.ts` 輸出
2. 測試結果：`pnpm test`
3. 前端控制台：瀏覽器開發者工具

---

**注意**：此指南假設您已有完整的 Python 分割腳本。如需幫助編寫腳本，請提供機票檔案樣本。
