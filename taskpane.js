/* global Office, Excel */

const CONFIG_KEY = "siliconflow_excel_cleaner_config";

function log(message) {
  const el = document.getElementById("log");
  console.log("[AI 数据清洗助手]", message);
  if (!el) return;
  const time = new Date().toLocaleTimeString();
  el.textContent = `[${time}] ${message}\n` + el.textContent;
}

function loadConfig() {
  try {
    const raw = window.localStorage.getItem(CONFIG_KEY);
    if (!raw) return {};
    return JSON.parse(raw);
  } catch (e) {
    log("读取本地配置失败：" + e.message);
    return {};
  }
}

function saveConfigToStorage(config) {
  try {
    window.localStorage.setItem(CONFIG_KEY, JSON.stringify(config));
    log("配置已保存到本机。");
  } catch (e) {
    log("保存配置失败：" + e.message);
  }
}

function initUiFromConfig() {
  const cfg = loadConfig();
  if (cfg.apiKey) document.getElementById("apiKey").value = cfg.apiKey;
  if (cfg.apiBaseUrl) document.getElementById("apiBaseUrl").value = cfg.apiBaseUrl;
  if (cfg.modelName) document.getElementById("modelName").value = cfg.modelName;
}

async function callSiliconflowApi(values, instruction) {
  const apiKey = document.getElementById("apiKey").value.trim();
  const apiBaseUrl = document.getElementById("apiBaseUrl").value.trim();
  const modelName = document.getElementById("modelName").value.trim();

  if (!apiKey) {
    throw new Error("请先填写 API Key。");
  }
  if (!apiBaseUrl) {
    throw new Error("请先填写 API 地址。");
  }
  if (!modelName) {
    throw new Error("请先填写模型名称。");
  }

  const systemPrompt =
    "你是一个专业的数据清洗助手。用户会给你一个二维数组（Excel 单元格值），" +
    "请根据用户的清洗指令，对数据进行清洗和标准化，例如：去掉首尾空格、统一日期格式、" +
    "规范大小写、删除完全空行等。只返回清洗后的二维数组 JSON，不要包含额外文字或代码块标记。";

  const userPrompt =
    "原始数据（JSON 二维数组）如下：\n" +
    JSON.stringify(values) +
    "\n\n清洗指令：" +
    (instruction || "默认：去除首尾空格，删除完全空行，保留原有结构。") +
    "\n\n请只返回清洗后的二维数组 JSON，不要包含 ```json 等代码块标记，直接返回例如：[[\"a\",\"b\"],[\"c\",\"d\"]]。";

  const body = {
    model: modelName,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    temperature: 0,
  };

  log("正在调用 SiliconFlow 接口...");
  log("API 地址: " + apiBaseUrl);
  log("模型: " + modelName);

  const resp = await fetch(apiBaseUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text();
    log("API 错误响应: " + text);
    throw new Error("SiliconFlow 请求失败：" + resp.status + " - " + text);
  }

  const data = await resp.json();

  // OpenAI Chat Completion 兼容结构
  const content =
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;

  if (!content) {
    log("API 返回数据: " + JSON.stringify(data));
    throw new Error("未从接口结果中解析到内容。");
  }

  log("接口返回内容：" + content.slice(0, 300) + "...");

  // 清理可能存在的代码块标记
  let cleanContent = content.trim();
  if (cleanContent.startsWith("```json")) {
    cleanContent = cleanContent.slice(7);
  } else if (cleanContent.startsWith("```")) {
    cleanContent = cleanContent.slice(3);
  }
  if (cleanContent.endsWith("```")) {
    cleanContent = cleanContent.slice(0, -3);
  }
  cleanContent = cleanContent.trim();

  // content 预期是 JSON 字符串形式的二维数组
  let cleaned;
  try {
    cleaned = JSON.parse(cleanContent);
  } catch (e) {
    log("尝试解析的内容: " + cleanContent.slice(0, 500));
    throw new Error("解析返回 JSON 失败：" + e.message);
  }

  if (!Array.isArray(cleaned) || (cleaned.length > 0 && !Array.isArray(cleaned[0]))) {
    throw new Error("返回结果不是二维数组，请检查模型输出。");
  }

  return cleaned;
}

async function onCleanDataClick() {
  const btn = document.getElementById("cleanDataBtn");
  btn.disabled = true;
  btn.textContent = "处理中...";

  try {
    const instruction = document
      .getElementById("cleanInstruction")
      .value.trim();

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "address", "rowCount", "columnCount"]);
      await context.sync();

      log(
        `已选择区域：${range.address}，行数：${range.rowCount}，列数：${range.columnCount}`
      );

      const originalValues = range.values;
      log("原始数据: " + JSON.stringify(originalValues).slice(0, 200) + "...");

      const cleanedValues = await callSiliconflowApi(
        originalValues,
        instruction
      );

      if (
        cleanedValues.length !== originalValues.length ||
        cleanedValues[0].length !== originalValues[0].length
      ) {
        log(
          "注意：清洗后的二维数组尺寸与原始选择区域不一致，将自动截断或填充。"
        );
      }

      // 调整到与当前选择区域相同的尺寸
      const rows = range.rowCount;
      const cols = range.columnCount;
      const finalValues = [];

      for (let r = 0; r < rows; r++) {
        const row = [];
        for (let c = 0; c < cols; c++) {
          row.push(
            cleanedValues[r] && cleanedValues[r][c] !== undefined
              ? cleanedValues[r][c]
              : ""
          );
        }
        finalValues.push(row);
      }

      range.values = finalValues;
      await context.sync();

      log("数据清洗完成并已写回选中区域。");
    });
  } catch (e) {
    log("清洗失败：" + e.message);
    console.error(e);
  } finally {
    btn.disabled = false;
    btn.textContent = "清洗选中区域数据";
  }
}

function onSaveConfigClick() {
  const config = {
    apiKey: document.getElementById("apiKey").value.trim(),
    apiBaseUrl: document.getElementById("apiBaseUrl").value.trim(),
    modelName: document.getElementById("modelName").value.trim(),
  };
  saveConfigToStorage(config);
}

function wireUpUi() {
  initUiFromConfig();
  const saveBtn = document.getElementById("saveConfigBtn");
  const cleanBtn = document.getElementById("cleanDataBtn");
  if (saveBtn) {
    saveBtn.addEventListener("click", onSaveConfigClick);
  }
  if (cleanBtn) {
    cleanBtn.addEventListener("click", onCleanDataClick);
  }
  log("界面已初始化。请在 Excel 中选择数据区域后点击清洗按钮。");
}

// Office Add-in 初始化
Office.onReady((info) => {
  log("Office.onReady 触发");
  log("Host: " + info.host);
  wireUpUi();
});