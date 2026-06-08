// api/quote.js
// 企业微信「内部智能机器人 · API 模式」回调（对齐官方 Python3 JSON Demo）
// 功能：
// 1）GET：URL 校验（解密 echostr 返回明文）
// 2）POST：解密用户消息 → 调用业务逻辑 → 按 msgtype=stream 格式加密返回
//
// 全程只用 Token + EncodingAESKey，不需要 corpsecret / send_url / access_token。

const crypto = require("crypto");
const path = require("path");
const XLSX = require("xlsx");

// Excel 放在和 quote.js 同一目录（api/）下
const XLS_FILE = path.join(__dirname, process.env.PRICE_WORKBOOK || "VMC系列价格表.xlsx");



// ===== 1. 机器人回调配置（用你的实际配置替换） =====
const TOKEN = process.env.WECOM_TOKEN || "";
const EncodingAESKey = process.env.WECOM_ENCODING_AES_KEY || "";
// 智能机器人场景 receiveid 为空字符串（官方文档说明）
const RECEIVE_ID = process.env.WECOM_RECEIVE_ID || "";

// ===== 2. 签名计算 / 校验 =====
function calcSignature(token, timestamp, nonce, encrypt) {
  const arr = [token, timestamp, nonce, encrypt].sort();
  return crypto.createHash("sha1").update(arr.join("")).digest("hex");
}

function verifySignature(token, timestamp, nonce, encrypt, msgSignature) {
  const sig = calcSignature(token, timestamp, nonce, encrypt);
  return sig === msgSignature;
}

// ===== 3. PKCS#7 补位 / 去补位 =====
function pkcs7Unpad(buf) {
  const pad = buf[buf.length - 1];
  if (pad < 1 || pad > 32) {
    throw new Error("invalid padding");
  }
  return buf.slice(0, buf.length - pad);
}

function pkcs7Pad(buf) {
  const blockSize = 32;
  const pad = blockSize - (buf.length % blockSize || blockSize);
  const padBuf = Buffer.alloc(pad, pad);
  return Buffer.concat([buf, padBuf]);
}

// ===== 4. AES key / 解密 =====
function aesKeyBuf() {
  if (!EncodingAESKey) {
    throw new Error("Missing WECOM_ENCODING_AES_KEY");
  }
  // EncodingAESKey 43 位，要补一个 "=" 再按 base64 解
  return Buffer.from(EncodingAESKey + "=", "base64");
}

function decryptWeCom(encrypt) {
  const key = aesKeyBuf();
  const iv = key.slice(0, 16);

  const cipherText = Buffer.from(encrypt, "base64");
  const decipher = crypto.createDecipheriv("aes-256-cbc", key, iv);
  decipher.setAutoPadding(false);

  let decrypted = Buffer.concat([decipher.update(cipherText), decipher.final()]);
  decrypted = pkcs7Unpad(decrypted);

  // 明文结构：16字节随机串 + 4字节msg_len + msg + receiveId
  const msgLen = decrypted.slice(16, 20).readUInt32BE(0);
  const msgBuf = decrypted.slice(20, 20 + msgLen);
  const msg = msgBuf.toString("utf8");
  const rest = decrypted.slice(20 + msgLen).toString("utf8"); // receiveId（这里为空）

  return { msg, receiveId: rest };
}

// ===== 5. 加密明文 JSON，生成 encrypt + msgsignature + timestamp + nonce =====
function encryptWeCom(plainJsonStr, nonceFromReq) {
  const key = aesKeyBuf();
  const iv = key.slice(0, 16);

  const random16 = crypto.randomBytes(16);
  const msgBuf = Buffer.from(plainJsonStr, "utf8");
  const msgLenBuf = Buffer.alloc(4);
  msgLenBuf.writeUInt32BE(msgBuf.length, 0);

  const plainBuf = Buffer.concat([
    random16,
    msgLenBuf,
    msgBuf,
    Buffer.from(RECEIVE_ID, "utf8"),
  ]);

  const padded = pkcs7Pad(plainBuf);
  const cipher = crypto.createCipheriv("aes-256-cbc", key, iv);
  cipher.setAutoPadding(false);
  const encryptedBuf = Buffer.concat([cipher.update(padded), cipher.final()]);
  const encrypt = encryptedBuf.toString("base64");

  const timestamp = Math.floor(Date.now() / 1000).toString();
  const nonce = nonceFromReq || crypto.randomBytes(8).toString("hex");
  const msgsignature = calcSignature(TOKEN, timestamp, nonce, encrypt);

  return {
    encrypt,
    msgsignature,
    timestamp,
    nonce,
  };
}

// ===== 6. 查价引擎逻辑（JS 版，来自 quote.py + server.js） =====


// 与 Python 版一致的路由/规则配置
const DEFAULT_SHEETS = [
  "VMC"
];

const SUFFIX_ENDINGS = [
  "03LFK", "03LF", "03SFLFK", "03SFLF",
  "04HTECKLF", "04HTECLF", "07LF", "07LFK"
];

const M_HEAD_WHITELIST = [
  "M010", "M015", "M020", "M025", "M032", "M040"
];

// 读取工作簿（缓存，避免每条消息都重新读一次文件）
let _workbookCache = null;
function getWorkbook() {
  if (_workbookCache) return _workbookCache;
  _workbookCache = XLSX.readFile(XLS_FILE, { cellDates: false });
  return _workbookCache;
}

// normalize：和 Python 版保持一致的清洗逻辑
function normalize(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  const rep = { "，": ",", "。": ".", "．": ".", "・": ".", "　": " " };
  for (const [k, v] of Object.entries(rep)) {
    s = s.split(k).join(v);
  }
  return s.trim();
}

// 型号解析：parse_model
function parseModel(model) {
  // 先做 normalize（去全角、空格），再统一转大写
  const s = normalize(model).toUpperCase();
  const parts = s.split(".").filter(Boolean);
  if (parts.length < 2) return { colKey: null, rowKey: null };

  const head = parts[0];

  if (head.startsWith("VMCE")) {
    return {
      sheetName: "VMCE",
      colKey: parts.slice(1).join("."),
      rowKey: head
    };
  }

  if (
    head.startsWith("VMC") ||
    head.startsWith("VMP")
  ) {
    if (parts.length < 4) return { colKey: null, rowKey: null };
    return {
      sheetName: parts.slice(2).join("."),
      colKey: parts[1],
      rowKey: head
    };
  }

  if (head.startsWith("MUC")) {
    const match = head.match(/^(MUC\d+)(.+)$/);
    if (!match) return { colKey: null, rowKey: null };
    return {
      sheetName: "MUC",
      colKey: [match[2], ...parts.slice(1)].join("."),
      rowKey: match[1]
    };
  }

  if (head.startsWith("GMC")) {
    return {
      sheetName: "GMC",
      colKey: parts.slice(1).join("."),
      rowKey: head
    };
  }

  if (head.startsWith("M")) {
    return {
      sheetName: "M",
      colKey: head,
      rowKey: parts.slice(1).join(".")
    };
  }

  return { colKey: null, rowKey: null };
}


// 路由 sheet：route_sheets
function routeSheets(model) {
  const parsed = parseModel(model);
  if (parsed.sheetName) return [parsed.sheetName];
  return DEFAULT_SHEETS;
}

// 小工具：统一做 normalize + 大写
function normalizeModel(model) {
  return normalize(model).toUpperCase();
}

function routeSheetsFixed(model, parsed) {
  const m = normalizeModel(model);  // 等价于你现在那一行

  if (parsed && parsed.sheetName) return [parsed.sheetName];
  if (m.startsWith("VMC")) return DEFAULT_SHEETS;
  return DEFAULT_SHEETS;
}


// 载入某个 sheet，返回 [sheetData, headers, rowKeys]
// sheetData: 2D 数组，sheetData[rowIndex][colIndex]
// headers: 第一行列头（已 normalize）
// rowKeys: 行标（来自第一列）
function loadSheet(sheetName) {
  const wb = getWorkbook();
  const sheet = wb.Sheets[sheetName];
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (!aoa || aoa.length === 0) {
    throw new Error("Empty sheet: " + sheetName);
  }

  const rawHeaders = aoa[0] || [];
  const headers = rawHeaders.map(normalize);

  // 行标：第一列，从第 2 行开始
  function normalizeRow(v) {
    let x = normalize(v);
    if (x.endsWith(".0")) x = x.slice(0, -2);
    return x;
  }

  const rowKeys = [];
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    rowKeys.push(normalizeRow(row[0]));
  }

  return { sheetData: aoa, headers, rowKeys };
}

// find_exact：精确匹配 + 兼容 '71' vs '71.0'
function findExact(list, key) {
  // 统一大写，避免大小写差异导致找不到
  let k = normalize(key).toUpperCase();
  if (k.endsWith(".0")) k = k.slice(0, -2);

  for (let i = 0; i < list.length; i++) {
    let v = normalize(list[i]).toUpperCase();
    if (v.endsWith(".0")) v = v.slice(0, -2);
    if (v === k) return i;
  }
  return -1;
}


// apply_rule：VMC 系列报价调整规则
function applyRule(model, base) {
  const s = normalizeModel(model);
  const parts = s.split(".").filter(Boolean);

  if (s.startsWith("VMCX")) {
    const adjusted = round2(base * 1.4);
    return {
      rule: "VMCX_MULTIPLY_1.4",
      adjusted,
      formula: `${base.toFixed(2)} * 1.4 = ${adjusted.toFixed(2)}`
    };
  }

  if (s.startsWith("VMPX") && parts[2] && parts[2].includes("50")) {
    const adjusted = round2(base * 1.4);
    return {
      rule: "VMPX_50_MULTIPLY_1.4",
      adjusted,
      formula: `${base.toFixed(2)} * 1.4 = ${adjusted.toFixed(2)}`
    };
  }

  return {
    rule: "NONE",
    adjusted: round2(base),
    formula: `${base.toFixed(2)} = ${round2(base).toFixed(2)}`
  };
}

function round2(x) {
  return Math.round((x + Number.EPSILON) * 100) / 100;
}

function buildFail(status, reason, extra) {
  return Object.assign({ status, reason }, extra || {});
}

function extractModelText(text) {
  return normalize(text)
    .replace(/^(查价|查询|报价|价格)\s*[:：]?\s*/i, "")
    .trim();
}

// quote_model：完全复刻 Python quote.py 的主逻辑
function quoteModel(model) {
  const m = model || "";
  if (!m.trim()) {
    return buildFail("FAIL_B", "NO_MODEL_INPUT");
  }

  const parsed = parseModel(m);
  const { sheetName, colKey, rowKey } = parsed;
  if (!colKey || !rowKey) {
    return buildFail("FAIL_B", "PARSE_ERROR", { model: m });
  }

  const sheets = routeSheetsFixed(m, parsed);

  for (const sheetName of sheets) {
    let sheetObj;
    try {
      sheetObj = loadSheet(sheetName);
    } catch (e) {
      console.error("loadSheet error:", sheetName, e);
      continue;
    }

    const { sheetData, headers, rowKeys } = sheetObj;
    const colIdx = findExact(headers, colKey);
    const rowIdx = findExact(rowKeys, rowKey);

    if (colIdx < 0 || rowIdx < 0) {
      // 在当前 sheet 找不到，换下一个 sheet
      continue;
    }

    // sheetData 第 0 行是表头，所以数据行从 1 开始
    const rawRow = sheetData[rowIdx + 1] || [];
    const cellVal = rawRow[colIdx];
    const sVal = normalize(cellVal);

    if (!sVal || sVal.toLowerCase() === "nan") {
      return buildFail("FAIL_A", "EMPTY_CELL", {
        sheet: sheetName,
        column_key: colKey,
        row_key: rowKey
      });
    }

    const base = parseFloat(sVal);
    if (Number.isNaN(base)) {
      return buildFail("FAIL_A", "NON_NUMERIC_CELL", {
        sheet: sheetName,
        column_key: colKey,
        row_key: rowKey,
        raw: sVal
      });
    }

    const baseRounded = round2(base);
    const { rule, adjusted, formula } = applyRule(m, baseRounded);

    const rate = 12.5;
    const cny = round2(adjusted * rate);

    return {
      status: "OK",
      sheet: sheetName,
      model: m,
      column_key: colKey,
      row_key: rowKey,
      base_price_eur: baseRounded,
      adjusted_price_eur: adjusted,
      rule_applied: rule,
      rule_formula: formula,
      sales_multiplier: rate,
      sales_price_cny: cny
    };
  }

  return buildFail("FAIL_B", "NOT_FOUND", {
    model: m,
    sheet: sheetName,
    column_key: colKey,
    row_key: rowKey
  });
}

// 把 JSON 结果转成回复文本（等价于 server.js 里的 formatReply）
function formatQuoteReply(data) {
  if (!data || data.status !== "OK") {
    const r = data || {};
    return [
      "未找到对应价格或不允许报价。",
      r.reason ? `原因: ${r.reason}` : "",
      r.model ? `型号: ${r.model}` : ""
    ].filter(Boolean).join("\n");
  }

  return [
    `表：${data.sheet}`,
    `定位：${data.column_key} × ${data.row_key}`,
    `原值(EUR)：${data.base_price_eur.toFixed(2)}`,
    `规则：${data.rule_applied}`,
    `计算公式：${data.rule_formula}`,
    `调整后(EUR)：${data.adjusted_price_eur.toFixed(2)}`,
    `销售价格系数：${data.sales_multiplier}`,
    `销售价格(CNY)：${data.sales_price_cny.toFixed(2)}`
  ].join("\n");
}

// ===== 7. 业务逻辑入口：在这里塞“生意逻辑和脑子” =====
// eventObj: 企微解密后的完整 JSON
// userText: 用户发来的文本内容（string）
async function runBusinessLogic(eventObj, userText) {
  // 1）空消息兜底
  if (!userText || !userText.trim()) {
    return "请发送要查询的 VMC 型号或问题。";
  }

  const text = userText.trim();

  // 2）帮助指令
  if (text === "帮助" || text.toLowerCase() === "help") {
    return [
      "👋 我是 VMC 报价助手。",
      "",
      "用法示例：",
      "1）直接发型号：",
      "   VMC25.03XK.50N.50",
      "",
      "2）带前缀说明也行：",
      "   查价 VMC25.03XK.50N.50",
      "",
      "我会在《VMC系列价格表.xlsx》中严格定位列头/行标，并按 VMC 规则给出报价。",
    ].join("\n");
  }

  // 3）尝试从文本中提取型号
  const model = extractModelText(text);

  const quoteResult = quoteModel(model);
  const replyText = formatQuoteReply(quoteResult);

  // 如果完全没匹配到（NOT_FOUND / PARSE_ERROR），再加一句提示
  if (quoteResult.status !== "OK") {
    return replyText + "\n\n（提示：请检查型号格式是否与 Excel 表头/行标一致）";
  }

  return replyText;
}

async function readRequestBody(req) {
  if (typeof req.body === "string") return req.body;
  if (Buffer.isBuffer(req.body)) return req.body.toString("utf8");
  if (req.body && typeof req.body === "object") return JSON.stringify(req.body);

  return new Promise((resolve, reject) => {
    let bodyStr = "";
    req.on("data", (chunk) => {
      bodyStr += Buffer.isBuffer(chunk) ? chunk.toString("utf8") : chunk;
    });
    req.on("end", () => resolve(bodyStr));
    req.on("error", reject);
  });
}

async function sendResponseUrl(responseUrl, content) {
  if (!responseUrl) return false;
  if (typeof fetch !== "function") {
    console.error("response_url send skipped: fetch is not available");
    return false;
  }

  try {
    const resp = await fetch(responseUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({
        msgtype: "markdown",
        markdown: {
          content,
        },
      }),
    });

    const text = await resp.text();
    console.log("response_url result:", {
      status: resp.status,
      ok: resp.ok,
      body: text,
    });

    if (!resp.ok) return false;
    try {
      const json = JSON.parse(text);
      return json.errcode === undefined || json.errcode === 0;
    } catch (e) {
      return true;
    }
  } catch (e) {
    console.error("response_url send error:", e);
    return false;
  }
}


// ===== 7. Vercel Handler =====
module.exports = async function handler(req, res) {
  try {
    const { method, url, query = {} } = req;
    const { msg_signature, timestamp, nonce, echostr } = query;

    console.log("Incoming:", { method, url, query });

    // ---------- 7.1 URL 验证（GET） ----------
    if (method === "GET") {
      if (!echostr) {
        res.status(200).send("ok");
        return;
      }

      if (!msg_signature || !timestamp || !nonce) {
        console.error("GET missing signature params");
        res.status(200).send(echostr);
        return;
      }

      const ok = verifySignature(TOKEN, timestamp, nonce, echostr, msg_signature);
      if (!ok) {
        console.error("GET verify signature failed");
        res.status(200).send(echostr);
        return;
      }

      try {
        const { msg } = decryptWeCom(echostr);
        console.log("GET decrypt echostr success, msg:", msg);
        res.status(200).send(msg);
      } catch (e) {
        console.error("GET decrypt echostr error:", e);
        res.status(200).send(echostr);
      }
      return;
    }

    // ---------- 7.2 接收消息（POST） ----------
    if (method === "POST") {
      try {
        const bodyStr = await readRequestBody(req);
        console.log("raw body:", bodyStr);

        let encrypt;
        try {
          const json = JSON.parse(bodyStr || "{}");
          encrypt = json.encrypt;
        } catch (e) {
          console.error("POST JSON parse error:", e);
          res.status(200).send("invalid json");
          return;
        }

        if (!encrypt) {
          console.error("POST missing encrypt");
          res.status(200).send("missing encrypt");
          return;
        }

        if (!msg_signature || !timestamp || !nonce) {
          console.error("POST missing signature params");
          res.status(200).send("missing signature");
          return;
        }

        const ok = verifySignature(
          TOKEN,
          timestamp,
          nonce,
          encrypt,
          msg_signature
        );
        if (!ok) {
          console.error("POST verify signature failed");
          res.status(200).send("sig error");
          return;
        }

        // 解密 encrypt 得到明文 JSON 字符串
        let plainMsg;
        try {
          const { msg } = decryptWeCom(encrypt);
          plainMsg = msg;
          console.log("decrypt success, plain msg:", plainMsg);
        } catch (e) {
          console.error("decrypt error:", e);
          res.status(200).send("decrypt error");
          return;
        }

        // 解析明文 JSON（用户消息）
        let eventObj = {};
        try {
          eventObj = JSON.parse(plainMsg);
        } catch (e) {
          console.error("plain msg is not valid JSON:", e);
          eventObj = {};
        }

        // 提取用户文本
        let userText = "";
        if (
          eventObj.msgtype === "text" &&
          eventObj.text &&
          typeof eventObj.text.content === "string"
        ) {
          userText = eventObj.text.content;
        }

        // ===== 核心：调用你的业务逻辑“大脑” =====
        const replyContent = await runBusinessLogic(eventObj, userText);

        if (eventObj.response_url) {
          const sent = await sendResponseUrl(eventObj.response_url, replyContent);
          if (sent) {
            res.setHeader("Content-Type", "text/plain; charset=utf-8");
            res.status(200).send("ok");
            return;
          }
        }

        // 构造 stream 明文回复（对齐官方 Demo）
        const streamId =
          eventObj.msgid ||
          (crypto.randomUUID
            ? crypto.randomUUID()
            : crypto.randomBytes(8).toString("hex"));
        const finish = true;

        const replyPlainObj = {
          msgtype: "stream",
          stream: {
            id: streamId,
            finish,
            content: replyContent,
          },
        };

        const replyPlainStr = JSON.stringify(replyPlainObj);
        console.log("reply plain (stream):", replyPlainStr);

        // 加密回复
        const replyPacket = encryptWeCom(replyPlainStr, nonce);
        console.log("replyPacket:", replyPacket);

        res.setHeader("Content-Type", "text/plain; charset=utf-8");
        res.status(200).send(JSON.stringify(replyPacket));
      } catch (e) {
        console.error("POST handler error:", e);
        res.status(200).send("");
      }
      return;
    }

    // 其它方法
    res.status(405).send("Only GET/POST allowed");
  } catch (e) {
    console.error("handler fatal error:", e);
    res.status(500).send("internal error");
  }
};
