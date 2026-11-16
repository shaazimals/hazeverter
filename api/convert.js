/**
 * Hazeverter - Adobe PDF Services API
 * Vercel Serverless Function: /api/convert
 *
 * Environment variables:
 *   ADOBE_CLIENT_ID
 *   ADOBE_CLIENT_SECRET
 * Optional analytics:
 *   UPSTASH_REDIS_REST_URL / KV_REST_API_URL
 *   UPSTASH_REDIS_REST_TOKEN / KV_REST_API_TOKEN
 * Optional admin:
 *   ADMIN_DASHBOARD_PASSWORD
 *   ADOBE_MONTHLY_LIMIT
 */

"use strict";

const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  ExportPDFResult,
  CreatePDFJob,
  CreatePDFResult
} = require("@adobe/pdfservices-node-sdk");

const Busboy = require("busboy");
const { Readable } = require("stream");
const crypto = require("crypto");

const MAX_FILE_SIZE = 20 * 1024 * 1024;
const MAX_FIELDS = 20;
const ADOBE_OPERATIONS = new Set([
  "word-to-pdf", "excel-to-pdf", "ppt-to-pdf", "pdf-export"
]);

const OUTPUTS = {
  docx: {
    adobeFormat: ExportPDFTargetFormat.DOCX,
    extension: ".docx",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  },
  xlsx: {
    adobeFormat: ExportPDFTargetFormat.XLSX,
    extension: ".xlsx",
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  },
  pptx: {
    adobeFormat: ExportPDFTargetFormat.PPTX,
    extension: ".pptx",
    mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  },
  pdf: { extension: ".pdf", mime: "application/pdf" }
};

const OFFICE_INPUTS = {
  "word-to-pdf": {
    mimeType: MimeType.DOCX,
    label: "Word DOCX",
    extensions: [".docx"]
  },
  "excel-to-pdf": {
    mimeType: MimeType.XLSX,
    label: "Excel XLSX",
    extensions: [".xlsx"]
  },
  "ppt-to-pdf": {
    mimeType: MimeType.PPTX,
    label: "PowerPoint PPTX",
    extensions: [".pptx"]
  }
};

class HttpError extends Error {
  constructor(message, statusCode = 400, code = "REQUEST_ERROR") {
    super(message);
    this.name = "HttpError";
    this.statusCode = statusCode;
    this.code = code;
  }
}

async function handler(req, res) {
  setCommonHeaders(res);

  if (req.method === "OPTIONS") {
    res.setHeader("Allow", "GET, POST, OPTIONS");
    return res.status(204).end();
  }

  if (req.method === "GET" && req.query?.admin === "stats") {
    return handleAdminStats(req, res);
  }

  if (req.method === "POST" && req.query?.event) {
    return handleAnalyticsEvent(req, res);
  }

  if (req.method !== "POST") {
    res.setHeader("Allow", "GET, POST, OPTIONS");
    return res.status(405).json({
      error: "Method tidak didukung.",
      code: "METHOD_NOT_ALLOWED"
    });
  }

  const startedAt = Date.now();
  let operation = "unknown";

  try {
    const clientId = String(process.env.ADOBE_CLIENT_ID || "").trim();
    const clientSecret = String(process.env.ADOBE_CLIENT_SECRET || "").trim();

    if (!clientId || !clientSecret) {
      throw new HttpError(
        "Kredensial Adobe belum dikonfigurasi. Periksa ADOBE_CLIENT_ID dan ADOBE_CLIENT_SECRET di Environment Variables.",
        503,
        "ADOBE_CREDENTIALS_MISSING"
      );
    }

    const upload = await readMultipart(req);
    operation = normalizeOperation(req.query?.type || upload.fields.conversionType);

    const officeInput = OFFICE_INPUTS[operation];
    const format = officeInput
      ? "pdf"
      : normalizeFormat(req.query?.format || upload.fields.format || "docx");
    const outputConfig = OUTPUTS[format];

    validateUploadedFile(upload, operation, officeInput);

    console.info("[Hazeverter] Conversion started", {
      operation,
      format,
      fileName: upload.originalName,
      fileSize: upload.buffer.length
    });

    const credentials = new ServicePrincipalCredentials({
      clientId,
      clientSecret
    });

    const pdfServices = new PDFServices({ credentials });

    // PPTX uses the original PPTX asset directly. Do not rasterize, resize,
    // or rebuild slides here: Adobe performs the Office-to-PDF rendering.
    const inputAsset = await pdfServices.upload({
      readStream: Readable.from(upload.buffer),
      mimeType: officeInput ? officeInput.mimeType : MimeType.PDF
    });

    let job;
    let resultType;

    if (officeInput) {
      job = new CreatePDFJob({ inputAsset });
      resultType = CreatePDFResult;
    } else {
      const params = new ExportPDFParams({
        targetFormat: outputConfig.adobeFormat
      });
      job = new ExportPDFJob({ inputAsset, params });
      resultType = ExportPDFResult;
    }

    const pollingURL = await pdfServices.submit({ job });
    if (!pollingURL) {
      throw new Error("Adobe tidak memberikan alamat pemantauan proses konversi.");
    }

    const jobResponse = await pdfServices.getJobResult({
      pollingURL,
      resultType
    });

    const resultAsset = jobResponse?.result?.asset;
    if (!resultAsset) {
      throw new Error("Proses Adobe selesai tanpa mengembalikan file hasil.");
    }

    const content = await pdfServices.getContent({ asset: resultAsset });
    if (!content?.readStream) {
      throw new Error("File hasil Adobe tidak dapat dibaca (stream kosong).");
    }

    const chunks = [];
    let outputSize = 0;
    for await (const chunk of content.readStream) {
      const part = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
      outputSize += part.length;

      // Protect the serverless function from unexpectedly large responses.
      if (outputSize > 100 * 1024 * 1024) {
        throw new HttpError("Ukuran hasil konversi melebihi batas 100 MB.", 502, "OUTPUT_TOO_LARGE");
      }
      chunks.push(part);
    }

    const finalBuffer = Buffer.concat(chunks);
    validateAdobeOutput(finalBuffer, format);

    const outputName = createOutputName(upload.originalName, outputConfig.extension);

    res.setHeader("Content-Type", outputConfig.mime);
    res.setHeader("Content-Disposition", contentDisposition(outputName));
    res.setHeader("Content-Length", String(finalBuffer.length));
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("X-Hazeverter-Conversion", operation);

    console.info("[Hazeverter] Conversion succeeded", {
      operation,
      format,
      outputSize: finalBuffer.length,
      durationMs: Date.now() - startedAt
    });

    // Analytics must never turn a successful document conversion into a failure.
    recordAdobeSuccess(operation).catch(error => {
      console.error("[Hazeverter] Analytics write failed:", safeErrorMessage(error));
    });

    return res.status(200).send(finalBuffer);
  } catch (error) {
    const normalized = normalizeAdobeError(error);
    console.error("[Hazeverter] Conversion failed", {
      operation,
      code: normalized.code,
      status: normalized.status,
      message: normalized.message,
      details: safeErrorDetails(error),
      durationMs: Date.now() - startedAt
    });

    if (res.headersSent) {
      return res.end();
    }

    return res.status(normalized.status).json({
      error: normalized.message,
      code: normalized.code,
      name: error?.name || "Error",
      retryable: normalized.retryable
    });
  }
}

function setCommonHeaders(res) {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("Referrer-Policy", "same-origin");
}

function normalizeOperation(value) {
  const normalized = String(value || "").trim().toLowerCase();

  // Frontend sends these operation names when exporting a PDF to Office.
  // They all use Adobe ExportPDFJob; the requested target is selected by
  // the separate `format` field (docx, xlsx, or pptx).
  if (["pdf-to-word", "pdf-to-excel", "pdf-to-ppt"].includes(normalized)) {
    return "pdf-export";
  }

  if (ADOBE_OPERATIONS.has(normalized)) return normalized;

  throw new HttpError(
    `Jenis konversi "${normalized || "(kosong)"}" tidak didukung.`,
    400,
    "UNSUPPORTED_OPERATION"
  );
}

function normalizeFormat(value) {
  const normalized = String(value || "docx").trim().toLowerCase();
  if (["docx", "word"].includes(normalized)) return "docx";
  if (["xlsx", "excel"].includes(normalized)) return "xlsx";
  if (["pptx", "ppt", "powerpoint"].includes(normalized)) return "pptx";
  throw new HttpError("Format output tidak didukung. Pilih DOCX, XLSX, atau PPTX.", 400, "UNSUPPORTED_FORMAT");
}

function validateUploadedFile(upload, operation, officeInput) {
  if (!upload?.buffer?.length) {
    throw new HttpError("File tidak ditemukan atau berkas yang dikirim kosong.", 400, "EMPTY_UPLOAD");
  }

  if (upload.buffer.length > MAX_FILE_SIZE) {
    throw new HttpError("Ukuran file maksimal 20 MB.", 413, "FILE_TOO_LARGE");
  }

  if (officeInput) {
    if (!hasAllowedExtension(upload.originalName, officeInput.extensions)) {
      throw new HttpError(
        `Format input harus ${officeInput.extensions.join(" atau ")}.`,
        415,
        "INVALID_FILE_EXTENSION"
      );
    }
    if (!isZipOfficeFile(upload.buffer)) {
      throw new HttpError(
        `File ${officeInput.label} tidak valid atau rusak. Pastikan file benar-benar ${officeInput.extensions[0]}.`,
        415,
        "INVALID_OFFICE_FILE"
      );
    }
    return;
  }

  if (operation === "pdf-export" && upload.buffer.subarray(0, 5).toString("ascii") !== "%PDF-") {
    throw new HttpError("File yang dikirim bukan PDF yang valid.", 415, "INVALID_PDF");
  }
}

function validateAdobeOutput(buffer, format) {
  if (!buffer?.length) {
    throw new Error(`Adobe mengembalikan file ${format.toUpperCase()} kosong.`);
  }

  if (format === "pdf" && buffer.subarray(0, 5).toString("ascii") !== "%PDF-") {
    throw new Error("Adobe mengembalikan hasil yang bukan PDF valid.");
  }

  if (format !== "pdf" && !(buffer[0] === 0x50 && buffer[1] === 0x4b)) {
    throw new Error(`Adobe mengembalikan hasil ${format.toUpperCase()} yang tidak valid.`);
  }
}

function readMultipart(req) {
  return new Promise((resolve, reject) => {
    let parser;
    try {
      parser = Busboy({
        headers: req.headers,
        limits: {
          files: 1,
          fileSize: MAX_FILE_SIZE,
          fields: MAX_FIELDS,
          fieldSize: 16 * 1024
        }
      });
    } catch (error) {
      reject(new HttpError(`Request upload tidak valid: ${error.message}`, 400, "INVALID_MULTIPART"));
      return;
    }

    const chunks = [];
    const fields = Object.create(null);
    let fileFound = false;
    let fileTooLarge = false;
    let originalName = "uploaded-file";
    let settled = false;

    const fail = error => {
      if (settled) return;
      settled = true;
      reject(error);
    };

    parser.on("field", (name, value, info) => {
      if (info?.valueTruncated) {
        fail(new HttpError("Data form terlalu panjang.", 400, "FIELD_TOO_LARGE"));
        return;
      }
      fields[name] = value;
    });

    parser.on("file", (fieldName, file, info) => {
      if (fileFound) {
        file.resume();
        fail(new HttpError("Hanya satu file yang boleh dikirim dalam satu proses.", 400, "TOO_MANY_FILES"));
        return;
      }

      fileFound = true;
      originalName = String(info?.filename || "uploaded-file").split(/[\\/]/).pop();

      file.on("data", chunk => chunks.push(Buffer.from(chunk)));
      file.on("limit", () => { fileTooLarge = true; });
      file.on("error", fail);
    });

    parser.on("filesLimit", () => {
      fail(new HttpError("Request berisi terlalu banyak file.", 400, "TOO_MANY_FILES"));
    });
    parser.on("fieldsLimit", () => {
      fail(new HttpError("Request berisi terlalu banyak field.", 400, "TOO_MANY_FIELDS"));
    });
    parser.on("error", error => {
      fail(new HttpError(`Gagal membaca upload: ${error.message}`, 400, "MULTIPART_READ_ERROR"));
    });

    parser.on("finish", () => {
      if (settled) return;
      if (!fileFound) {
        fail(new HttpError("File tidak ditemukan pada field upload 'file'.", 400, "FILE_FIELD_MISSING"));
        return;
      }
      if (fileTooLarge) {
        fail(new HttpError("Ukuran file maksimal 20 MB.", 413, "FILE_TOO_LARGE"));
        return;
      }
      settled = true;
      resolve({
        buffer: Buffer.concat(chunks),
        originalName,
        fields
      });
    });

    req.on("aborted", () => fail(new HttpError("Upload dibatalkan sebelum selesai.", 400, "UPLOAD_ABORTED")));
    req.pipe(parser);
  });
}

function isZipOfficeFile(buffer) {
  return buffer.length >= 4 &&
    buffer[0] === 0x50 &&
    buffer[1] === 0x4b &&
    [0x03, 0x05, 0x07].includes(buffer[2]) &&
    [0x04, 0x06, 0x08].includes(buffer[3]);
}

function hasAllowedExtension(filename, extensions) {
  const lower = String(filename || "").toLowerCase();
  return extensions.some(extension => lower.endsWith(extension));
}

function createOutputName(originalName, extension) {
  const base = String(originalName || "converted")
    .replace(/\.[^.]+$/, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 100);
  return `${base || "converted"}${extension}`;
}

function contentDisposition(filename) {
  const ascii = filename.replace(/[^\x20-\x7E]/g, "_").replace(/["\\]/g, "_");
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(filename)}`;
}

function safeErrorMessage(error) {
  return String(error?.message || "Unknown error").slice(0, 500);
}

function safeErrorDetails(error) {
  const details = {};
  for (const key of ["name", "message", "status", "statusCode", "code", "requestId"]) {
    if (error?.[key] !== undefined) details[key] = String(error[key]).slice(0, 500);
  }
  if (error?.response?.status) details.responseStatus = error.response.status;
  return details;
}

function normalizeAdobeError(error) {
  if (error instanceof HttpError) {
    return {
      status: error.statusCode,
      code: error.code,
      message: error.message,
      retryable: error.statusCode >= 500
    };
  }

  const message = safeErrorMessage(error);
  const lower = message.toLowerCase();
  const statusFromError = Number(error?.status || error?.statusCode || error?.response?.status || 0);

  if (/unauthorized|invalid.client|invalid_client|authentication|credential|token/.test(lower) || statusFromError === 401) {
    return {
      status: 502,
      code: "ADOBE_AUTH_FAILED",
      message: "Autentikasi ke Adobe gagal. Periksa pasangan ADOBE_CLIENT_ID dan ADOBE_CLIENT_SECRET serta status kredensial di Adobe Developer Console.",
      retryable: false
    };
  }

  if (/quota|rate.?limit|too many requests|429/.test(lower) || statusFromError === 429) {
    return {
      status: 503,
      code: "ADOBE_RATE_LIMIT",
      message: "Layanan Adobe sedang mencapai batas penggunaan. Tunggu beberapa saat lalu coba kembali.",
      retryable: true
    };
  }

  if (/timeout|timed out|etimedout|econnreset|eai_again|socket hang up/.test(lower)) {
    return {
      status: 504,
      code: "ADOBE_TIMEOUT",
      message: "Adobe membutuhkan waktu terlalu lama atau koneksi terputus saat memproses dokumen. Coba lagi beberapa saat; file asli tidak diubah.",
      retryable: true
    };
  }

  if (/unsupported|not supported|invalid.*(file|document)|corrupt|damaged|password.protected/.test(lower)) {
    return {
      status: 422,
      code: "ADOBE_UNSUPPORTED_DOCUMENT",
      message: "Adobe tidak dapat membaca dokumen ini. Pastikan file tidak rusak, tidak diproteksi password, dan disimpan dalam format Office/PDF yang didukung.",
      retryable: false
    };
  }

  if (statusFromError >= 500 || statusFromError === 0) {
    return {
      status: 502,
      code: "ADOBE_CONVERSION_FAILED",
      message: "Layanan Adobe gagal menyelesaikan konversi. Coba kembali. Jika berulang, periksa log Vercel untuk detail Adobe request.",
      retryable: true
    };
  }

  return {
    status: 502,
    code: "ADOBE_CONVERSION_FAILED",
    message: "Konversi gagal diproses oleh Adobe. Periksa format dan kondisi dokumen, lalu coba kembali.",
    retryable: false
  };
}

/* ----------------------------- Analytics ----------------------------- */

function analyticsConfig() {
  return {
    url: (process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL || "").replace(/\/$/, ""),
    token: process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN || ""
  };
}

function dateParts() {
  const now = new Date();
  return {
    day: now.toISOString().slice(0, 10),
    month: now.toISOString().slice(0, 7)
  };
}

async function redisPipeline(commands) {
  const config = analyticsConfig();
  if (!config.url || !config.token) {
    throw new Error("Database statistik belum dikonfigurasi.");
  }

  const response = await fetch(`${config.url}/pipeline`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${config.token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(commands)
  });

  if (!response.ok) {
    throw new Error(`Database statistik merespons HTTP ${response.status}.`);
  }
  return response.json();
}

async function handleAnalyticsEvent(req, res) {
  res.setHeader("Cache-Control", "no-store");
  const event = String(req.query.event || "").toLowerCase();
  const { day, month } = dateParts();

  try {
    if (event === "visit") {
      await redisPipeline([
        ["INCR", "hz:visits:total"],
        ["INCR", `hz:visits:${day}`]
      ]);
    } else if (event === "tool_run") {
      const tool = String(req.query.tool || "unknown")
        .toLowerCase()
        .replace(/[^a-z0-9-]/g, "")
        .slice(0, 40) || "unknown";
      await redisPipeline([
        ["INCR", "hz:tools:total"],
        ["INCR", `hz:tools:${month}`],
        ["INCR", `hz:tool:${month}:${tool}`]
      ]);
    } else {
      return res.status(400).json({ error: "Event statistik tidak didukung." });
    }

    return res.status(204).end();
  } catch (error) {
    console.error("[Hazeverter] Analytics event failed:", safeErrorMessage(error));
    // Analytics is non-critical. The frontend can safely ignore this response.
    return res.status(503).json({ error: "Penyimpanan statistik belum tersedia." });
  }
}

async function recordAdobeSuccess(operation) {
  const { month } = dateParts();
  await redisPipeline([
    ["INCR", "hz:adobe:total"],
    ["INCR", `hz:adobe:${month}`],
    ["INCR", `hz:adobe:${month}:${operation}`]
  ]);
}

function safePasswordEqual(value, expected) {
  const a = Buffer.from(String(value || ""));
  const b = Buffer.from(String(expected || ""));
  return a.length === b.length && a.length > 0 && crypto.timingSafeEqual(a, b);
}

async function handleAdminStats(req, res) {
  res.setHeader("Cache-Control", "no-store, private");
  const expectedPassword = process.env.ADMIN_DASHBOARD_PASSWORD || "";

  if (!expectedPassword || !safePasswordEqual(req.headers["x-admin-password"], expectedPassword)) {
    return res.status(401).json({ error: "Password admin salah." });
  }

  const { day, month } = dateParts();
  const tools = [
    "ocr-ba", "pdf-to-word", "pdf-to-excel", "pdf-to-ppt",
    "word-to-pdf", "excel-to-pdf", "ppt-to-pdf", "pdf-to-jpg",
    "jpg-to-pdf", "html-to-pdf", "split-pdf", "organize-pdf",
    "merge", "edit-pdf"
  ];

  const recentDays = Array.from({ length: 7 }, (_, index) => {
    const date = new Date();
    date.setUTCDate(date.getUTCDate() - (6 - index));
    return date.toISOString().slice(0, 10);
  });

  const commands = [
    ["GET", "hz:visits:total"],
    ["GET", `hz:visits:${day}`],
    ["GET", "hz:tools:total"],
    ["GET", `hz:tools:${month}`],
    ["GET", "hz:adobe:total"],
    ["GET", `hz:adobe:${month}`],
    ...tools.map(tool => ["GET", `hz:tool:${month}:${tool}`]),
    ...recentDays.map(date => ["GET", `hz:visits:${date}`])
  ];

  try {
    const result = await redisPipeline(commands);
    const values = result.map(item => Number(item?.result || 0));
    const [visitorsTotal, visitorsToday, toolsTotal, toolsMonth, adobeTotal, adobeMonth] = values;
    const toolValues = values.slice(6, 6 + tools.length);
    const dailyValues = values.slice(6 + tools.length);
    const monthlyLimit = Math.max(0, Number(process.env.ADOBE_MONTHLY_LIMIT || 0));

    return res.status(200).json({
      period: month,
      visitors: { total: visitorsTotal, today: visitorsToday },
      processes: { total: toolsTotal, month: toolsMonth },
      adobe: {
        total: adobeTotal,
        month: adobeMonth,
        limit: monthlyLimit,
        remaining: monthlyLimit > 0 ? Math.max(0, monthlyLimit - adobeMonth) : null,
        percentage: monthlyLimit > 0
          ? Math.min(100, Math.round((adobeMonth / monthlyLimit) * 100))
          : null
      },
      byTool: Object.fromEntries(tools.map((tool, index) => [tool, toolValues[index] || 0])),
      visitorsLast7Days: recentDays.map((date, index) => ({
        date,
        value: dailyValues[index] || 0
      }))
    });
  } catch (error) {
    console.error("[Hazeverter] Admin stats failed:", safeErrorMessage(error));
    return res.status(503).json({
      error: "Database statistik belum dikonfigurasi atau tidak dapat diakses."
    });
  }
}

// IMPORTANT: Vercel must not parse multipart request bodies before Busboy.
// Attach config to the handler AFTER defining it. Assigning module.exports.config
// before replacing module.exports with the function silently loses this setting.
handler.config = {
  api: {
    bodyParser: false
  }
};

module.exports = handler;
