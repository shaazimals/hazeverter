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

const MAX_FILE_SIZE = 20 * 1024 * 1024;

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
  pdf: {
    extension: ".pdf",
    mime: "application/pdf"
  }
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

module.exports.config = {
  api: {
    bodyParser: false
  }
};

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const clientId = process.env.ADOBE_CLIENT_ID?.trim();
    const clientSecret = process.env.ADOBE_CLIENT_SECRET?.trim();

    if (!clientId || !clientSecret) {
      throw new Error("Adobe credential tidak ditemukan");
    }

    const upload = await readMultipart(req);
    const operation = normalizeOperation(
      req.query?.type || upload.fields.conversionType
    );
    const officeInput = OFFICE_INPUTS[operation];
    const format = officeInput
      ? "pdf"
      : normalizeFormat(req.query?.format || upload.fields.format || "docx");
    const outputConfig = OUTPUTS[format];

    if (!upload.buffer.length) {
      throw new Error("Berkas tidak terbaca");
    }

    if (operation === "pdf-export" && upload.buffer.subarray(0, 5).toString("ascii") !== "%PDF-") {
      throw new Error("File yang dikirim bukan PDF yang valid");
    }

    if (officeInput && !isZipOfficeFile(upload.buffer)) {
      throw new Error(`File ${officeInput.label} tidak valid`);
    }

    if (officeInput && !hasAllowedExtension(upload.originalName, officeInput.extensions)) {
      throw new Error(`Format input harus ${officeInput.extensions.join(" atau ")}`);
    }

    console.log("CONVERT START:", {
      operation,
      format,
      fileSize: upload.buffer.length
    });

    const credentials = new ServicePrincipalCredentials({
      clientId,
      clientSecret
    });

    const pdfServices = new PDFServices({ credentials });

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

    const response = await pdfServices.getJobResult({
      pollingURL,
      resultType
    });

    const resultAsset = response?.result?.asset;
    if (!resultAsset) {
      throw new Error(`Adobe tidak mengembalikan file ${format.toUpperCase()}`);
    }

    const content = await pdfServices.getContent({ asset: resultAsset });
    if (!content?.readStream) {
      throw new Error(`Stream hasil ${format.toUpperCase()} tidak tersedia`);
    }

    const chunks = [];
    for await (const chunk of content.readStream) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }

    const finalBuffer = Buffer.concat(chunks);
    if (!finalBuffer.length) {
      throw new Error(`File ${format.toUpperCase()} hasil konversi kosong`);
    }

    const outputName = createOutputName(
      upload.originalName,
      outputConfig.extension
    );

    res.setHeader("Content-Type", outputConfig.mime);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${outputName}"`
    );
    res.setHeader("Content-Length", finalBuffer.length);
    res.setHeader("Cache-Control", "no-store");

    console.log("CONVERT SUCCESS:", {
      format,
      outputSize: finalBuffer.length
    });

    return res.status(200).send(finalBuffer);
  } catch (error) {
    console.error("FULL CONVERT ERROR:", error);

    const message = error?.message || "Gagal memproses dokumen";
    const status =
      /format|tidak ditemukan|tidak terbaca|bukan PDF|DOCX|XLSX|PPTX|tidak valid|terlalu besar|multipart/i.test(message)
        ? 400
        : 500;

    return res.status(status).json({
      error: message,
      name: error?.name || "Error",
      details: JSON.stringify(
        error,
        Object.getOwnPropertyNames(error || {})
      )
    });
  }
};

function normalizeOperation(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["word-to-pdf", "excel-to-pdf", "ppt-to-pdf"].includes(normalized)) {
    return normalized;
  }
  return "pdf-export";
}

function normalizeFormat(value) {
  const normalized = String(value || "docx").trim().toLowerCase();

  if (["docx", "word"].includes(normalized)) return "docx";
  if (["xlsx", "excel"].includes(normalized)) return "xlsx";
  if (["pptx", "ppt", "powerpoint"].includes(normalized)) return "pptx";

  throw new Error("Format output tidak didukung. Gunakan docx, xlsx, atau pptx");
}

function readMultipart(req) {
  return new Promise((resolve, reject) => {
    let busboy;

    try {
      busboy = Busboy({
        headers: req.headers,
        limits: {
          files: 1,
          fileSize: MAX_FILE_SIZE,
          fields: 10
        }
      });
    } catch (error) {
      reject(new Error(`Form upload tidak valid: ${error.message}`));
      return;
    }

    const chunks = [];
    const fields = {};
    let fileFound = false;
    let originalName = "converted.pdf";
    let tooLarge = false;

    busboy.on("field", (name, value) => {
      fields[name] = value;
    });

    busboy.on("file", (_fieldName, file, info) => {
      if (fileFound) {
        file.resume();
        return;
      }

      fileFound = true;
      originalName = info?.filename || "converted.pdf";

      file.on("data", chunk => chunks.push(chunk));
      file.on("limit", () => {
        tooLarge = true;
      });
      file.on("error", reject);
    });

    busboy.on("finish", () => {
      if (!fileFound) {
        reject(new Error("File PDF tidak ditemukan pada request multipart"));
        return;
      }

      if (tooLarge) {
        reject(new Error("Ukuran PDF terlalu besar. Maksimal 20 MB"));
        return;
      }

      resolve({
        buffer: Buffer.concat(chunks),
        originalName,
        fields
      });
    });

    busboy.on("error", reject);
    req.pipe(busboy);
  });
}

function createOutputName(originalName, extension) {
  const baseName = String(originalName || "converted")
    .replace(/\.pdf$/i, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 100);

  return `${baseName || "converted"}${extension}`;
}

function isZipOfficeFile(buffer) {
  return buffer.length >= 4 && buffer[0] === 0x50 && buffer[1] === 0x4B;
}

function hasAllowedExtension(filename, allowedExtensions) {
  const lowerName = String(filename || "").toLowerCase();
  return allowedExtensions.some(extension => lowerName.endsWith(extension));
}
