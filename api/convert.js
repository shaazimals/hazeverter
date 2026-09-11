const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  ExportPDFResult
} = require("@adobe/pdfservices-node-sdk");

const Busboy = require("busboy");
const { Readable } = require("stream");

const MAX_FILE_SIZE = 20 * 1024 * 1024;

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

    const { buffer: fileBuffer, originalName } = await readUploadedPdf(req);

    if (!fileBuffer.length) {
      throw new Error("PDF tidak terbaca");
    }

    // Validasi signature PDF, bukan hanya ekstensi atau MIME dari browser.
    if (fileBuffer.subarray(0, 5).toString("ascii") !== "%PDF-") {
      throw new Error("File yang dikirim bukan PDF yang valid");
    }

    console.log("PDF TO EXCEL - FILE SIZE:", fileBuffer.length);

    const credentials = new ServicePrincipalCredentials({
      clientId,
      clientSecret
    });

    const pdfServices = new PDFServices({ credentials });

    console.log("PDF TO EXCEL - UPLOAD START");
    const inputAsset = await pdfServices.upload({
      readStream: Readable.from(fileBuffer),
      mimeType: MimeType.PDF
    });
    console.log("PDF TO EXCEL - UPLOAD SUCCESS");

    const params = new ExportPDFParams({
      targetFormat: ExportPDFTargetFormat.XLSX
    });

    const job = new ExportPDFJob({
      inputAsset,
      params
    });

    console.log("PDF TO EXCEL - SUBMIT START");
    const pollingURL = await pdfServices.submit({ job });

    const response = await pdfServices.getJobResult({
      pollingURL,
      resultType: ExportPDFResult
    });

    const resultAsset = response?.result?.asset;
    if (!resultAsset) {
      throw new Error("Adobe tidak mengembalikan file Excel hasil konversi");
    }

    const content = await pdfServices.getContent({
      asset: resultAsset
    });

    if (!content?.readStream) {
      throw new Error("Stream file Excel dari Adobe tidak tersedia");
    }

    const outputChunks = [];
    for await (const chunk of content.readStream) {
      outputChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }

    const finalBuffer = Buffer.concat(outputChunks);
    if (!finalBuffer.length) {
      throw new Error("File Excel hasil konversi kosong");
    }

    const outputName = createOutputName(originalName);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${outputName}"`
    );
    res.setHeader("Content-Length", finalBuffer.length);
    res.setHeader("Cache-Control", "no-store");

    console.log("PDF TO EXCEL - SUCCESS:", finalBuffer.length);
    return res.status(200).send(finalBuffer);
  } catch (error) {
    console.error("FULL PDF TO EXCEL ERROR:", error);

    const message = error?.message || "Gagal mengubah PDF menjadi Excel";
    const status =
      /tidak ditemukan|tidak terbaca|bukan PDF|terlalu besar|multipart/i.test(message)
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

function readUploadedPdf(req) {
  return new Promise((resolve, reject) => {
    let busboy;

    try {
      busboy = Busboy({
        headers: req.headers,
        limits: {
          files: 1,
          fileSize: MAX_FILE_SIZE
        }
      });
    } catch (error) {
      reject(new Error(`Form upload tidak valid: ${error.message}`));
      return;
    }

    const chunks = [];
    let fileFound = false;
    let originalName = "converted.pdf";
    let tooLarge = false;

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

    busboy.on("filesLimit", () => {
      console.warn("Lebih dari satu file dikirim; hanya file pertama diproses");
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
        originalName
      });
    });

    busboy.on("error", reject);
    req.pipe(busboy);
  });
}

function createOutputName(originalName) {
  const baseName = String(originalName || "converted")
    .replace(/\.pdf$/i, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 100);

  return `${baseName || "converted"}.xlsx`;
}
