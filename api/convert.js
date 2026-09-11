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

    const fileBuffer = await readUploadedPdf(req);

    if (!fileBuffer.length) {
      throw new Error("PDF tidak terbaca");
    }

    console.log("FILE SIZE:", fileBuffer.length);

    const credentials = new ServicePrincipalCredentials({
      clientId,
      clientSecret
    });

    const pdfServices = new PDFServices({ credentials });

    console.log("UPLOAD START");
    const inputAsset = await pdfServices.upload({
      readStream: Readable.from(fileBuffer),
      mimeType: MimeType.PDF
    });
    console.log("UPLOAD SUCCESS");

    const params = new ExportPDFParams({
      targetFormat: ExportPDFTargetFormat.DOCX
    });

    const job = new ExportPDFJob({
      inputAsset,
      params
    });

    console.log("SUBMIT START");
    const pollingURL = await pdfServices.submit({ job });
    console.log("POLLING URL RECEIVED");

    // PENTING: resultType harus memakai ExportPDFResult,
    // bukan ExportPDFJob.resultType yang nilainya undefined.
    const response = await pdfServices.getJobResult({
      pollingURL,
      resultType: ExportPDFResult
    });

    const resultAsset = response?.result?.asset;
    if (!resultAsset) {
      throw new Error("Adobe tidak mengembalikan file DOCX hasil konversi");
    }

    console.log("JOB FINISHED");

    const content = await pdfServices.getContent({
      asset: resultAsset
    });

    const output = [];
    for await (const chunk of content.readStream) {
      output.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }

    const finalBuffer = Buffer.concat(output);
    if (!finalBuffer.length) {
      throw new Error("File DOCX hasil konversi kosong");
    }

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="converted.docx"'
    );
    res.setHeader("Content-Length", finalBuffer.length);

    return res.status(200).send(finalBuffer);
  } catch (error) {
    console.error("FULL ADOBE ERROR:", error);

    return res.status(500).json({
      error: error?.message || "Gagal memproses dokumen",
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
      busboy = Busboy({ headers: req.headers });
    } catch (error) {
      reject(new Error(`Form upload tidak valid: ${error.message}`));
      return;
    }

    const chunks = [];
    let fileFound = false;

    busboy.on("file", (_fieldName, file) => {
      fileFound = true;
      file.on("data", chunk => chunks.push(chunk));
      file.on("error", reject);
    });

    busboy.on("finish", () => {
      if (!fileFound) {
        reject(new Error("File PDF tidak ditemukan pada request"));
        return;
      }
      resolve(Buffer.concat(chunks));
    });

    busboy.on("error", reject);
    req.pipe(busboy);
  });
}
