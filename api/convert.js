const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  CreatePDFJob,
  CreatePDFParams,
  SDKError,
  ServiceUsageError,
  ServiceApiError
} = require('@adobe/pdfservices-node-sdk');

const fs = require('fs');
const path = require('path');
const formidable = require('formidable');

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // 1. Cek Environment Variables
  const clientId = process.env.ADOBE_CLIENT_ID;
  const clientSecret = process.env.ADOBE_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    return res.status(500).json({ 
      error: 'Konfigurasi Server Gagal: ADOBE_CLIENT_ID atau ADOBE_CLIENT_SECRET belum dipasang di Environment Variables Vercel.' 
    });
  }

  const form = formidable({
    uploadDir: '/tmp',
    keepExtensions: true,
  });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      return res.status(500).json({ error: 'Gagal memproses file upload: ' + err.message });
    }

    const uploadedFile = Array.isArray(files.file) ? files.file[0] : files.file;
    const conversionType = Array.isArray(fields.conversionType) ? fields.conversionType[0] : fields.conversionType;

    if (!uploadedFile) {
      return res.status(400).json({ error: 'Tidak ada berkas yang diunggah.' });
    }

    let inputFilePath = uploadedFile.filepath;
    let outputFilePath = path.join('/tmp', `output_${Date.now()}`);

    try {
      // 2. Inisialisasi Adobe SDK
      const credentials = new ServicePrincipalCredentials({
        clientId: clientId.trim(),
        clientSecret: clientSecret.trim(),
      });

      const pdfServices = new PDFServices({ credentials });

      let readStream = fs.createReadStream(inputFilePath);

      // 3. Eksekusi Konversi Sesuai Opsi
      if (['pdf-to-word', 'pdf-to-excel', 'pdf-to-ppt'].includes(conversionType)) {
        const inputAsset = await pdfServices.upload({
          readStream,
          mimeType: MimeType.PDF,
        });

        let targetFormat = ExportPDFTargetFormat.DOCX;
        if (conversionType === 'pdf-to-excel') targetFormat = ExportPDFTargetFormat.XLSX;
        if (conversionType === 'pdf-to-ppt') targetFormat = ExportPDFTargetFormat.PPTX;

        const params = new ExportPDFParams({ targetFormat });
        const job = new ExportPDFJob({ inputAsset, params });

        const pollingURL = await pdfServices.submit(job);
        const pdfServicesResponse = await pdfServices.getJobResult({
          pollingURL,
          resultType: ExportPDFJob.resultType,
        });

        const resultAsset = pdfServicesResponse.result.asset;
        const streamAsset = await pdfServices.getContent({ asset: resultAsset });

        outputFilePath += (conversionType === 'pdf-to-word' ? '.docx' : conversionType === 'pdf-to-excel' ? '.xlsx' : '.pptx');
        const outputStream = fs.createWriteStream(outputFilePath);

        await new Promise((resolve, reject) => {
          streamAsset.readStream.pipe(outputStream);
          streamAsset.readStream.on('end', resolve);
          streamAsset.readStream.on('error', reject);
        });

      } else if (['word-to-pdf', 'excel-to-pdf', 'ppt-to-pdf'].includes(conversionType)) {
        let mimeType = MimeType.DOCX;
        if (conversionType === 'excel-to-pdf') mimeType = MimeType.XLSX;
        if (conversionType === 'ppt-to-pdf') mimeType = MimeType.PPTX;

        const inputAsset = await pdfServices.upload({ readStream, mimeType });
        const params = new CreatePDFParams({});
        const job = new CreatePDFJob({ inputAsset, params });

        const pollingURL = await pdfServices.submit(job);
        const pdfServicesResponse = await pdfServices.getJobResult({
          pollingURL,
          resultType: CreatePDFJob.resultType,
        });

        const resultAsset = pdfServicesResponse.result.asset;
        const streamAsset = await pdfServices.getContent({ asset: resultAsset });

        outputFilePath += '.pdf';
        const outputStream = fs.createWriteStream(outputFilePath);

        await new Promise((resolve, reject) => {
          streamAsset.readStream.pipe(outputStream);
          streamAsset.readStream.on('end', resolve);
          streamAsset.readStream.on('error', reject);
        });
      } else {
        return res.status(400).json({ error: 'Tipe konversi tidak dikenali: ' + conversionType });
      }

      // 4. Kirim File Hasil Konversi ke Client
      const fileBuffer = fs.readFileSync(outputFilePath);
      
      // Cleanup temporary files
      if (fs.existsSync(inputFilePath)) fs.unlinkSync(inputFilePath);
      if (fs.existsSync(outputFilePath)) fs.unlinkSync(outputFilePath);

      res.setHeader('Content-Type', 'application/octet-stream');
      return res.status(200).send(fileBuffer);

    } catch (error) {
      console.error('Adobe API Runtime Error:', error);
      
      // Tangkap Error Spesifik Adobe SDK
      let detailMessage = error.message || 'Terjadi kesalahan sistem pada Adobe SDK.';
      if (error instanceof SDKError || error instanceof ServiceApiError || error instanceof ServiceUsageError) {
        detailMessage = `[Adobe API Error]: ${error.message} (Status Code: ${error.statusCode || 'Unknown'})`;
      }

      return res.status(500).json({ error: detailMessage });
    }
  });
}
