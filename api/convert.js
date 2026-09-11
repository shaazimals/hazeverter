const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  CreatePDFJob,
  CreatePDFParams
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

  // 1. Ambil Environment Variables
  const clientId = (process.env.ADOBE_CLIENT_ID || '').trim();
  const clientSecret = (process.env.ADOBE_CLIENT_SECRET || '').trim();

  if (!clientId || !clientSecret) {
    return res.status(500).json({ 
      error: 'Environment Variable Hilang: ADOBE_CLIENT_ID atau ADOBE_CLIENT_SECRET tidak terdeteksi di Vercel.' 
    });
  }

  const form = formidable({
    uploadDir: '/tmp',
    keepExtensions: true,
  });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error('Formidable Error:', err);
      return res.status(500).json({ error: 'Gagal membaca berkas unggahan: ' + err.message });
    }

    try {
      const uploadedFile = Array.isArray(files.file) ? files.file[0] : files.file;
      const conversionType = Array.isArray(fields.conversionType) ? fields.conversionType[0] : fields.conversionType;

      if (!uploadedFile || !uploadedFile.filepath) {
        return res.status(400).json({ error: 'Tidak ada berkas yang dikirimkan.' });
      }

      const inputFilePath = uploadedFile.filepath;
      let outputFilePath = path.join('/tmp', `converted_${Date.now()}`);

      // 2. Inisialisasi Adobe Service Principal Credentials
      const credentials = new ServicePrincipalCredentials({
        clientId: clientId,
        clientSecret: clientSecret,
      });

      const pdfServices = new PDFServices({ credentials });
      const readStream = fs.createReadStream(inputFilePath);

      // 3. Proses Konversi PDF ke Office (Word/Excel/PPT)
      if (['pdf-to-word', 'pdf-to-excel', 'pdf-to-ppt'].includes(conversionType)) {
        const inputAsset = await pdfServices.upload({
          readStream,
          mimeType: MimeType.PDF,
        });

        let targetFormat = ExportPDFTargetFormat.DOCX;
        let ext = '.docx';
        if (conversionType === 'pdf-to-excel') { targetFormat = ExportPDFTargetFormat.XLSX; ext = '.xlsx'; }
        if (conversionType === 'pdf-to-ppt') { targetFormat = ExportPDFTargetFormat.PPTX; ext = '.pptx'; }

        const params = new ExportPDFParams({ targetFormat });
        const job = new ExportPDFJob({ inputAsset, params });

        const pollingURL = await pdfServices.submit(job);
        const pdfServicesResponse = await pdfServices.getJobResult({
          pollingURL,
          resultType: ExportPDFJob.resultType,
        });

        const resultAsset = pdfServicesResponse.result.asset;
        const streamAsset = await pdfServices.getContent({ asset: resultAsset });

        outputFilePath += ext;
        const outputStream = fs.createWriteStream(outputFilePath);

        await new Promise((resolve, reject) => {
          streamAsset.readStream.pipe(outputStream);
          streamAsset.readStream.on('end', resolve);
          streamAsset.readStream.on('error', reject);
        });

      // 4. Proses Konversi Office (Word/Excel/PPT) ke PDF
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
        return res.status(400).json({ error: 'Tipe konversi tidak valid: ' + conversionType });
      }

      // 5. Kirimkan Hasil Berkas ke Client
      const fileBuffer = fs.readFileSync(outputFilePath);

      // Cleanup berkas sementara di /tmp
      try {
        if (fs.existsSync(inputFilePath)) fs.unlinkSync(inputFilePath);
        if (fs.existsSync(outputFilePath)) fs.unlinkSync(outputFilePath);
      } catch (e) {}

      res.setHeader('Content-Type', 'application/octet-stream');
      return res.status(200).send(fileBuffer);

    } catch (error) {
      console.error('Adobe API Runtime Failure:', error);
      const errorMsg = error.message || 'Gagal mengeksekusi konversi pada Adobe API.';
      return res.status(500).json({ error: `[Adobe API Error]: ${errorMsg}` });
    }
  });
}
