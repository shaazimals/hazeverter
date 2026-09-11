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

const { Readable } = require('stream');
const Busboy = require('busboy');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const clientId = (process.env.ADOBE_CLIENT_ID || '').trim();
  const clientSecret = (process.env.ADOBE_CLIENT_SECRET || '').trim();

  if (!clientId || !clientSecret) {
    return res.status(500).json({ error: 'Kredensial ADOBE_CLIENT_ID / ADOBE_CLIENT_SECRET belum dipasang di Vercel.' });
  }

  try {
    const { fileBuffer, conversionType } = await new Promise((resolve, reject) => {
      const busboy = Busboy({ headers: req.headers });
      let fileBuffer = null;
      let conversionType = '';

      busboy.on('field', (fieldname, val) => {
        if (fieldname === 'conversionType') conversionType = val;
      });

      busboy.on('file', (fieldname, file) => {
        const buffers = [];
        file.on('data', (data) => buffers.push(data));
        file.on('end', () => { fileBuffer = Buffer.concat(buffers); });
      });

      busboy.on('finish', () => resolve({ fileBuffer, conversionType }));
      busboy.on('error', (err) => reject(err));

      req.pipe(busboy);
    });

    if (!fileBuffer) {
      return res.status(400).json({ error: 'Tidak ada berkas yang dikirimkan.' });
    }

    const credentials = new ServicePrincipalCredentials({ clientId, clientSecret });
    const pdfServices = new PDFServices({ credentials });
    const readStream = Readable.from(fileBuffer);

    let job;
    const type = (conversionType || '').toLowerCase();

    // 1. PDF ke Word (DOCX)
    if (type.includes('word') || type.includes('docx') || type === 'pdf-to-word') {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.PDF });
      const params = new ExportPDFParams({ targetFormat: ExportPDFTargetFormat.DOCX });
      job = new ExportPDFJob({ inputAsset, params });

    // 2. PDF ke Excel (XLSX)
    } else if (type.includes('excel') || type.includes('xlsx') || type === 'pdf-to-excel') {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.PDF });
      const params = new ExportPDFParams({ targetFormat: ExportPDFTargetFormat.XLSX });
      job = new ExportPDFJob({ inputAsset, params });

    // 3. PDF ke PowerPoint (PPTX)
    } else if (type.includes('ppt') || type.includes('pptx') || type === 'pdf-to-ppt') {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.PDF });
      const params = new ExportPDFParams({ targetFormat: ExportPDFTargetFormat.PPTX });
      job = new ExportPDFJob({ inputAsset, params });

    // 4. Word ke PDF
    } else if (type === 'word-to-pdf' || type === 'docx-to-pdf') {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.DOCX });
      job = new CreatePDFJob({ inputAsset, params: new CreatePDFParams({}) });

    // 5. Excel ke PDF
    } else if (type === 'excel-to-pdf' || type === 'xlsx-to-pdf') {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.XLSX });
      job = new CreatePDFJob({ inputAsset, params: new CreatePDFParams({}) });

    // 6. PowerPoint ke PDF
    } else if (type === 'ppt-to-pdf' || type === 'pptx-to-pdf') {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.PPTX });
      job = new CreatePDFJob({ inputAsset, params: new CreatePDFParams({}) });

    } else {
      return res.status(400).json({ error: `Tipe konversi '${conversionType}' tidak dikenali oleh server.` });
    }

    const pollingURL = await pdfServices.submit(job);
    const pdfServicesResponse = await pdfServices.getJobResult({
      pollingURL,
      resultType: ExportPDFJob.resultType,
    });

    const resultAsset = pdfServicesResponse.result.asset;
    const streamAsset = await pdfServices.getContent({ asset: resultAsset });

    const resultBuffers = [];
    for await (const chunk of streamAsset.readStream) {
      resultBuffers.push(chunk);
    }
    const finalBuffer = Buffer.concat(resultBuffers);

    res.setHeader('Content-Type', 'application/octet-stream');
    return res.status(200).send(finalBuffer);

  } catch (error) {
    console.error('Adobe Execution Error:', error);
    return res.status(500).json({ error: 'Gagal diproses server Adobe: ' + error.message });
  }
};
