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

    let job, ext;

    // 1. PDF ke Office
    if (['pdf-to-word', 'pdf-to-excel', 'pdf-to-ppt'].includes(conversionType)) {
      const inputAsset = await pdfServices.upload({ readStream, mimeType: MimeType.PDF });
      
      let targetFormat = ExportPDFTargetFormat.DOCX;
      ext = 'docx';
      if (conversionType === 'pdf-to-excel') { targetFormat = ExportPDFTargetFormat.XLSX; ext = 'xlsx'; }
      if (conversionType === 'pdf-to-ppt') { targetFormat = ExportPDFTargetFormat.PPTX; ext = 'pptx'; }

      const params = new ExportPDFParams({ targetFormat });
      job = new ExportPDFJob({ inputAsset, params });

    // 2. Office ke PDF
    } else if (['word-to-pdf', 'excel-to-pdf', 'ppt-to-pdf'].includes(conversionType)) {
      let mimeType = MimeType.DOCX;
      if (conversionType === 'excel-to-pdf') mimeType = MimeType.XLSX;
      if (conversionType === 'ppt-to-pdf') mimeType = MimeType.PPTX;
      ext = 'pdf';

      const inputAsset = await pdfServices.upload({ readStream, mimeType });
      const params = new CreatePDFParams({});
      job = new CreatePDFJob({ inputAsset, params });
    } else {
      return res.status(400).json({ error: 'Jenis konversi tidak valid.' });
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
