const { 
  PDFServices, 
  Credentials, 
  ExportPDFJob, 
  ExportPDFParams, 
  ExportPDFTargetFormat,
  CreatePDFJob
} = require('@adobe/pdfservices-node-sdk');
const formidable = require('formidable');
const fs = require('fs');
const { Readable } = require('stream');

module.exports.config = {
  api: {
    bodyParser: false,
  },
};

const MIME_TYPES = {
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pdf: 'application/pdf'
};

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let tempFilePath = null;

  try {
    const form = formidable({ uploadDir: '/tmp', keepExtensions: true });
    
    // Parse form data
    const [fields, files] = await new Promise((resolve, reject) => {
      form.parse(req, (err, fields, files) => {
        if (err) reject(err);
        else resolve([fields, files]);
      });
    });

    const clientId = process.env.ADOBE_CLIENT_ID;
    const clientSecret = process.env.ADOBE_CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      return res.status(500).json({ error: 'API Key Adobe belum diatur di Vercel Environment Variables.' });
    }

    const credentials = Credentials.servicePrincipalCredentialsBuilder()
      .withClientId(clientId)
      .withClientSecret(clientSecret)
      .build();

    const pdfServices = new PDFServices({ credentials });

    const fileItem = Array.isArray(files.file) ? files.file[0] : files.file;
    const rawType = Array.isArray(fields.conversionType) ? fields.conversionType[0] : fields.conversionType;
    const rawTarget = Array.isArray(fields.targetType) ? fields.targetType[0] : fields.targetType;
    
    const conversionType = rawType || (rawTarget ? `pdf-to-${rawTarget}` : 'pdf-to-word');

    if (!fileItem || !fileItem.filepath) {
      return res.status(400).json({ error: 'File tidak terdeteksi.' });
    }

    tempFilePath = fileItem.filepath;
    const originalName = fileItem.originalFilename || '';
    const ext = originalName.split('.').pop().toLowerCase();

    // Baca file ke Buffer agar aman di serverless
    const fileBuffer = fs.readFileSync(tempFilePath);

    let job, contentTypeOutput, outputFilename;

    // 1. PDF TO (WORD / PPT / EXCEL)
    if (['pdf-to-excel', 'pdf-to-word', 'pdf-to-ppt'].includes(conversionType)) {
      const inputAsset = await pdfServices.upload({ 
        stream: Readable.from(fileBuffer), 
        mimeType: MIME_TYPES.pdf 
      });

      let targetFormat = ExportPDFTargetFormat.DOCX;
      if (conversionType === 'pdf-to-excel') targetFormat = ExportPDFTargetFormat.XLSX;
      if (conversionType === 'pdf-to-ppt') targetFormat = ExportPDFTargetFormat.PPTX;

      const params = new ExportPDFParams({ targetFormat });
      job = new ExportPDFJob({ inputAsset, params });

      if (conversionType === 'pdf-to-excel') {
        contentTypeOutput = MIME_TYPES.xlsx;
        outputFilename = 'converted.xlsx';
      } else if (conversionType === 'pdf-to-ppt') {
        contentTypeOutput = MIME_TYPES.pptx;
        outputFilename = 'converted.pptx';
      } else {
        contentTypeOutput = MIME_TYPES.docx;
        outputFilename = 'converted.docx';
      }
    } 
    // 2. (EXCEL / WORD / PPT) TO PDF
    else if (['excel-to-pdf', 'word-to-pdf', 'ppt-to-pdf'].includes(conversionType)) {
      let mimeTypeInput = MIME_TYPES[ext];
      if (!mimeTypeInput) {
        if (conversionType === 'excel-to-pdf') mimeTypeInput = MIME_TYPES.xlsx;
        else if (conversionType === 'word-to-pdf') mimeTypeInput = MIME_TYPES.docx;
        else if (conversionType === 'ppt-to-pdf') mimeTypeInput = MIME_TYPES.pptx;
      }

      const inputAsset = await pdfServices.upload({ 
        stream: Readable.from(fileBuffer), 
        mimeType: mimeTypeInput 
      });

      job = new CreatePDFJob({ inputAsset });

      contentTypeOutput = MIME_TYPES.pdf;
      outputFilename = 'converted.pdf';
    } else {
      return res.status(400).json({ error: 'Jenis konversi tidak valid.' });
    }

    // Polling Job Result dari Adobe Cloud
    const pollingURL = await pdfServices.submit({ job });
    const jobClass = ['excel-to-pdf', 'word-to-pdf', 'ppt-to-pdf'].includes(conversionType) ? CreatePDFJob : ExportPDFJob;
    
    const pdfServicesResponse = await pdfServices.getJobResult({ 
      pollingURL, 
      resultType: jobClass 
    });

    const resultAsset = pdfServicesResponse.result?.asset;
    if (!resultAsset) throw new Error('Adobe Cloud tidak mengembalikan hasil file.');

    const streamAsset = await pdfServices.getContent({ asset: resultAsset });

    res.setHeader('Content-Type', contentTypeOutput);
    res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);

    const resultStream = streamAsset.readStream || streamAsset.stream;
    
    if (resultStream && typeof resultStream.pipe === 'function') {
      resultStream.pipe(res);
      resultStream.on('end', () => {
        if (tempFilePath) try { fs.unlinkSync(tempFilePath); } catch (e) {}
      });
    } else if (streamAsset) {
      res.send(Buffer.from(streamAsset));
      if (tempFilePath) try { fs.unlinkSync(tempFilePath); } catch (e) {}
    }

  } catch (error) {
    console.error('Adobe Serverless Error:', error);
    if (tempFilePath) try { fs.unlinkSync(tempFilePath); } catch (e) {}
    res.status(500).json({ error: error.message || 'Terjadi kesalahan saat mengonversi dokumen.' });
  }
};
