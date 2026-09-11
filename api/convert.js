const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat
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
    return res.status(405).json({
      error: "Method not allowed"
    });
  }


  try {

    const clientId =
      process.env.ADOBE_CLIENT_ID?.trim();

    const clientSecret =
      process.env.ADOBE_CLIENT_SECRET?.trim();


    if (!clientId || !clientSecret) {

      throw new Error(
        "ADOBE_CLIENT_ID / ADOBE_CLIENT_SECRET belum tersedia"
      );

    }



    // ==========================
    // READ UPLOADED FILE
    // ==========================

    const fileBuffer = await new Promise((resolve, reject) => {


      const busboy = Busboy({
        headers: req.headers
      });


      let chunks = [];


      busboy.on("file", (fieldname, file) => {


        file.on("data", (chunk) => {
          chunks.push(chunk);
        });


        file.on("error", reject);


      });



      busboy.on("finish", () => {


        resolve(
          Buffer.concat(chunks)
        );


      });



      busboy.on("error", reject);


      req.pipe(busboy);


    });



    if (!fileBuffer || fileBuffer.length === 0) {

      throw new Error(
        "File PDF tidak terbaca"
      );

    }



    console.log(
      "PDF RECEIVED:",
      fileBuffer.length
    );



    // ==========================
    // ADOBE AUTH
    // ==========================


    const credentials =
      new ServicePrincipalCredentials({

        clientId,

        clientSecret

      });



    const pdfServices =
      new PDFServices({

        credentials

      });



    // ==========================
    // UPLOAD PDF
    // ==========================


    console.log(
      "Uploading asset..."
    );


    const inputAsset =
      await pdfServices.upload({

        readStream:
          Readable.from(fileBuffer),

        mimeType:
          MimeType.PDF

      });



    console.log(
      "Asset uploaded"
    );



    // ==========================
    // CREATE JOB
    // ==========================


    const params =
      new ExportPDFParams({

        targetFormat:
          ExportPDFTargetFormat.DOCX

      });



    const job =
      new ExportPDFJob({

        inputAsset,

        params

      });



    console.log(
      "JOB CREATED:",
      job !== undefined
    );



    if (!job) {

      throw new Error(
        "Adobe ExportPDFJob gagal dibuat"
      );

    }



    // ==========================
    // EXECUTE JOB
    // ==========================


    console.log(
      "Submitting job..."
    );


    const pollingURL =
      await pdfServices.submit(job);



    console.log(
      "Polling:",
      pollingURL
    );



    const response =
      await pdfServices.getJobResult({

        pollingURL,

        resultType:
          ExportPDFJob.resultType

      });



    // ==========================
    // DOWNLOAD RESULT
    // ==========================


    const resultAsset =
      response.result.asset;



    const result =
      await pdfServices.getContent({

        asset: resultAsset

      });



    const output = [];


    for await (
      const chunk of result.readStream
    ) {

      output.push(chunk);

    }


    const finalBuffer =
      Buffer.concat(output);



    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );


    res.setHeader(
      "Content-Disposition",
      'attachment; filename="converted.docx"'
    );


    return res
      .status(200)
      .send(finalBuffer);



  } catch (error) {


    console.error(
      "ADOBE ERROR:",
      error
    );


    return res.status(500).json({

      error:
        error.message ||
        "Adobe conversion gagal"

    });


  }

};
