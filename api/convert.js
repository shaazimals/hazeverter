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
        "Adobe credential tidak ditemukan"
      );
    }


    const fileBuffer = await new Promise((resolve, reject) => {

      const busboy = Busboy({
        headers: req.headers
      });

      const chunks = [];


      busboy.on("file", (name, file) => {

        file.on("data", chunk => {
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
        "PDF tidak terbaca"
      );

    }



    console.log(
      "FILE SIZE:",
      fileBuffer.length
    );



    const credentials =
      new ServicePrincipalCredentials({

        clientId,
        clientSecret

      });



    const pdfServices =
      new PDFServices({

        credentials

      });



    console.log(
      "UPLOAD START"
    );


    const inputAsset =
      await pdfServices.upload({

        readStream:
          Readable.from(fileBuffer),

        mimeType:
          MimeType.PDF

      });



    console.log(
      "UPLOAD SUCCESS"
    );



    const params =
      new ExportPDFParams({

        targetFormat:
          ExportPDFTargetFormat.DOCX

      });



    console.log(
      "CREATE JOB"
    );


    const job =
      new ExportPDFJob({

        inputAsset,
        params

      });



    console.log(
      "JOB:",
      job
    );



    console.log(
      "SUBMIT START"
    );


    const pollingURL =
      await pdfServices.submit({

        job

      });



    console.log(
      "POLLING URL:",
      pollingURL
    );



    const response =
      await pdfServices.getJobResult({

        pollingURL,

        resultType:
          ExportPDFJob.resultType

      });



    console.log(
      "JOB FINISHED"
    );



    const resultAsset =
      response.result.asset;



    const content =
      await pdfServices.getContent({

        asset: resultAsset

      });



    const output = [];


    for await (const chunk of content.readStream) {

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


    return res.status(200).send(finalBuffer);



  } catch (error) {


    console.error(
      "FULL ADOBE ERROR:",
      error
    );


    return res.status(500).json({

      error: error.message,

      name: error.name,

      details:
        JSON.stringify(
          error,
          Object.getOwnPropertyNames(error)
        )

    });

  }

};
