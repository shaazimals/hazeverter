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
      error:"Method not allowed"
    });
  }


  try {


    const clientId =
      process.env.ADOBE_CLIENT_ID?.trim();


    const clientSecret =
      process.env.ADOBE_CLIENT_SECRET?.trim();



    if(!clientId || !clientSecret){

      throw new Error(
        "Adobe ENV tidak ditemukan"
      );

    }



    const fileBuffer =
      await new Promise((resolve,reject)=>{


        const busboy =
          Busboy({
            headers:req.headers
          });


        let chunks=[];


        busboy.on(
          "file",
          (field,file)=>{


            file.on(
              "data",
              chunk=>{
                chunks.push(chunk);
              }
            );


          }
        );


        busboy.on(
          "finish",
          ()=>{

            resolve(
              Buffer.concat(chunks)
            );

          }
        );


        busboy.on(
          "error",
          reject
        );


        req.pipe(busboy);


      });



    if(
      !fileBuffer ||
      fileBuffer.length === 0
    ){

      throw new Error(
        "PDF kosong"
      );

    }



    console.log(
      "PDF SIZE:",
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
      "Uploading PDF..."
    );



    const asset =
      await pdfServices.upload({

        readStream:
          Readable.from(fileBuffer),

        mimeType:
          MimeType.PDF

      });



    console.log(
      "Creating Adobe Job..."
    );



    const params =
      new ExportPDFParams({

        targetFormat:
          ExportPDFTargetFormat.DOCX

      });



    const job =
      new ExportPDFJob({

        inputAsset: asset,

        params

      });



    console.log(
      "JOB:",
      job ? "OK" : "FAILED"
    );



    const pollingURL =
      await pdfServices.submit(job);



    console.log(
      "POLLING:",
      pollingURL
    );



    const result =
      await pdfServices.getJobResult({

        pollingURL,

        resultType:
          ExportPDFJob.resultType

      });



    const content =
      await pdfServices.getContent({

        asset:
          result.result.asset

      });



    const chunks=[];


    for await(
      const chunk of content.readStream
    ){

      chunks.push(chunk);

    }



    const output =
      Buffer.concat(chunks);



    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );


    res.setHeader(
      "Content-Disposition",
      "attachment; filename=result.docx"
    );


    return res.status(200).send(output);



  } catch(err){


    console.error(
      "FULL ERROR:",
      err
    );


    return res.status(500).json({

      error:
        err.message

    });

  }

};
