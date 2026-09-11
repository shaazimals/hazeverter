const {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  CreatePDFJob,
  CreatePDFParams,
} = require("@adobe/pdfservices-node-sdk");

const { Readable } = require("stream");
const Busboy = require("busboy");


module.exports.config = {
  api: {
    bodyParser: false,
  },
};


module.exports = async function handler(req, res) {

  if (req.method !== "POST") {
    return res.status(405).json({
      error: "Method not allowed"
    });
  }


  const clientId = process.env.ADOBE_CLIENT_ID;
  const clientSecret = process.env.ADOBE_CLIENT_SECRET;


  if (!clientId || !clientSecret) {

    console.error(
      "Missing Adobe ENV",
      {
        clientId: !!clientId,
        clientSecret: !!clientSecret
      }
    );

    return res.status(500).json({
      error:"Adobe credentials missing"
    });
  }


  try {


    const upload = await new Promise((resolve,reject)=>{


      const busboy = Busboy({
        headers:req.headers
      });


      let bufferChunks=[];
      let conversionType="";


      busboy.on("field",(name,value)=>{

        if(
          [
            "conversionType",
            "type",
            "format",
            "action"
          ].includes(name)
        ){

          conversionType=value;

        }

      });



      busboy.on("file",(name,file)=>{


        file.on("data",(chunk)=>{

          bufferChunks.push(chunk);

        });


        file.on("end",()=>{

        });


      });



      busboy.on("finish",()=>{


        resolve({

          fileBuffer:
          Buffer.concat(bufferChunks),

          conversionType

        });


      });



      busboy.on("error",reject);



      req.pipe(busboy);


    });



    const {
      fileBuffer,
      conversionType
    } = upload;



    if(
      !fileBuffer ||
      fileBuffer.length===0
    ){

      return res.status(400).json({

        error:"File kosong"

      });

    }



    const credentials =
      new ServicePrincipalCredentials({

        clientId,
        clientSecret

      });



    const pdfServices =
      new PDFServices({
        credentials
      });



    const type =
      (conversionType || "pdf-to-word")
      .toLowerCase();



    const readStream =
      Readable.from(fileBuffer);



    let job;
    let resultType;



    /*
      PDF -> Excel
    */

    if(
      type.includes("excel") ||
      type.includes("xlsx")
    ){


      const asset =
        await pdfServices.upload({

          readStream,
          mimeType:MimeType.PDF

        });


      job =
      new ExportPDFJob({

        inputAsset:asset,

        params:
        new ExportPDFParams({

          targetFormat:
          ExportPDFTargetFormat.XLSX

        })

      });


      resultType =
      ExportPDFJob.resultType;


    }



    /*
      PDF -> Word
    */

    else {


      const asset =
      await pdfServices.upload({

        readStream,

        mimeType:MimeType.PDF

      });



      job =
      new ExportPDFJob({

        inputAsset:asset,

        params:
        new ExportPDFParams({

          targetFormat:
          ExportPDFTargetFormat.DOCX

        })

      });


      resultType =
      ExportPDFJob.resultType;


    }



    console.log(
      "Submitting Adobe Job"
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

        resultType

      });



    const asset =
      response.result.asset;



    const content =
      await pdfServices.getContent({

        asset

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
      "Content-Disposition",
      "attachment; filename=result.docx"
    );


    res.setHeader(
      "Content-Type",
      "application/octet-stream"
    );


    return res.status(200).send(output);



  }

  catch(error){


    console.error(
      "FULL ADOBE ERROR",
      error
    );


    return res.status(500).json({

      error:
      error.message ||
      "Adobe processing failed"

    });


  }

};
