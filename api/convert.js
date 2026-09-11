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
    bodyParser: false,
  },
};


module.exports = async function handler(req, res) {

  if (req.method !== "POST") {
    return res.status(405).json({
      error: "Method not allowed"
    });
  }


  const clientId = process.env.ADOBE_CLIENT_ID?.trim();
  const clientSecret = process.env.ADOBE_CLIENT_SECRET?.trim();


  if (!clientId || !clientSecret) {
    return res.status(500).json({
      error: "Adobe credentials belum tersedia"
    });
  }


  try {


    const {
      fileBuffer,
      conversionType
    } = await new Promise((resolve,reject)=>{


      const busboy = Busboy({
        headers:req.headers
      });


      let chunks=[];
      let type="";


      busboy.on("file",(name,file)=>{

        file.on("data",(data)=>{
          chunks.push(data);
        });

      });


      busboy.on("field",(name,value)=>{

        if(
          name === "conversionType" ||
          name === "type" ||
          name === "format"
        ){
          type=value;
        }

      });


      busboy.on("finish",()=>{

        resolve({
          fileBuffer:Buffer.concat(chunks),
          conversionType:type
        });

      });


      busboy.on("error",reject);


      req.pipe(busboy);


    });



    if(!fileBuffer || fileBuffer.length === 0){

      return res.status(400).json({
        error:"File PDF kosong"
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



    const asset =
      await pdfServices.upload({

        readStream:
        Readable.from(fileBuffer),

        mimeType:
        MimeType.PDF

      });



    /*
       Default:
       PDF -> Word
    */


    let targetFormat =
      ExportPDFTargetFormat.DOCX;



    if(
      conversionType &&
      conversionType.toLowerCase().includes("excel")
    ){

      targetFormat =
      ExportPDFTargetFormat.XLSX;

    }



    const params =
      new ExportPDFParams({

        targetFormat

      });



    const job =
      new ExportPDFJob({

        inputAsset:asset,

        params

      });



    if(!job){

      throw new Error(
        "Adobe job gagal dibuat"
      );

    }



    console.log(
      "Adobe job created"
    );


    const pollingURL =
      await pdfServices.submit(job);



    const response =
      await pdfServices.getJobResult({

        pollingURL,

        resultType:
        ExportPDFJob.resultType

      });



    const resultAsset =
      response.result.asset;



    const content =
      await pdfServices.getContent({

        asset:resultAsset

      });



    const output=[];


    for await(
      const chunk of content.readStream
    ){

      output.push(chunk);

    }


    const finalBuffer =
      Buffer.concat(output);



    res.setHeader(
      "Content-Type",
      "application/octet-stream"
    );


    res.setHeader(
      "Content-Disposition",
      'attachment; filename="converted.docx"'
    );


    return res.status(200)
      .send(finalBuffer);



  } catch(error){


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
