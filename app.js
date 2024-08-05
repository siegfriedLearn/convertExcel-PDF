const controladorExcel = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const puppeteer = require('puppeteer');
const pug = require('pug');





const processFiles = () => {

    //Info del archivo viene en req.file
    // const rutaArchivo = req.file.path
    // const rutaArchivo = '/Users/carlosmedina/Downloads/test.xlsx'
    const rutaArchivo = '/Users/carlosmedina/Downloads/filtrado.xlsx'
    //console.log(req.file)
    //Leer archivo
    const archivoExcel = controladorExcel.readFile(rutaArchivo)
    //leer hojas del archivo
    const hojasArchivo = archivoExcel.SheetNames;
    //console.log(hojasArchivo)
    //almacenar nombre primera hoja
    const nombreHoja = hojasArchivo[0]
    //guardar info en formato json
    const infoHojaJson = controladorExcel.utils.sheet_to_json(archivoExcel.Sheets[nombreHoja])
    //console.log(infoHojaJson)
    // res.status(200).download(
    //     'public/files/plantilla.xlsx'
    //     );

    return infoHojaJson;

    }

     const infoHoja =  processFiles();
     //console.log(infoHoja)

     //buildPDF((data)=>{},(data)=>{})

// const crearDocumento = async (nombre, propiedades) =>{

//     // Cargar plantilla
//     const compiledFunction = pug.compileFile('./views/ticket.pug');

//     // Enviar propiedades
//     const html = compiledFunction(propiedades);
//     // console.log(html)    

//     const browser = await puppeteer.launch();
//     const page = await browser.newPage();
//     await page.setContent(html);
//     await page.pdf({path: `./documentos/${nombre}.pdf`, format: 'A4'});
//     await browser.close();
// }    

// infoHoja.forEach((ticket, index) => {
//     crearDocumento(ticket.TICKET, ticket)
// });


function buildPDF(data){
    const doc = new PDFDocument();

    doc.fontSize(12).text(
        `ticket: ${data.TICKET}      
descripcion: ${data.DESCRIPCION}
tipo: ${data.TIPO}
grupo_resolutor: ${data.RESOLUTORGROUP}
clasificacion_nivel_3: ${data.ID_CLASIFICACION3}
cumple_fcr: ${data.CUMPLEFCR}
detalle: ${data.DETALLE}
solucion: ${data.SOLUCION}`
)

    doc.pipe(fs.createWriteStream(`./documentos/${data.TICKET}.pdf`))
    doc.end();
}


infoHoja.forEach((ticket) => {
    buildPDF(ticket)
});




   