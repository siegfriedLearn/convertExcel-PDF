const controladorExcel = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const puppeteer = require('puppeteer');
const pug = require('pug');





const processFiles = () => {

    //Info del archivo viene en req.file
    // const rutaArchivo = req.file.path
    // const rutaArchivo = '/Users/carlosmedina/Downloads/test.xlsx'
    const rutaArchivo = '/Users/carlosmedina/Downloads/SR.xlsx'
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


function buildPDF(){
    const doc = new PDFDocument();

    infoHoja.forEach((ticket)=>{
        doc.fontSize(10).text(`INICIO SECCIÓN`, {align: 'center'})
        doc.fontSize(8).text(`TICKET: ${ticket.TICKET}`);      
    doc.fontSize(8).text(`DESCRIPCION: ${ticket.DESCRIPCION.replace(/\t/g, ' ')}`);
    doc.fontSize(8).text(`TIPO: ${ticket.TIPO}`);
    doc.fontSize(8).text(`GRUPO_RESOLUTOR: ${ticket.RESOLUTORGROUP}`);
    doc.fontSize(8).text(`CLASIFICACION_NIVEL_3: ${ticket.ID_CLASIFICACION3}`);
    doc.fontSize(8).text(`CUMPLE_FCR: ${ticket.CUMPLEFCR}`);
    doc.fontSize(8).text(`DETALLE: ${ticket.DETALLE.replace(/\r\n|\r/g, '\n').replace(/\t/g, ' ')}`);
    doc.fontSize(8).text(`SOLUCION: ${ticket.SOLUCION.replace(/\r\n|\r/g, '\n').replace(/\t/g, ' ')}`);
    doc.fontSize(10).text(`FIN SECCIÓN`, {align: 'center'})
    //doc.addPage()
    })
   

    doc.pipe(fs.createWriteStream(`SR.pdf`))
    doc.end();
}


// infoHoja.forEach((ticket) => {
//     buildPDF(ticket)
// });

buildPDF()



   