// imports
const { PDFNet } = require('@pdftron/pdfnet-node');
const pdfConvert = async (appliacntName) => {
    await PDFNet.initialize();

    // New pdf doc
    const pdfDoc = await PDFNet.PDFDoc.create();

    // create page,write data to it and add to pdfDoc
    const page2 = await pdfDoc.pageCreate();
    // ElementBuilder is used to build new Element objects
    const eb = await PDFNet.ElementBuilder.create();
    // ElementWriter is used to write Elements to the page
    const writer = await PDFNet.ElementWriter.create();
    let element; 
    let gstate;
    // begin writing to this page
    writer.beginOnPage(page2);
    // Reset the GState to default
    eb.reset();
    // Begin writing a block of text. Set Font and font size.
    element = await eb.createTextBeginWithFont(await PDFNet.Font.create(pdfDoc, PDFNet.Font.StandardType1Font.e_times_roman), 16);
    writer.writeElement(element);
    // write text
    element = await eb.createNewTextRun(`Applicant Name: ${appliacntName}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 100, 700); // location on page of text
    gstate = await element.getGState();
    writer.writeElement(element);
    writer.end(); // save changes to the current page
    pdfDoc.pagePushBack(page2); // push dynamic page to pdfDoc.

    let pdfDocPageCount;

    // cover letter
    const coverletterPdf = await PDFNet.Convert.office2PDF('https://cdn.geckoandfly.com/wp-content/downloads/cover-letter-templates/cover-letter-free-template-03.docx');
    const coverletterPdfPageCount = await coverletterPdf.getPageCount();
    pdfDocPageCount = await pdfDoc.getPageCount();
    pdfDoc.insertPages(pdfDocPageCount + 1, coverletterPdf, 1, coverletterPdfPageCount, PDFNet.PDFDoc.InsertFlag.e_none);

    // cv
    const cvPdf = await PDFNet.Convert.office2PDF('https://www.coolfreecv.com/doc/coolfreecv_resume_en_03_n.docx');
    const cvPdfPageCount = await cvPdf.getPageCount();
    pdfDocPageCount = await pdfDoc.getPageCount();
    pdfDoc.insertPages(pdfDocPageCount + 1, cvPdf, 1, cvPdfPageCount, PDFNet.PDFDoc.InsertFlag.e_none);

    // docx to pdf from url tests
    // 23kb file (no images)
    // const wordDocPdf = await PDFNet.Convert.office2PDF('https://www.coolfreecv.com/doc/coolfreecv_resume_en_06_n.docx');

    // 72kb file
    // const wordDocPdf = await PDFNet.Convert.office2PDF('https://www.coolfreecv.com/doc/coolfreecv_resume_en_03_n.docx');
    
    // 1mb word doc
    // const wordDocPdf = await PDFNet.Convert.office2PDF('https://file-examples-com.github.io/uploads/2017/02/file-sample_1MB.docx');

    // 13.51mb word doc
    // const wordDocPdf = await PDFNet.Convert.office2PDF('https://filesamples.com/samples/document/docx/sample4.docx');

    // merge pages. Append wordDocPdf by passing 1 for first page and count for last page.
    // const wordDocPdfPageCount = await wordDocPdf.getPageCount();
    // const pdfDocPageCount = await pdfDoc.getPageCount();
    // pdfDoc.insertPages(pdfDocPageCount + 1, wordDocPdf, 1, wordDocPdfPageCount, PDFNet.PDFDoc.InsertFlag.e_none);

    // load test to compare to aspose: 5 sample docs at 115kb (23kb each)
    // for (let i = 0; i < 5; i++) {
    //     const wordDocPdf = await PDFNet.Convert.office2PDF('https://www.coolfreecv.com/doc/coolfreecv_resume_en_06_n.docx');
    //     // merge pages. Append wordDocPdf by passing 1 for first page and count for last page.
    //     const wordDocPdfPageCount = await wordDocPdf.getPageCount();
    //     const pdfDocPageCount = await pdfDoc.getPageCount();
    //     pdfDoc.insertPages(pdfDocPageCount + 1, wordDocPdf, 1, wordDocPdfPageCount, PDFNet.PDFDoc.InsertFlag.e_none);
    //   }

    // set security handler.
    await pdfDoc.initSecurityHandler();

    // save to a buffer.
    const pdfBuff =  await pdfDoc.saveMemoryBuffer(PDFNet.SDFDoc.SaveOptions.e_linearized);
    return pdfBuff;  
};

module.exports = async function (context, req) {
    context.log('Processing a request...');

    const appliacntName = (req.query.name || (req.body && req.body.name));
    const responseMessage = appliacntName
        ? "Hello, " + appliacntName + ". This HTTP triggered function executed successfully."
        : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";
    if (!appliacntName) {
        // invalid params, pass error.
        context.res = {
            status: 400,
            body: 'Please Pass in the name!!'
        };   

    }
    else {    
        context.log('inside else...');
            try {
                context.log('inside try...');
               const memoryBuffer = await pdfConvert(appliacntName);
               context.log(typeof memoryBuffer);
               context.res = {
                // status: 200, /* Defaults to 200 */
                headers: {
                    "Content-Type" : "application/pdf"
                },
                body: memoryBuffer
            };
            } catch (error) {
                context.log(error);
            }
    }
}