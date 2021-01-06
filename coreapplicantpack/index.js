// imports
const { PDFNet } = require('@pdftron/pdfnet-node');
const pdfConvert = async (applicant) => {
    const { personalInfo, applicantDocs } = applicant;
    await PDFNet.initialize();

    // New pdf doc
    const pdfDoc = await PDFNet.PDFDoc.create();

    // create page,write data to it and add to pdfDoc
    const coverPage = await pdfDoc.pageCreate();
    // ElementBuilder is used to build new Element objects
    const eb = await PDFNet.ElementBuilder.create();
    // ElementWriter is used to write Elements to the page
    const writer = await PDFNet.ElementWriter.create();
    let element; 
    // let gstate;

    // begin writing to this page
    writer.beginOnPage(coverPage);
    // Reset the GState to default
    eb.reset();
    // Cover Page. Set Font and font size.
    element = await eb.createTextBeginWithFont(await PDFNet.Font.create(pdfDoc, PDFNet.Font.StandardType1Font.e_times_roman), 16);
    writer.writeElement(element);

    // To-do error checking that required params are there. Otherwise throw an error.
    // Name
    element = await eb.createNewTextRun(`Applicant Name: ${personalInfo.fullName}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 100, 700); // location of text on the page
    writer.writeElement(element); // write element
    writer.writeElement(await eb.createTextNewLine()); // New line

    // Applicant No
    element = await eb.createNewTextRun(`Applicant No: ${personalInfo.applicantNo}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 100, 600);
    writer.writeElement(element);
    writer.writeElement(await eb.createTextNewLine());

    // Channel Type: Internal or External applicant
    element = await eb.createNewTextRun(`Channel Type: ${personalInfo.channelType}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 100, 500);
    writer.writeElement(element);
    writer.writeElement(await eb.createTextNewLine());

    // Job Ref
    element = await eb.createNewTextRun(`Job Reference: ${personalInfo.jobRef}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 100, 400);
    writer.writeElement(element);
    writer.writeElement(await eb.createTextNewLine());

    // Job Title
    element = await eb.createNewTextRun(`Advertised Job Title: ${personalInfo.jobTitle}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 100, 300);
    writer.writeElement(element);
    writer.writeElement(await eb.createTextNewLine());
    
    // gstate = await element.getGState(); gstate used for settings i.e Spacing, Fill Colour etc
    writer.end(); // save changes to the current page
    pdfDoc.pagePushBack(coverPage); // append dynamic page to pdfDoc.

    let pdfDocPageCount;
    // cover letter
    const coverletterPdf = await PDFNet.Convert.office2PDF(applicantDocs.coverLetter);
    // const coverletterPdfPageCount = await coverletterPdf.getPageCount();
    // pdfDocPageCount = await pdfDoc.getPageCount();
    pdfDoc.insertPages(await pdfDoc.getPageCount() + 1, coverletterPdf, 1, await coverletterPdf.getPageCount(), PDFNet.PDFDoc.InsertFlag.e_none);

    // cv
    const cvPdf = await PDFNet.Convert.office2PDF(applicantDocs.cv);
    pdfDoc.insertPages(await pdfDoc.getPageCount() + 1, cvPdf, 1, await cvPdf.getPageCount(), PDFNet.PDFDoc.InsertFlag.e_none);

    // set security handler.
    await pdfDoc.initSecurityHandler();

    // save to a buffer.
    const pdfBuff =  await pdfDoc.saveMemoryBuffer(PDFNet.SDFDoc.SaveOptions.e_linearized);
    if (!pdfBuff) throw Error('Empty Document, potential with saveMemoryBuffer function!');
    return pdfBuff;
};

module.exports = async function (context, req) {
    context.log('Processing a request...');
    // const appliacntName = (req.query.name || (req.body && req.body.name));
    // To-Do: Error checking for all the relevant info?
    if (!req.body) {
        // invalid params, pass error.
        context.res = {
            status: 400,
            body: 'Please pass in request body!'
        };   
    } else if (!req.body.applicant) {
        context.res = {
            status: 400,
            body: 'Please pass in the applicant details!'
        };   

    } else { 
        const applicant = req.body.applicant;
        context.log(`applicant: ${applicant}`);
        context.log('inside else, try pdfConvert...');
            try {
               const memoryBuffer = await pdfConvert(applicant);
               // to-do if memory buffer is null return an error
               context.res = {
                // status: 200, /* Defaults to 200 */
                headers: {
                    "Content-Type" : "application/pdf"
                },
                body: memoryBuffer
            };
            } catch (error) {
                context.log(error);
                // error with pdf generation.
                context.res = {
                    status: 400,
                    body: 'Error creating the applicant PDF. See logs/insights for more info'
                };  
            }
    }
}