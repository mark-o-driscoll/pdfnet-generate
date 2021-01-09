// imports
const { PDFNet } = require('@pdftron/pdfnet-node');
const pdfConvert = async (applicant) => {
    const { personalInfo, applicantDocs, questions } = applicant;
    await PDFNet.initialize();

    // New pdf doc
    const pdfDoc = await PDFNet.PDFDoc.create();

    // ElementBuilder is used to build new Element objects
    const eb = await PDFNet.ElementBuilder.create();
    // ElementWriter is used to write Elements to the page
    const writer = await PDFNet.ElementWriter.create();
    let element; 
    let gstate;
    let questNum;

    // create cover page,write personalInfo to it and add to pdfDoc
    const coverPage = await pdfDoc.pageCreate();
    // begin writing to this page
    writer.beginOnPage(coverPage);
    // Reset the GState to default
    eb.reset();
    // Cover Page. Set Font and font size.
    element = await eb.createTextBeginWithFont(await PDFNet.Font.create(pdfDoc, PDFNet.Font.StandardType1Font.e_times_roman), 16);
    writer.writeElement(element);

    // To-do error checking that required params are there. Otherwise throw an error.
    // Heading
    element = await eb.createNewTextRun('Applicant PDF Pack');
    element.setTextMatrixEntries(2, 0, 0, 2, 180, 720); // location of text on the page
    writer.writeElement(element); // write element

    // Name
    element = await eb.createNewTextRun(`Name: ${personalInfo.fullName}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 80, 650);
    writer.writeElement(element);

    // Applicant No
    element = await eb.createNewTextRun(`Applicant No: ${personalInfo.applicantNo}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 80, 600);
    writer.writeElement(element);

    // Channel Type: Internal or External applicant
    element = await eb.createNewTextRun(`Channel Type: ${personalInfo.channelType}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 80, 550);
    writer.writeElement(element);

    // Job Ref
    element = await eb.createNewTextRun(`Job Reference: ${personalInfo.jobRef}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 80, 500);
    writer.writeElement(element);

    // Job Title
    element = await eb.createNewTextRun(`Advertised Job Title: ${personalInfo.jobTitle}`);
    element.setTextMatrixEntries(1, 0, 0, 1, 80, 450);
    writer.writeElement(element);
    
    // gstate = await element.getGState(); // gstate used for settings i.e Spacing, Fill Colour etc
    writer.end(); // save changes to the current page
    pdfDoc.pagePushBack(coverPage); // append dynamic page to pdfDoc.

    // Screening Questions page
    const screeningQPage = await pdfDoc.pageCreate();
    // begin writing to this page
    writer.beginOnPage(screeningQPage);
    // Reset the GState to default
    eb.reset();

    // Begin writing a block of text
    element = await eb.createTextBeginWithFont(await PDFNet.Font.create(pdfDoc, PDFNet.Font.StandardType1Font.e_times_roman), 14);
    writer.writeElement(element);

    element = await eb.createNewTextRun('Screening Questions');
    element.setTextMatrixEntries(2, 0, 0, 2, 190, 720);
    gstate = await element.getGState();
    // Set the spacing between lines
    gstate.setLeading(16);
    writer.writeElement(element);

    // output questions and answers
    writer.writeElement(await eb.createTextNewLine());
    element.setTextMatrixEntries(1, 0, 0, 1, 50, 705);
    writer.writeElement(element);
    for (let i = 0; i < questions.screening.length; i++) {
        writer.writeElement(await eb.createTextNewLine()); // New line
        questNum = (i + 1).toString(); 
        element = await eb.createNewTextRun(`${questNum}: ${questions.screening[i].question}`);
        gstate = await element.getGState();
        gstate.setTextRenderMode(PDFNet.GState.TextRenderingMode.e_fill_text);
        gstate.setFillColorSpace(await PDFNet.ColorSpace.createDeviceRGB());
        gstate.setFillColorWithColorPt(await PDFNet.ColorPt.init(1, 0, 0)); // red
        writer.writeElement(element);
    
        writer.writeElement(await eb.createTextNewLine()); // New line
    
        element = await eb.createNewTextRun(questions.screening[i].answer);
        gstate = await element.getGState();
        gstate.setTextRenderMode(PDFNet.GState.TextRenderingMode.e_fill_text);
        gstate.setFillColorSpace(await PDFNet.ColorSpace.createDeviceRGB());
        gstate.setFillColorWithColorPt(await PDFNet.ColorPt.init(0, 0, 0)); // black.
        writer.writeElement(element);
    
        writer.writeElement(await eb.createTextNewLine()); // New line
        writer.writeElement(await eb.createTextNewLine()); // New line      
    }     

    //word block test 
    // const para = 'A PDF text object consists of operators that can show '
    // + 'text strings, move the text position, and set text state and certain '
    // + 'other parameters. In addition, there are three parameters that are '
    // + 'defined only within a text object and do not persist from one text '
    // + 'object to the next: Tm, the text matrix, Tlm, the text line matrix, '
    // + 'Trm, the text rendering matrix, actually just an intermediate result '
    // + 'that combines the effects of text state parameters, the text matrix '
    // + '(Tm), and the current transformation matrix';

    // const paraEnd = para.length;
    // let textRun = 0;
    // let textRunEnd;

    // const paraWidth = 500; // paragraph width in units
    // let curWidth = 0;

    // while (textRun < paraEnd) {
    //     textRunEnd = para.indexOf(' ', textRun);
    //     if (textRunEnd < 0) {
    //         textRunEnd = paraEnd - 1;
    //     }
    //     let text = para.substring(textRun, textRunEnd - textRun + 1);
    //     element = await eb.createNewTextRun(text);
    //     if (curWidth + (await element.getTextLength()) < paraWidth) {
    //         writer.writeElement(element);
    //         curWidth += await element.getTextLength();
    //     } else {
    //         writer.writeElement(await eb.createTextNewLine()); // New line
    //         text = para.substr(textRun, textRunEnd - textRun + 1);
    //         element = await eb.createNewTextRun(text);
    //         curWidth = await element.getTextLength();
    //         writer.writeElement(element);
    //     }
    //     textRun = textRunEnd + 1;
    // }
    //word block test

    // Finish the block of text
    writer.writeElement(await eb.createTextEnd());

    writer.end(); // save changes to the current page
    pdfDoc.pagePushBack(screeningQPage);
    // Screening Questions Page(s) End

    // Application Form Questions page
    const applicationQPage = await pdfDoc.pageCreate();
    // begin writing to this page
    writer.beginOnPage(applicationQPage);
    // Reset the GState to default
    eb.reset();
    // Begin writing a block of text
    element = await eb.createTextBeginWithFont(await PDFNet.Font.create(pdfDoc, PDFNet.Font.StandardType1Font.e_times_roman), 14);
    writer.writeElement(element);

    element = await eb.createNewTextRun('Application Form Questions');
    element.setTextMatrixEntries(2, 0, 0, 2, 150, 720);
    gstate = await element.getGState();
    // Set the spacing between lines
    gstate.setLeading(16);
    writer.writeElement(element);

    // output questions and answers
    writer.writeElement(await eb.createTextNewLine());
    element.setTextMatrixEntries(1, 0, 0, 1, 50, 700);
    writer.writeElement(element);
    // questNum already defined. Reset to 0
    questNum = 0;
    for (let i = 0; i < questions.appform.length; i++) {
        writer.writeElement(await eb.createTextNewLine()); // New line
        questNum = (i + 1).toString(); 
        element = await eb.createNewTextRun(`${questNum}: ${questions.appform[i].question}`);
        gstate = await element.getGState();
        gstate.setTextRenderMode(PDFNet.GState.TextRenderingMode.e_fill_text);
        gstate.setFillColorSpace(await PDFNet.ColorSpace.createDeviceRGB());
        gstate.setFillColorWithColorPt(await PDFNet.ColorPt.init(1, 0, 0)); // red
        writer.writeElement(element);
    
        writer.writeElement(await eb.createTextNewLine()); // New line
    
        element = await eb.createNewTextRun(questions.appform[i].answer);
        gstate = await element.getGState();
        gstate.setTextRenderMode(PDFNet.GState.TextRenderingMode.e_fill_text);
        gstate.setFillColorSpace(await PDFNet.ColorSpace.createDeviceRGB());
        gstate.setFillColorWithColorPt(await PDFNet.ColorPt.init(0, 0, 0)); // black.
        writer.writeElement(element);
    
        writer.writeElement(await eb.createTextNewLine()); // New line
        writer.writeElement(await eb.createTextNewLine()); // New line      
    }     

    // Finish the block of text
    writer.writeElement(await eb.createTextEnd());

    writer.end(); // save changes to the current page
    pdfDoc.pagePushBack(applicationQPage);
    // Application Form Questions Page(s) End

    // cover letter
    const coverletterPdf = await PDFNet.Convert.office2PDF(applicantDocs.coverLetter);
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