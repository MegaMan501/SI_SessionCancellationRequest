// Author: Mohamed Rahaman
// Date: April 28, 2018
// Email: nabilrahaman501@gmail.com

/******************************* Global Variables *******************************/
var officePhoneNumber = ''; // NE SI Office Phone Number
var supervisorName = '';
var supervisorEmail = '';   // The person recieving the cancellation request. Only use a TCCD email.

/*
    This were information about the
    learning centers can be changed.
    Be carefull with the information
    you change here because it can
    will affect the rest of the code.
*/
var learningCenter = {
  science:{
           courseNum: 2,
           courses:['BIOL', 'CHEM'],
           details:  "Science help can also be found in:\n"
           + "NE Science Learning Center NHSC 1215\n"
           + "[Mon-Thu 8:30 AM - 9:00 PM]\n[Fri-Sat 8:30 AM - 5:00 PM]",
           status:false
          },
  mathPhys:{
            courseNum: 2,
            courses:['PHYS', 'MATH'],
            details: "Math help can also be found in:\n"
            + "NE Math & Physics Lab in NTAB 2204\n"
            + "[Mon-Thu 7:00 AM - 10:00 PM; Fri-Sat 7:00 AM - 5:00 PM]\n"
            + "Academic Learning Center in NACB 1114\n"
            + "[Mon-Thru 7:30 AM - 9:00 PM; Fri 7:30 AM - 6:30 PM; Sat 10:00 AM - 2:00 PM]",
            status:false
           },
  url:'https://www.tccd.edu/academics/academic-help/labs-tutoring/#ne'
};


/******************************* DO NOT EDIT THE DOCUEMENT BELOW THIS LINE *******************************/

// Functions:

// On form summation get the data from sheets and create a document
function onFormSubmit(e) {
    const leader = {
                    info:{
                        timeStamp:e.values[0],
                        email:e.values[1],
                        fName:e.values[2],
                        lName:e.values[3],
                        subj:e.values[4],
                        reason:e.values[9],
                        reasonD:e.values[10],
                        moreThanOne:e.values[11],
                        comments:e.values[16]
                    },
                    one:{
                        cDate:e.values[5],
                        sTime:e.values[6],
                        eTime:e.values[7],
                        rmNum:e.values[8]
                    },
                    multi:{
                        sDate:e.values[12],
                        eDate:e.values[13],
                        times:e.values[14],
                        rmNums:e.values[15]
                    }

                };
    var sendExtendedEmail = false;
    var docId;

    // Test
    Logger.log(leader);

    // Main Call
    try
    {
        if (leader.info.moreThanOne == 'Yes')
        {
            sendExtendedEmail = true;
            docId = createDocument(leader,sendExtendedEmail);
            sendExtendCancellation(leader, docId);
        }
        else
        {
            sendExtendedEmail = false;
            docId = createDocument(leader,sendExtendedEmail);
        }
    } catch(err) {
    Logger.log(err + "\n");
    }
}

// Create the Document in
function createDocument(leader, sendExtendedEmail)
{
    var docDate = getDate();
    var doc = DocumentApp.create(docDate + ' SI Session Cancelation For ' + leader.info.fName + " " + leader.info.lName);
    var docId = doc.getId();
    var docUrl = doc.getUrl();
    var header = doc.addHeader();
    var body = doc.getBody();

    // Header
    var headerStyle1 ={};
        headerStyle1[DocumentApp.Attribute.FONT_SIZE] = 12;
        headerStyle1[DocumentApp.Attribute.BOLD] = true;
    var headerText = header.getParent().getChild(1).asText().setText("\t\t\t\t\t\t\t\t\t\t\t" + leader.one.rmNum);
    headerText.setAttributes(headerStyle1);

    // Body Section 1: Leader's Session to Cancel
    var bodyStyle = {};
    bodyStyle[DocumentApp.Attribute.FONT_SIZE] = 30;
    bodyStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

    var bodyText = body.appendParagraph(leader.info.fName + " " + leader.info.lName + "\'s"
                                   + "\nSI Session for " + leader.info.subj
                                   + "\nat " + leader.one.sTime
                                   + "\nis CANCELED"
                                   + "\n" + leader.one.cDate).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    bodyText.setAttributes(bodyStyle);

    // Get User subject
    var subject = "";
    for(var i=0; i < 4; i++)
    {
        subject += leader.info.subj[i];
    }

    // Determine the Correct Learing Center
    for(var i=0; i < learningCenter.science.courseNum;i++)
    {
        if (subject == learningCenter.science.courses[i])
        {
          learningCenter.science.status = true;
        }
        else if (subject == learningCenter.mathPhys.courses[i])
        {
          learningCenter.mathPhys.status = true;
        }
    }

    // Body
    var bodyStyle2 = {};
    bodyStyle2[DocumentApp.Attribute.FONT_SIZE] = 18;
    bodyStyle2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

    // Footer
    var footerStyle1 = {};
        footerStyle1[DocumentApp.Attribute.FONT_SIZE] = 16;
        footerStyle1[DocumentApp.Attribute.BOLD] = true;
        footerStyle1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
    var footerStyle2 = {};
        footerStyle2[DocumentApp.Attribute.FONT_SIZE] = 12;
        footerStyle2[DocumentApp.Attribute.BOLD] = false;
        footerStyle2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;

    // Body Section 2: Learning Centers
    var bodyText2;
    var footerText1;
    if(learningCenter.science.status)
    {
        bodyText2 = body.appendParagraph( "\n\n" + learningCenter.science.details + "\n" + learningCenter.url + "\n\n\n\n");
    }
    else if (learningCenter.mathPhys.status)
    {
        bodyText2 = body.appendParagraph( "\n\n" + learningCenter.mathPhys.details + "\n" + learningCenter.url + "\n\n\n");
    }
    else
    {
        bodyText2 = body.appendParagraph( "\n\n" + learningCenter.url).setLinkUrl(learningCenter.url);
        body.appendParagraph("\n\n\n\n");
    }
    bodyText2.setAttributes(bodyStyle2);

    // Footer Section:
    footerText1 = body.appendParagraph("Sorry for any inconvenience.");
    footerText1.setAttributes(footerStyle2);
    footerText2 = body.appendParagraph("NE Supplemental Instruction Office\n" + officePhoneNumber);
    footerText2.setAttributes(footerStyle1);

    // Save and Close the document
    doc.saveAndClose();

    if(sendExtendedEmail)
    {
        Logger.log('Extended Email Sent');
        return docId;
    }
    else
    {
        var emailSubject = docDate + 'Cancellation Request for ' + leader.info.fName + " " + leader.info.lName;
        var emailBody =
                      "\nHello " + supervisorName + ",\n\n"
                      + leader.info.fName + ' ' + leader.info.lName + " has made a cancellation request.\n\n"
                      + "\nDate: " + leader.one.cDate
                      + "\nRoom Number: " + leader.one.rmNum
                      + "\nTime: " + leader.one.sTime + " to " + leader.one.eTime
                      + "\nReason: " + leader.info.reason
                      + "\nDetails: " + leader.info.reasonD
                      + "\nSI Leader Email: " + leader.info.email
                      + "\nGoogle Docs URL: " +  docUrl
                      + "\n\n";
        GmailApp.sendEmail(supervisorEmail, emailSubject, emailBody,
                           {attachments: [doc.getAs(MimeType.PDF)],name: 'Automated Cancellation Request',noReply:true });
  }

  return docId;
}

function sendExtendCancellation(leader, docId)
{
    var doc = DocumentApp.openById(docId);
    var docUrl = doc.getUrl();
    var subject = getDate() + ' Extend SI Session Cancelation For ' + leader.info.fName + " " + leader.info.lName;
    var body = "\nHello " + supervisorName + ",\n\n"
             + leader.info.fName + " " + leader.info.lName
             + " has requested and EXTENDED SESSION Cancellation for "
             + leader.info.subj + " course."
             + "\n\nDate Range: " + leader.multi.sDate + " to " + leader.multi.eDate + "."
             + "\nTimes: " + leader.multi.times
             + "\nRoom Numbers: " + leader.multi.rmNums
             + "\n\nUpcoming Session to cancel."
             + "\nUpcoming Date: " + leader.one.cDate
             + "\nTime: " + leader.one.sTime + " to " + leader.one.eTime
             + "\nRoom Number: " + leader.one.rmNum
             + "\nSI Leader Email: " + leader.info.email
             + "\nDoc URL: " + docUrl
             + "\nComments: " + leader.info.comments;

    doc.saveAndClose();
    GmailApp.sendEmail(supervisorEmail, subject, body,
                     {attachments: [doc.getAs(MimeType.PDF)],name:'Automated Cancellation Request',noReply:true });
}

function getDate()
{
    return Utilities.formatDate(new Date(), 'CST', 'MM-dd-yyyy - hh:mm:ss - ');
}
