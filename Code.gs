// original from: http://mashe.hawksey.info/2014/07/google-sheets-as-a-database-insert-with-apps-script-using-postget-methods-with-ajax-example/
// original gist: https://gist.github.com/willpatera/ee41ae374d3c9839c2d6 

function doGet(e){
  return handleResponse(e);
}

//  Enter sheet name where data is to be written below
var SHEET_NAME = "Sheet1";
var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

function handleResponse(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  //var lock = LockService.getPublicLock();
  //lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    //var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var doc = SpreadsheetApp.openById("16mihBcC5HokcYu8vC4ayLV-OMkf3F2r7JzD5r-M97tY");
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row = []; 
    // loop through the header columns
   for (var i = 0; i < headers.length; i++) { // start at 1 to avoid Timestamp column
            if (headers[i].length > 0) {
                row.push(e.parameter[headers[i]]); // add data to row
            }
        }
     var sname = String(row[0]);
            var sevis = String(row[1]);
            var cwid = String(row[2]);
            var email = String(row[3]);
            var visatype = String(row[4]);
            var classification = String(row[5]);
            var major = String(row[6]);
            var departure = String(row[7]);
            var rturn = String(row[8]);
            var coc = String(row[9]);
            var semesterenrolled = String(row[10]);
            var hoursenrolled = String(row[11]);
            var visa = String(row[12]);
            var bursar = String(row[13]);
            var request = String(row[14]);
			var comments = String(row[15]);
            var change_add = String(row[16]);
            var update_add = String(row[17]);
            var banner_major = String(row[18]);
            var review = String(row[19]);
            var grad = String(row[20]);
            var enrnext = String(row[21]);
            var enrhr = String(row[22]);
            var assis = String(row[23]);
            var enrdate = String(row[24]);
            var spa = String(row[25]);
            var prior = String(row[26]);
            var fn = String(row[27]);
            var ln = String(row[28]);
            var dateslot = String(row[29]);
            var timeslot = String(row[30]);

                    var name = cwid + " Travel Signature";
                    traveltemp = DriveApp.getFileById('1ICjTiNE3iiik5VYwBSNY4assRmnkNA5x69Zn3VlqjJU').makeCopy(name);
                    var id = traveltemp.getId(); 
                    var worddoc = DocumentApp.openById(id);
                    var contents = worddoc.getBody();

                    contents.replaceText("<fname>", fn);
                    contents.replaceText("<lname>", ln);
					contents.replaceText("<major>", major);
                    contents.replaceText("<sevis>", sevis);
                    contents.replaceText("<cwid>", cwid);
                    contents.replaceText("<visatype>", visatype);
                    contents.replaceText("<email>", email);
                    contents.replaceText("<classification>", classification);
					contents.replaceText("<departure>", departure);
					contents.replaceText("<ret>", rturn);
					contents.replaceText("<coc>", coc);
					contents.replaceText("<semesterenrolled>", semesterenrolled);
                    contents.replaceText("<hoursenrolled>", hoursenrolled);
					contents.replaceText("<bursar>", bursar);
					contents.replaceText("<visa>", visa);
                    contents.replaceText("<request>", request);
                    contents.replaceText("<change_add>", change_add);
                    contents.replaceText("<update_add>", update_add);
                    contents.replaceText("<banner_major>", banner_major);
                    contents.replaceText("<review>", review);
                    contents.replaceText("<grad>", grad);
                    contents.replaceText("<enrnext>", enrnext);
                    contents.replaceText("<prior>", prior);
                   
                    
                if (enrhr == "") {
                   contents.replaceText("<enrhr>", "N/A");

                }
                else {
                
                contents.replaceText("<enrhr>", enrhr);
                
                }
                    if (assis == "") {
                  
                    contents.replaceText("<assis>", "N/A");
                    
                    }
                    
                   else {
                   
                    contents.replaceText("<assis>", assis);
                   
                   }
                   
                   if (enrdate == "") {
                  
                    contents.replaceText("<enrdate>", "N/A");
                    
                    }
                    
                   else {
                   
                    contents.replaceText("<enrdate>", enrdate);
                   
                   }
                    
                    contents.replaceText("<spa>", spa);
                    
					
                    worddoc.saveAndClose();
                     row.push(traveltemp.getUrl());
               
              
               
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    //Prepare and send emails
        
                    var subjectLine = "Travel Signature Request for " + sname ;
                    var body = "<HTML><BODY>" +
"<p style=font-size:14px;>Dear  " + sname + ",</p> <p> </p> <p style=font-size:14px;><strong><span style=color:#000000;>" + 
"Your request for the Travel Singature has been approved</span></strong></p><p></p><p><strong><u>Next steps</u></strong>" + 
" for obtaining your DS-2019 or I-20 travel signature.</p><p></p><p><span style=color:#ff0000;><strong><u>You will need the following documents:" +
"</u></strong></span><br /><br /><span style=color:#ff0000;>1. Print and sign the attached Students Travel Signature form " +
"</span><br /><span style=color:#ff0000;>2. Your U.S Visa</span><br /><span style=color:#ff0000;>3. Your Passport</span><br />" +
"<span style=color:#ff0000;>4. Your most recent I-20</span><br /><span style=color:#ff0000;>" +
"5. Take all the Documents above (1,2,3 and 4) to the ISS Office for signature</span></p><p></p><p><span style=color:#ff0000;><strong>" +
"<u>Your ISS office appointment is on-<br />" + dateslot + " at " + timeslot + "</u></strong></span></p><p></p>" +
"<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>Useful Information: <br />What documents are needed for travel:</p><p></p>" +
"<p><strong>Take the following documents for travel out the U.S.<br />&bull; Valid passport, Visa, and I-94<br />&bull; " + 
"I-20 or DS-2019 signed for travel by an ISS staff member. <br />&bull; Official transcript from (Registrar&rsquo;s Office &ndash; " +
"transcript section, 322 Student Union<br />&bull; Letter of Enrollment (Certificate of Enrollment) (Registrar&rsquo;s Office &ndash; " +
"Certification section, 322 Student Union).<br />&bull; SEVIS Fee Receipt from your initial entry (first time travelers)<br />" +
"If you are renewing your visa, you will need proof of financial documentation for the visa-issuing officer.<br />" +
"The financial documentation could include the following - a personal bank statement, Research or Teaching Assistantship letter that " +
"includes your salary, tuition waivers, health insurance, or sponsor&rsquo;s letter.<br />You could also have a combination of documents. " +
"<br /><br /> Please check the CDC International Travel page (https://www.cdc.gov/coronavirus/2019-ncov/travelers/international-travel-during-covid19.html) for more information about traveling during pandemic" + 
"<br /><br />If you have any further questions, please let us know.<br /><br />Best regards,<br /><br />Office of International Students &amp; Scholars " +
"<br />309 WWC | Oklahoma State University<br />Phone: (405) 744-5459 | Fax: (405) 744-8120<br />Email: iss@okstate.edu<br />Homepage: http://iss.okstate.edu/</strong></p></BODY></HTML>";
                    MailApp.sendEmail({
                        to: email,
                        subject: subjectLine,
                        htmlBody: body,

                        name: "Travel Signature",
                        attachments: [traveltemp.getAs(MimeType.PDF)] //[traveltemp.getAs(MimeType.PDF)]
                    });
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    //lock.releaseLock();
  }
}