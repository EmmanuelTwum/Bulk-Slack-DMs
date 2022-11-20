//Slack Messenger App

/*
    THIS SCRIPT ALLOWS YOUR NEWLY CREATED SLACK BOT TO SEND BULK MESSAGES ON YOUR BEHALF.
    @see https://benblaine.medium.com/send-personal-bulk-slack-dms-using-google-sheets-script-6fc44590aeb4
*/


//Slack App

const RECIPIENT_COL  = "User ID";
const DM_SENT_COL = "Email Sent";



var POST_MESSAGE_ENDPOINT = "https://slack.com/api/chat.postMessage";

// This code adds a menu item to the Google Sheet that you can use to send your message

function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Slack Messager Menu')
       .addItem('Send to Recipients', 'postLoop')
       .addToUi();
 }

// This code gets the list of user ID's from your Google Sheet and sends your Slack message to them one-by-one.

function postLoop () {
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input")
 //var rangeValues = sheet.getDataRange().getValues();
 
// get the draft Gmail message to use as a template
  const msgTemplate = 'Hello *{{Name}}*, TYPE THE REST OF YOUR MESSAGE HERE';

  // get the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetch displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // assuming row 1 contains our column headings
  const heads = data.shift(); 
  
  // get the index of column named 'Email Status' (Assume header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const DMSentColIdx = heads.indexOf(DM_SENT_COL);
  
  // convert 2d array into object array
  // @see https://stackoverflow.com/a/22917499/1027723
  // for pretty version see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  
  // used to record sent emails
  const out = [];
 /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return  JSON.parse(template_string);
  }

   /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
  */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };


  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    if (row[DM_SENT_COL] == ''){
      try {
        const msgObj = fillInTemplateFromObject_(msgTemplate, row); 
        var channel = row[RECIPIENT_COL];
    postToSlack(channel)
 function postToSlack(channel) {
  var payload = {
    'channel' : channel,
    "type": "mrkdwn",
    'text' : msgObj,
    'as_user' : true
  }
  
return UrlFetchApp.fetch(
  POST_MESSAGE_ENDPOINT,
  {
    method             : 'post',
    contentType        : 'application/json',
    headers            : {
      Authorization : 'Bearer ' + 'xoxp-1262639085254-2000788509024-4330188717364-b2cb8c702f9d2edde83f94a99a8ea2dc'
    },
    payload            : JSON.stringify(payload),
    muteHttpExceptions : true,
})
}

// modify cell to record email sent date
      out.push([new Date()]);
    } catch(e) {
      // modify cell to record error
      out.push([e.message]);
    }
  } else {
    out.push([row[DM_SENT_COL]]);
  }
});
  
// updating the sheet with new data
sheet.getRange(2, DMSentColIdx+1, out.length).setValues(out);

// This is the code that sends your message to Slack, it is called by the above function postLoop()



}
