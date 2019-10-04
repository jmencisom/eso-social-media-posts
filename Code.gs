/*
 * Script to export social media posts to Orlo format
 * Author:  Javier Enciso
 * Email:   j4r.e4o@gmail.com
 * Version: 2019.10.04 04:09 COT
*/

//////////////////////////////////////////////////////////////////////

/**
 * Check if today is Central European Summer Time - CEST. CEST begins on the last Sunday in March and 
 * ends on the last Sunday in October each year.
 * @return {boolean} True is current date is between the last Sunday in March and the last Sunday in October each year.
 * @see https://en.wikipedia.org/wiki/Summer_Time_in_Europe
 * @see https://gist.github.com/danalloway/17b48fddab9028432c68
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function is_today_cest(){

  var currentDate = new Date();
  var currentYear = currentDate.getFullYear();
  
  // CEST Start
  var firstOfMarch = new Date(currentYear, 2, 1);
  var daysUntilFirstSundayInMarch = (7 - firstOfMarch.getDay()) % 7;
  var lastSundayInMarch = firstOfMarch.getDate() + daysUntilFirstSundayInMarch + 21;
  var cestStartDate = new Date(currentYear, 2, lastSundayInMarch);
  
  // CEST End
  var firstOfOctober = new Date(currentYear, 9, 1);
  var daysUntilFirstSundayInOctober = (7 - firstOfOctober.getDay()) % 7;
  var lastSundayInOctober = firstOfOctober.getDate() + daysUntilFirstSundayInOctober + 21;
  var cestEndDate = new Date(currentYear, 9, lastSundayInOctober);
  
  // Logs
  Logger.log("inputDate: " + currentDate);
  Logger.log("cestStartDate: " + cestStartDate);
  Logger.log("cestEndDate: " + cestEndDate);
  
  if (currentDate > cestStartDate && currentDate < cestEndDate){
    return true;
  }
  else{
    return false;
  }
}

var GERMAN_DST = "";

if (is_today_cest()){
  GERMAN_DST = "GMT+2";
}
else{
  GERMAN_DST = "GMT+1";
}

Logger.log("GERMAN_DST: " + GERMAN_DST);


//////////////////////////////////////////////////////////////////////

// Email address allowed to execute "Export Posts" and "Move to Done" menu items
var allowed_emails = [];
allowed_emails.push("marialejarojas@gmail.com");
allowed_emails.push("j4r.e4o@gmail.com");
allowed_emails.push("garamora@gmail.com");
allowed_emails.push("raquel.shida@gmail.com");
allowed_emails.push("sandu.oana@gmail.com");
allowed_emails.push("oana.bescience@gmail.com");
allowed_emails.push("oanasandu.eso@gmail.com");

// Messages and name of the tabs
var ACCESS_DENIED_MSG = 'Access denied. Please contact epodweb@eso.org to request access.';
var SM_POSTS_TAB_NAME = 'SM Posts';
var DONE_TAB_NAME = 'Done';

// Lenght tweet with image
var LENGHT_TWEET_WITH_IMG = 240;



/**
 * Add menu items to Google Spreadsheets
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu_entries = [];
  
  var menu_label_eso = "ESO Tools";
  var menu_label_export = "Export Posts to Orlo";
  var menu_label_move = "Move exported Posts to Done tab";
  
  if (is_today_cest()){
    menu_label_export += " (CEST)";
  }
  else{
    menu_label_export += " (CET)";
  }
  
  menu_entries.push({name: menu_label_export, functionName: "saveAsSSProxy"});
  menu_entries.push(null);
  menu_entries.push({name: menu_label_move, functionName: "moveToDoneProxy"});
  ss.addMenu(menu_label_eso, menu_entries);
};


/**
 * Enable to call moveToDone() method to allowed users
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function moveToDoneProxy(){
  
  var email = Session.getActiveUser().getEmail();   
  if (allowed_emails.indexOf(email) != -1){
    moveToDone()
  }
  else{
    Browser.msgBox(ACCESS_DENIED_MSG);
  }
}


/**
 * Enable to call saveAsSS() method to allowed users
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function saveAsSSProxy(){
  
  var email = Session.getActiveUser().getEmail();   
  if (allowed_emails.indexOf(email) != -1){
    saveAsSS()
  }
  else{
    Browser.msgBox(ACCESS_DENIED_MSG);
  }
}


/**
 * Move exported posts to Done tab
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function moveToDone() {
  
  // Current sheet
  var posts_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SM_POSTS_TAB_NAME);

  // Get Done sheet
  var done_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DONE_TAB_NAME);
  
  // Parameter for the Original Spreadsheet
  // Where is the Export indicator
  // Indexes start at 0
  var COL_EXPORT = 8;
  var ROW_FIRST = 4;

  
  var rows_to_export = [];
  
  // get available data range in the spreadsheet
  var activeRange = posts_sheet.getDataRange();
  try {
    var data = activeRange.getValues();    

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {

      // Preprocessing. Counting rows to export
      for (var row = ROW_FIRST; row < data.length; row++){ 
        if (data[row][COL_EXPORT] == "YES"){
          rows_to_export.push(row + 1);
          // This inserts a row below the column headers
          done_sheet.insertRowBefore(2);
          
          var source = posts_sheet.getRange(row + 1, 1, 1, posts_sheet.getMaxColumns());
          var destination = done_sheet.getRange(2, 1, 1, done_sheet.getMaxColumns());
          
          source.copyTo(destination);
          
        }
      }
      
      // Delete rows from SM Posts Tab
      for (var row = rows_to_export.length - 1; row >= 0 ; row--){
        posts_sheet.deleteRow(rows_to_export[row]);
      }
      
      Browser.msgBox('Moved ' + rows_to_export.length + ' posts');
    }
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

/**
 * Create Spreadsheet from selected posts
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function saveAsSS() {
  // Current Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SM_POSTS_TAB_NAME);
  
  // Current sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SM_POSTS_TAB_NAME);
  
  var ui = SpreadsheetApp.getUi();
  var RESPONSE_TITLE = "Timezone Difference";
  var RESPONSE_MESSAGE = "What's your timezone difference respect to German time?";
  RESPONSE_MESSAGE += "\nExamples: https://goo.gl/kP8tqS";
  var response = ui.prompt(RESPONSE_TITLE, RESPONSE_MESSAGE, ui.ButtonSet.OK_CANCEL);
  var user_time_offset = 0;

  try {
    if (response.getSelectedButton() == ui.Button.OK) {
      user_time_offset = Number(response.getResponseText());
      
      // Name of the output Spreadsheet without spaces
      var ss_name = ss.getName().toLowerCase().replace(/ /g,'_')
      
      // Create new Spreadsheet for output
      var newdate = new Date();
      var ss_output_name = ss_name + "_output_" + Utilities.formatDate(newdate, GERMAN_DST, "yyyy-MM-dd_HH:mm:ss");
      var ss_output = SpreadsheetApp.create(ss_output_name);
      
      // Add Sheet for output
      var ss_output_sheet = ss_output.getActiveSheet();
      
      // Create a file in the Docs List with the given name and the csv data
      var csv = convertRangeToXlsFile_(sheet, user_time_offset);
      
      for (var i = 0; i < csv.length; i++) {
        ss_output_sheet.appendRow(csv[i]);
      }
      
      Browser.msgBox('Output file created ' + ss_output_name);    
    }
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
  
}


/**
 * Remove special characters in post text
 * @param {String} text with special characters
 * @return {String} text without special characters
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function preprocess_text(text) {
  // Remove conflicting caracters from text
  var t = text;
  
  // Remove carreige returns and spaces from the end
  t =  t.toString().trim();
  
  // Replace inner carreige returns by single space
  //t =  t.replace(/\n/g, " ");
  //t =  t.replace(/\r/g, " ");
  
  // Straigh quotation marks
  t =  t.replace("'", "'");
  t =  t.replace("’", "'");
  t =  t.replace('"', '"');
  t =  t.replace('"', '"');
  t =  t.replace('"', '"');
  t =  t.replace('-', '-');
  
  return t;
  
}


/**
 * Determine if an account id belongs to a twitter account
 * @param {Number} account id
 * @return {Boolean} True if the account id belongs to a twitter account
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function is_twitter(account_id) {
  // Orlo account id is available in Bulk Upload Sample File
  var accounts = [6836, 7020, 7025, 7022, 7000, 6999, 7007, 7002, 7018, 7023, 7013, 7003, 7004, 7005, 7006, 7002, 7018, 7024, 7018, 7010, 7002, 7027, 7015, 7017, 7008, 7011, 7012, 7021, 7014, 7016, 7019];
  accounts.push(6879); // Hubble TW
  accounts.push(6881); // ESO TW Supernova
  accounts.push(6882); // ESO TW Supernova DE
  accounts.push(6880); // IAU TW
  
  if (accounts.indexOf(account_id) != -1){
    return true;
  }
  return false;
}


/**
 * Determine if an account id belongs to an instagram account
 * @param {Number} account id
 * @return {Boolean} True if the account id belongs to an instagram account
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function is_instagram(account_id) {
  // Orlo accounts id is available in Bulk Upload Sample File
  var accounts = [];
  accounts.push(17535); // Hubble Instagram
  accounts.push(17536); // ESO Instagram
  
  if (accounts.indexOf(account_id) != -1){
    return true;
  }
  return false;
}

/**
 * Determine if the language of an account is Spanish
 * @param {Number} column index
 * @return {Boolean} True if the language of column index is Spanish
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function is_spanish_col(column) {
  // Column index starts at 0
  var accounts = [];
  accounts.push(35); // ESO FB Chile
  accounts.push(36); // ESO FB Spain
  accounts.push(68); // ESO TW Chile
  accounts.push(69); // ESO TW Spain
  
  if (accounts.indexOf(column) != -1){
    return true;
  }
  return false;
}


/**
 * Determine if the language of an account is German
 * @param {Number} column index
 * @return {Boolean} True if the language of column index is German
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function is_german_col(column) {
  // Column index starts at 0
  var accounts = [];
  accounts.push(18); // ESO FB Supernova DE
  accounts.push(19); // ESO FB Austria
  accounts.push(20); // ESO FB Germany
  accounts.push(21); // ESO FB Belgium-de
  accounts.push(22); // ESO FB Swizerland-de
  accounts.push(51); // ESO TW Supernova DE
  accounts.push(52); // ESO TW Austria
  accounts.push(53); // ESO TW Germany
  accounts.push(54); // ESO TW Belgium-de
  accounts.push(55); // ESO TW Swizerland-de  
  
  if (accounts.indexOf(column) != -1){
    return true;
  }
  return false;
}


/**
 * Customize link based on country for eso, supernova and alma sites
 * @param {String} Link to customize
 * @param {String} Country name
 * @param {Number} column index
 * @return {String} Customized link
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function customize_link_col(link, country, column) {
  // Remove carreige returns and spaces from the end
  l = link.toString().trim();
  var alma_domain = "kids.alma.cl";
  var supernova_domain = "supernova.eso.org";
  var eso_domain = "eso.org/public";
  
  
  // Custom Supernova links
  if (l.indexOf(supernova_domain) != -1 ){
    // Link belongs to supernova domain
    if (is_german_col(column)){
      // Risky: Check in case the alma url format gets updated!
      l = l.replace("supernova.eso.org/","supernova.eso.org/germany/");
      l = l + "?lang"; 
    }
    return l;
  }
  
  // Custom ALMA links
  if (l.indexOf(alma_domain) != -1 ){
    // Link belongs to alma domain
    if (is_spanish_col(column)){
      // Risky: Check in case the alma url format gets updated!
      l = l.replace("lang=en","lang=es");
    }
    return l;
  }
  
  // Custom ESO Links
  if (l.indexOf(eso_domain) != -1 ){
    if (country){
      l = l.replace("eso.org/public/","eso.org/public/" + country + "/");
      l = l + "?lang";    
    }
    return l;
  }
  
  // Default case
  return l;
}


/**
 * Check if a text is a number
 * @param {String} text
 * @return {Boolean} True if the text is number
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function isNumeric(text) {
  return !isNaN(parseFloat(text)) && isFinite(text);
}


/**
 * Create XLS file from selected rows
 * @param {String} Spreadsheet id
 * @param {Number} Time offset of the user who exports the posts.
 * @return {String} String with the CSV values to export
 * @author Javier Enciso <jenciso@partner.eso.org>
 */
function convertRangeToXlsFile_(sheet, user_time_offset) {
  
  // Index of the export column
  var COL_EXPORT = 8;
  
  // Index of the live-local-now column
  var COL_LIVELOCAL = 7;

  // Ignored columns, 
  var COL_IGNORED_INDEX = [];
  // i.e., ESO Instagram column - Index start from 0
  // COL_IGNORED_INDEX.push(13);
  // i.e., Hubble Instagram column - Index start from 0
  // COL_IGNORED_INDEX.push(81);
  // No need to ignore columns since Social Sign In support Instagram too
  var COL_IGNORED_COUNT = COL_IGNORED_INDEX.length;

  // Index of the first column to export
  // Column A = 0, B = 1,..., L = 11
  var COL_FIRST = 11;
  
  var ROW_ACCOUNT_ID = 0;
  var ROW_COUNTRY_URL = 1;
  var ROW_COUNTRY_OFFSET = 2;
  var ROW_FIRST = 4;
  
  var rows_to_export = [];
  
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();    

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {

      // Preprocessing. Counting rows to export
      for (var row = ROW_FIRST; row < data.length; row++){ 
        if (data[row][COL_EXPORT] == "YES"){
          rows_to_export.push(row); 
        }
      }

      // Calculating the size of the CSV Array
      var post_count = (data[0].length - COL_FIRST - COL_IGNORED_COUNT) * rows_to_export.length;
      var csv = new Array(post_count);
      for (var i = 0; i < post_count; i++) {
        csv[i] = new Array(4); // 4 columns to export
      }


      // Processing
      var i = 0; // export row counter
      
      for (var row = 0; row < rows_to_export.length; row++){

        // Current row to export
        var r = rows_to_export[row];

        
        // Live/Local/Now indicator
        // Live: time_offset = user_time_offset + local_time_offset, for local_time_offset >= 0
        // time_offset = 0 otherwise
        // Local: time_offset = user_time_offset + local_time_offset
        // Now: time_offset = user_time_offset
        var livelocal_sel = data[r][COL_LIVELOCAL];
        
        for (var col = COL_FIRST; col < data[0].length; col++) {
          // Moving accross the available columns
          
          if (COL_IGNORED_INDEX.indexOf(col) == -1) {
            // Do only for non-ingnored columns
          
            var text = "";
            var link = "";
            var img_url = "";
            var account_id = "";
            
            
            
            text = preprocess_text(data[r][col]);
            
            if (text){
              // Add country to the link
              var country = data[ROW_COUNTRY_URL][col];
              // Account id
              account_id = data[ROW_ACCOUNT_ID][col];
              csv[i][3] = account_id;
              
              var post = "";
              
              if (is_instagram(account_id)){
                // Instagram posts don't take the Link
                post = text;
              }
              else{
                // Cusotmize link based on language for Facebook and Twitter accounts
                link = customize_link_col(data[r][2], country, col);
                // Text + Link
                post = text + " " + link;
              }
              
              
              // Browser.msgBox(post);
              csv[i][0] = post;
              
              // Image URL
              // Include links in twitter posts only with length is less than LENGHT_TWEET_WITH_IMG
              if (is_twitter(account_id)){
                if (text.length <= LENGHT_TWEET_WITH_IMG){
                  // Remove endlines
                  img_url = data[r][3].toString().trim();
                }
                else{
                  img_url = "";
                }
              }
              else{
                // Remove endlines
                img_url = data[r][3].toString().trim();
              }
              
              csv[i][1] = img_url;
              
              
              // Publication date
              var pub_date = new Date(data[r][4]);
              var local_time_offset = data[ROW_COUNTRY_OFFSET][col];
              
              if (isNumeric(local_time_offset)){
                //pass
              }
              else{
                // set to 0 when there is no numeric input in row 3
                local_time_offset = 0;
              }
              
              var time_offset = 0;
              switch (livelocal_sel) {
                case "LIVE":
                  if (local_time_offset >= 0){
                    time_offset = user_time_offset + local_time_offset;
                  }
                  else{
                    time_offset = user_time_offset;
                  }
                  break;
                case "LOCAL":
                  time_offset = user_time_offset + local_time_offset;
                  break;
                case "NOW":
                  time_offset = user_time_offset;
                  break;
              }
              
              pub_date.setHours(pub_date.getHours() + time_offset);
              
              //  Date-time to export
              var pub_date_format = Utilities.formatDate(pub_date, GERMAN_DST, "yyyy/MM/dd HH:mm:ss");
              
              // Browser.msgBox(pub_date_format);
              csv[i][2] = pub_date_format;
              
              // Counter of records to export
              i++;
            }
          }
        } // End if ignore columns
      } // Enod for
      // Remove empty rows
      var post_count_real = i;
      var ncsv = new Array(post_count_real);
      for (var j = 0; j < post_count_real; j++) {
        ncsv[j] = csv[j];
      }
      
    }
    return ncsv;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
