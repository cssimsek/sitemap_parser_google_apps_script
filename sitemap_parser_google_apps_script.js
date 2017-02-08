/* *************** */
/* FETCH SITEMAP URLS  
/* *************** */   
function fetchSitemapURLs(){
  
   //Hardcoded sheet index - 0 indexed
   var sitemapURLsSheetIndex = 0;
  
   //book and sheet references
   var sheetOfUrls = SpreadsheetApp.getActiveSpreadsheet();
   var sheetName = sheetOfUrls.getSheets()[sitemapURLsSheetIndex].getSheetName();
   var theSelectedSheet = sheetOfUrls.getSheetByName(sheetName);
   var operationalRange = theSelectedSheet.getDataRange();
   var dataFromRange = operationalRange.getValues();
   var lengthOfRange = operationalRange.getNumRows();
   Logger.log(lengthOfRange);
  
   //Hardcoded data range definitions
   var columnIndex = 4; // the index of the column that contains the sitemap URLs - 1 indexed
   var rangeStartRow = 1; // 1 indexed
   var rowDelimit = 2; // assign delimiting row number if not all sitemaps are to be fetched - 1 indexed
   var untilRow = rowDelimit || lengthOfRange;
  
   //XML Sitemap Namespace
   var xmlNS_String = "http://www.google.com/schemas/sitemap/0.9"; //
  
   
   //deleteTriggers if necessary
   deleteTriggers();
   function deleteTriggers(){
     if(ScriptApp.getProjectTriggers()){
       var triggers = ScriptApp.getProjectTriggers();
       for (var j = 0; j < triggers.length; j++) {
         ScriptApp.deleteTrigger(triggers[j]);
       }
     }
   }

   
  var urlFetchOptions = {
      "muteHttpExceptions": true,      
      "followRedirects": true
  };
    
  
   var theResponses = [];
  
    //timeKeeper
   var startTime = (new Date).getTime(), killTime = startTime + (24*(Math.pow(10,4)));
   
  //define UrlFetchApp options object
   var columnHeaders;
     
   //callRedirectChecker();
   sitemapParser:
   for(var i=(rangeStartRow-1);i<untilRow;i+=1){
      if((new Date).getTime()<killTime){
        var url2Check = encodeURI(decodeURI((dataFromRange[i][columnIndex-1]).trim()));
        Logger.log(url2Check);
        var ok = false;
        do{
          Logger.log(theResponses);
          try {
            var theResponse = UrlFetchApp.fetch(url2Check,urlFetchOptions);
            if(theResponse.getHeaders()["Content-Type"] &amp;&amp; theResponse.getHeaders()["Content-Type"].indexOf("xml")!=-1 &amp;&amp; (theResponse.getResponseCode()=="200")){
              var xmlFromResponse = theResponse.getContentText();
              var xmlDocument = XmlService.parse(xmlFromResponse);
              var xmlRootEl = xmlDocument.getRootElement();
              var xmlProtocol = XmlService.getNamespace(xmlNS_String);
              var urlEntries = xmlDocument.getRootElement().getChildren('url', xmlProtocol);
              //Logger.log("urlEntries.length:" + urlEntries.length);
              var urlsArray = [];
              for (var urlIndex= 0; urlIndex<urlEntries.length;urlIndex++) {
                urlsArray.push([url2Check,urlEntries[urlIndex].getChild('loc', xmlProtocol).getText(),(new Date).toUTCString()]);
              }
            }else{
              continue sitemapParser;
            }
            
            //Logger.log(urlsArr + "\n");
            theResponses.push(urlsArray);
            Utilities.sleep(50);
            ok = true;
        }catch(e) {
          Logger.log(e);
          if((e+"").indexOf("DNS error") != -1){
            theResponses.push(new Array(e+"",undefined));
            continue sitemapParser;
          }
          Utilities.sleep(2000);
        } 
      }while(!ok);
    }
    else{
      break; 
    }
   }
  
  Logger.log("theResponses.length: \n" + theResponses.length);
  
  
  
  Logger.log("theResponses[0][1][1]: \n" + theResponses[0][1][1].match(/https?:\/\/.*?\//ig)[0]);
  
  
  var newSheetsObj = {};
  //New Sheet forEach Sitemap
  for(var sitemap=0;sitemap<theResponses.length;sitemap+=1){
     var sheetOfUrls = SpreadsheetApp.getActiveSpreadsheet();
     var theURL = theResponses[sitemap][1][1].match(/https?:\/\/.*?\//ig)[0];
     sheetOfUrls.insertSheet(theURL,(sheetOfUrls.getNumSheets()+1));
     var currentSheet = sheetOfUrls.getSheetByName(theURL+"");
     currentSheet.activate();
     Utilities.sleep(500);
     columnHeaders = [["Sitemap URL","URLs in Sitemap", "Timestamp"]];
     
     for(var headerIndex=0;headerIndex<columnHeaders[0].length;headerIndex+=1){
        columnHeaders[0][headerIndex];
     }
    
     var headerRange = currentSheet.getRange(("A"+1+":"+"C"+columnHeaders.length));
     headerRange.setValues(columnHeaders);
    
     var range = currentSheet.getRange("A"+2+":"+"C"+((theResponses[sitemap].length)+1)); 
     range.setValues(theResponses[sitemap]);
  }
  
  //Column Title Defs
    Logger.log("rangeStartRow: " + rangeStartRow + " &amp;&amp; theResponses.length:" + theResponses.length + " &amp;&amp; lengthOfRange: " + lengthOfRange);
  
  /* Generic timeBased trigger */
  function prepTrigger(mins,functionName){
        mins = mins > 11 ? mins : 11;
        ScriptApp.newTrigger(functionName).timeBased().after(mins * 60 * 1000).create();
  }
  
  if(((rangeStartRow+theResponses.length)-1)<untilRow ){
       prepTrigger(12,"fetchSitemapURLs");
  }else{
  Logger.log("fetchSitemapURLs completed work on: " + ((rangeStartRow+theResponses.length)-1) + " rows"); 
  }
}

/* *************** */
/* CLOSE FETCH SITEMAP URLS */
/* *************** */ 