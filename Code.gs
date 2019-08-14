// Add-on logic
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createAddonMenu()
    .addItem("Begin Metadata Collection", "openSidebar")
    .addItem("Enter API key", "setKey")
    .addItem("Enter retmax", "setRetmax")
    .addToUi();
}

function onInstall() {
  onOpen();
}

function openSidebar() {
  var html = HtmlService
      .createTemplateFromFile('sidebar')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('SRA Metadata Collector');
  SpreadsheetApp.getUi().showSidebar(html);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// Parse XML to generate metadatatable

function processXMLFull(SRA_UIDs,header,save_pref){
   for(var key in SRA_UIDs){
     SRA_UIDs[key] = XmlService.parse(SRA_UIDs[key]);
   }
  var metadata = processXML_(SRA_UIDs);
  outputTable_(metadata,header,save_pref);
  return 'Metadata gathered.';
}

// Parse XML
function processXML_(sraXmlD) {
      // Receives an XML object obtained from SRA.
      // Returns a LoL object corresponding to a table of metadata.
  var metadata = {};
  for(var key in sraXmlD){
      var entries = sraXmlD[key].getRootElement().getChildren();
      var currentMetadata=[]; // LoL where each primary list corresponds with a run to be output.
      for(var j = 0; j < entries.length; j++) {
        // Works by extracting data to a common variable, currentElement, which is then appended to the currentData, which is appended to the LoL table.
        var currentData = [key];
        var currentElement = entries[j].getChild('SAMPLE');
        currentData.push(currentElement.getChild('IDENTIFIERS').getChildText('PRIMARY_ID')); // Sample Accession
        
        var currentElement = entries[j].getChild('EXPERIMENT');
        currentData.push(currentElement.getChildText('TITLE')); // Sample title
        
        var currentElement = entries[j].getChild('SAMPLE'); // Fields from SAMPLE
        // Gathers sample attributes or gathers description if absent.
        if(currentElement.getChild('SAMPLE_ATTRIBUTES')!=null){
          currentData.push(currentElement.getChild('SAMPLE_ATTRIBUTES')
                         .getChildren()
                         .map(function(x) {return x.getChildText('TAG')+': '+x.getChildText('VALUE')})
                         .join(" || ")); // Creates a string of sample attributes.
        } else if (currentElement.getChildText('DESCRIPTION')!=null){
           currentData.push(currentElement.getChildText('DESCRIPTION'));
        } else {currentData.push('')};
        
        var currentElement = entries[j].getChild('EXPERIMENT') // Fields from EXPERIMENT
        currentData.push(currentElement.getChild('PLATFORM').getChildren()[0].getChildText('INSTRUMENT_MODEL')); // Get platform
        var currentElement = entries[j].getChild('STUDY').getChild('IDENTIFIERS').getChildren();
        currentData.push(currentElement.map(function(x) {return x.getValue();}).join());
        
        var currentElement = entries[j].getChild('EXPERIMENT').getChild('DESIGN').getChild('LIBRARY_DESCRIPTOR');
        currentData.push(currentElement.getChildText('LIBRARY_STRATEGY'));
        currentData.push(currentElement.getChild('LIBRARY_LAYOUT').getChildren()[0].getName());
        currentData.push(currentElement.getChildText('LIBRARY_SOURCE'));
        currentData.push(currentElement.getChildText('LIBRARY_SELECTION'));
        // Construction design can be placed in different places sometimes.
        var design = currentElement.getChildText('LIBRARY_CONSTRUCTION_PROTOCOL');
        if(design==null){
          currentData.push(entries[j].getChild('EXPERIMENT').getChild('DESIGN').getChildText('DESIGN_DESCRIPTION'));
        } else {
          currentData.push(design);
        }
        // Getting run information, which may have a multiple mapping to the rest of the metadata.
        var currentData = entries[j]
           .getChild('RUN_SET')
           .getChildren()
           .map(function(x) {return [x.getChild('IDENTIFIERS').getChildText('PRIMARY_ID')].concat(currentData)
                                                                                          .concat([(+x.getAttribute('total_spots').getValue()/1000000).toFixed(1).toString()+'M'])});
        currentMetadata = currentMetadata.concat(currentData);
        // Adds parsed metadata to the dictionary, overwriting the XML.
      }
        metadata[key] = currentMetadata;
  }
  return metadata;
}
// Display metadata table
function outputTable_(metadata,header,save_pref_param) {
  Logger.log("Began output");
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  if(save_pref_param){
    for(var key in metadata){
    var cSheet = sheet.getSheetByName(key);
    if (cSheet != null) {
      // Fails if this is the last sheet in the spreadsheet. Simply fail out if so.
      try {
        sheet.deleteSheet(cSheet);
      } catch(e) {}
    } 
    cSheet = sheet.insertSheet();
    cSheet.setName(key);
    // Plus 1 since you need to add a header.
    metadata[key].unshift(header);
    var range = cSheet.getRange(1,1,metadata[key].length,metadata[key][0].length);
    range.setValues(metadata[key]);
    }
  } else {
  var cSheet = sheet.getSheetByName('Results');

    if (cSheet != null) {
      // Fails if this is the last sheet in the spreadsheet. Simply fail out if so.
      try {
        sheet.deleteSheet(cSheet);
      } catch(e) {}
    } 
    cSheet = sheet.insertSheet();
    cSheet.setName('Results');
  // Merging all metadata entries
  var mMetadata = [header];
  for (var key in metadata) {
  mMetadata = mMetadata.concat(metadata[key])
  }
  var range = cSheet.getRange(1, 1, mMetadata.length, mMetadata[0].length);
  range.setValues(mMetadata);
  }
}




// API-key get set
function setKey() {
  key=showAPIKEYPrompt_();
  if(key==null){
    PropertiesService.getUserProperties().deleteProperty('api-key');
  } else {
    PropertiesService.getUserProperties().setProperty('api-key',key);}
}
function showAPIKEYPrompt_() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please enter your ncbi API key.',
    'Key:',
      ui.ButtonSet.OK_CANCEL);
  // Process response
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if(text==""){return null;}
    return text;
  }
function getKey() {
  //Logger.log(PropertiesService.getUserProperties().getProperty('api-key'));
  return PropertiesService.getUserProperties().getProperty('api-key');
}

// retmax get set
function setRetmax() {
  result=showRETMAXPrompt_();
  if(result==null){
    PropertiesService.getUserProperties().deleteProperty('retmax');
  } else {
    PropertiesService.getUserProperties().setProperty('retmax',result);}
}
function showRETMAXPrompt_() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please enter the desired retmax for E-utils fetch function.',
    'Number:',
      ui.ButtonSet.OK_CANCEL);
  // Process response
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if(isNaN(text)) {ui.alert( 'Invalid retmax submitted. Defaulting to 400.');
                 return "400";}
  if(+text>1000) {ui.alert('Retmax values greater than 1000 may result in an error for very large studies.');}
    return (+text).toFixed(0).toString();
  }
function getRetmax() {
  return PropertiesService.getUserProperties().getProperty('retmax');
}
