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
      .setTitle('SRA Metadata Collector');
  SpreadsheetApp.getUi().showSidebar(html);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// Metadata generation logic
function getMetadataRequest(response) {
  // Pass through function.
  //Various interesting test cases.
  //var response = {accessions:['GSE48812'],parameters:{save_pref:false}};
  //var response = {accessions:['PRJDB7477'],parameters:{save_pref:false}};
  //var response = {accessions:['SRP158392'],parameters:{save_pref:false}};
  //var response = {accessions:['PRJNA523380'],parameters:{save_pref:false}}; // Extremely large.
  //var response = {accessions:['HT-29'],parameters:{save_pref:false}};
  //var response = {accessions:['SRP009888'],parameters:{save_pref:false}}; // Does not have sample attributes.
  //var response = {accessions:['GSE100839'],parameters:{save_pref:false}}; // Has technical replicates.
  //
  return getMetadata_(response);
}

function getMetadata_(response) {
  // This function runs through the metadata collection process, utilizing various other functions and the eutils object.
  var RETMAX = (getRetmax_() != null && !isNaN(getRetmax_())) ? +getRetmax_():400; // Global parameter. This limit will impact total I/O'
  // Creating Eutils instance
  var e = new Eutils_(getKey_(),RETMAX);
  // Creating dictionary for UIDs (values) for samples associated with a given accessions (keys)
  var SRA_UIDs = {};
  // Creating a regex pattern for determing what type of accession has been entered.
  var regex = new RegExp("^GSE|^SRP|^PRJNA|^PRJDB");
  for (var i = 0; i < response.accessions.length; i++) {
    // Check if PRJNA, SRP, or GSE. Also validate whether all are correctly formatted. Each has its own process.
    Utilities.sleep(e.api_rate); 
    if(regex.exec(response.accessions[i])=='PRJNA' || regex.exec(response.accessions[i])=='PRJDB'){
       // Bioproject Bioproject Bioproject
       var BIOP_UID =  e.esearch(response.accessions[i],'sra',use_history=true);
       SRA_UIDs[response.accessions[i]] = [BIOP_UID.esearchresult.webenv,BIOP_UID.esearchresult.querykey];
      
     } else if(regex.exec(response.accessions[i])=='GSE'){
       // GDS GDS GDS GDS GDS GDS
        var GDS_UID = e.esearch(response.accessions[i],'gds').esearchresult.idlist[0]; // extracts first entry from returned ID's, which is always the GEO study.
        if(GDS_UID==null){throw response.accessions[i]+' was not found.';}
        Utilities.sleep(e.api_rate);
        var SRA_ELINK = e.elink(GDS_UID,'gds','sra',cmd='neighbor_history');
        SRA_UIDs[response.accessions[i]] = [SRA_ELINK.linksets[0].webenv,SRA_ELINK.linksets[0].linksetdbhistories[0].querykey];
      
    } else {
        // SRA SRA SRA SRA SRA. Any entered term that does not match one of the above formats will be simply searched with SRA.
        var SRP_UID = e.esearch(response.accessions[i],'sra',use_history=true);
        SRA_UIDs[response.accessions[i]] = [SRP_UID.esearchresult.webenv,SRP_UID.esearchresult.querykey];
    }
  }
  // This avoids repetitive function calls in the above loop
  for (var key in SRA_UIDs) {
    Utilities.sleep(e.api_rate);
    SRA_UIDs[key] = XmlService.parse(e.efetch(null,'sra',SRA_UIDs[key]).getContentText());
  }
  var header = ["SRR","Submitted Term","Sample Accession","Title","Sample Attributes","Platform","Accessions","Library Strategy","Library Layout","Library Source","Library Selection","Library Construction Protocol","Bases"];
  metadata = processXML_(SRA_UIDs);
  outputTable_(metadata, header, response.parameters.save_pref);
  return 'Metadata gathered.';
}

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

function outputTable_(metadata,header,save_pref_param) {
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

function Eutils_(key,retmax) {
  var handler = {
  request:function request(url) {return UrlFetchApp.fetch(url);}
  }
  this.key = key;
  this.keyBool = key!=null; // True if key exists.
  this.url = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/'; // Needs to be the Eutils url.
  this.retmax = retmax.toString();
  function key_to_rate(key) {if(key==null){return (1/2.75)*1000;} else {return (1/9.75)*1000;}}
  this.api_rate = key_to_rate(key);
  this.esearch = function(term,db,usehistory) {
    if(usehistory==null){
      usehistory=false
    }
    var url = this.url+'esearch.fcgi?db='+db+'&term='+term+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    if(usehistory) {url=url+'&usehistory=y';}
    return JSON.parse(handler.request(url));
  };
  this.elink = function(id,dbfrom,db,cmd) {
    var url = this.url+'elink.fcgi?db='+db+'&dbfrom='+dbfrom+'&id='+id+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    if(cmd!=null) {url=url+'&cmd='+cmd;}
    return JSON.parse(handler.request(url));
  };
  this.esummary = function(id,db) {
    var url = this.url+'esummary.fcgi?db='+db+'&id='+id+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(handler.request(url));
  };
  this.efetch = function(id,db,webenv) {
    // returns an xml
    if(webenv!=null){
      var url = this.url+'efetch.fcgi?db='+db+'&query_key='+webenv[1]+'&WebEnv='+webenv[0]+'&retmax='+this.retmax;
      if(this.keyBool) {url=url+'&api_key='+this.key;}
      return handler.request(url)
      } else {
        var url = this.url+'efetch.fcgi?db='+db+'&id='+id+'&retmax='+this.retmax;
        if(this.keyBool) {url=url+'&api_key='+this.key;}
        return handler.request(url);
      }};
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
function getKey_() {
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
function getRetmax_() {
  return PropertiesService.getUserProperties().getProperty('retmax');
}
