// Eutils path: esearch(GSE) -> esummary(GSE UID) -> esearch(bioproject) -> elink(bioproject UID) -> efetch(sra UID's)

// Add-on logic
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Enter API-key', functionName: 'setKey_'},
    {name: 'Begin Run Collection', functionName: 'openSidebar_'}
  ];
  spreadsheet.addMenu('SRA Run Collector', menuItems);
}
function onInstall() {
  onOpen();
}

function openSidebar_() {
  var html = HtmlService
      .createTemplateFromFile('sidebar')
      .evaluate()
      .setTitle('SRA Run Collector');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Metadata generation logic.
function validateResponse(response) {
  //var response = {accessions:['GSE99454'],parameters:{save_pref:false}};
  if(response.accessions[0]=="") {throw 'Please enter SRP, PRJNA, or GSE accessions for HTS studies.';}
  if(getKey_()==null || getKey_()==""){throw 'Please enter a ncbi API key through the Reannotator toolbar.';}
  var blankEntries = [];
  var vRegex = new RegExp("^GSE|^SRP|^PRJNA");
  for (var i = 0; i < response.accessions.length; i++) {
    // Check if PRJNA, SRP, or GSE. Also validate whether all are correctly formatted. Each has its own process.
     if (response.accessions[i]=='') {
      blankEntries.push(i);
      } else if(vRegex.exec(response.accessions[i])==null){
           throw response.accessions[i]+' is not a valid accession.';
      }
  }
  for (var i=0; i < blankEntries.length; i++) {
        response.accessions.splice(blankEntries[i]-i,1);
        }
  return response
}

function getMetadata(response) {
  //var response = {accessions:['GSE99454'],parameters:{save_pref:false}};
  var RETMAX = 400 // Global parameter. This limit will impact total I/O'
  // Creating Eutils instance
  var e = new Eutils_(getKey_(),RETMAX);
  var SRA_UIDs = {};
  for (var i = 0; i < response.accessions.length; i++) {
    // Check if PRJNA, SRP, or GSE. Also validate whether all are correctly formatted. Each has its own process.
    var PRJNA_regex = new RegExp("^PRJNA");
    var SRP_regex = new RegExp("^SRP");
    var GSE_regex = new RegExp("^GSE");    
     if(PRJNA_regex.exec(response.accessions[i])=='PRJNA'){
       // PRJNA PRJNA PRJNA PRJNA
     Utilities.sleep(e.api_rate);
     var BIOP_UID =  e.esearch(response.accessions[i],'bioproject').esearchresult.idlist[0];
      // First check for PRJNA route is whether esearch yielded no results.
     if(BIOP_UID==null){throw response.accessions[i]+' was not found.';}
     Utilities.sleep(e.api_rate);
     var SRA_ELINK = e.elink(BIOP_UID,'bioproject','sra');
     Utilities.sleep(e.api_rate);
       // Second check for PRJNA route is whether elink yielded no results.
     if(SRA_ELINK.linksets[0].linksetdbs==null) {throw response.accessions[i]+' has no associated run information.';}
     SRA_UIDs[response.accessions[i]] = SRA_ELINK.linksets[0].linksetdbs[0].links;
     
     } else if(SRP_regex.exec(response.accessions[i])=='SRP'){
      // SRP SRP SRP SRP SRP
      Utilities.sleep(e.api_rate);
      var SRP_UID = e.esearch(response.accessions[i],'sra');
      Utilities.sleep(e.api_rate);
       // First check for SRP is whether esearch yielded hits, which indicates associated runs. WIP may fail if more than 1:1 mapping for SRP
      if(SRP_UID.esearchresult.idlist.length<1) {throw response.accessions[i]+' has no associated run information.';}
       Utilities.sleep(e.api_rate);
      SRA_UIDs[response.accessions[i]] = SRP_UID.esearchresult.idlist;
      
    } else if(GSE_regex.exec(response.accessions[i])=='GSE'){
      // GSE GSE GSE GSE GSE
      Utilities.sleep(e.api_rate);
      var GDS_UID = e.esearch(response.accessions[i],'gds').esearchresult.idlist[0]; // extracts first entry from returned ID's, which is always the GEO study.
      // First check is simply that the esearch was not null
      if(GDS_UID==null){throw response.accessions[i]+' was not found.';}
      Utilities.sleep(e.api_rate);
      var GDS_SUMMARY = e.esummary(GDS_UID,'gds').result[GDS_UID];
      // Second check is that selected UID exactly matches the submitted accession
      if(response.accessions[i]!=GDS_SUMMARY.accession) {throw response.accessions[i]+' did not create a 1:1 match.';}
      Utilities.sleep(e.api_rate);
      var BIOP_UID = e.esearch(GDS_SUMMARY.bioproject,'bioproject').esearchresult.idlist[0];
      Utilities.sleep(e.api_rate);
      var SRA_ELINK = e.elink(BIOP_UID,'bioproject','sra');
      if(SRA_ELINK.linksets[0].linksetdbs==null) {throw response.accessions[i]+' has no associated run information.';}
      Utilities.sleep(e.api_rate);
      SRA_UIDs[response.accessions[i]] = SRA_ELINK.linksets[0].linksetdbs[0].links;
      }
  }
  // This avoids repetitive function calls in the above loop
  for (var key in SRA_UIDs) {
    Utilities.sleep(e.api_rate);
    SRA_UIDs[key] = XmlService.parse(e.efetch(SRA_UIDs[key].join(),'sra').getContentText());
  }
  Utilities.sleep(e.api_rate);
  metadata = processXML(SRA_UIDs);
  var header = ["SRR","Submitted Accession","Sample Accession","Title","Sample Attributes","Platform","Accession","Library Strategy","Library Layout","Library Source","Library Selection","Library Construction Protocol","Bases"];
  outputTable(metadata, header, response.parameters.save_pref);
  return 'Metadata gathered.';
}

function processXML(SRA_METADATA) {
      // Receives an XML object obtained from SRA.
      // Returns a LoL object corresponding to a table of metadata.
  for(var key in SRA_METADATA){
      var root = SRA_METADATA[key].getRootElement();
      var entries = new Array();
      entries = root.getChildren();
      var metadata=[]; // Where you would specify additional headers.
      for(var j = 0; j < entries.length; j++) {
        // Works by extracting data to a common variable, currentElement, which is then appended to the currentData, which is appended to the LoL table.
        var currentData = [key];
        var currentElement = entries[j].getChild('SAMPLE');
        currentData.push(currentElement.getChild('IDENTIFIERS').getChildText('PRIMARY_ID')); // Sample Accession
        
        var currentElement = entries[j].getChild('EXPERIMENT')
        currentData.push(currentElement.getChildText('TITLE')); // Sample title
        
        var currentElement = entries[j].getChild('SAMPLE'); // Fields from SAMPLE
        currentData.push(currentElement.getChild('SAMPLE_ATTRIBUTES')
                         .getChildren()
                         .map(function(x) {return x.getChildText('TAG')+': '+x.getChildText('VALUE')})
                         .join(" || ")); // Creates a string of sample attributes.
        var currentElement = entries[j].getChild('EXPERIMENT') // Fields from EXPERIMENT
        currentData.push(currentElement.getChild('PLATFORM').getChildren()[0].getChildText('INSTRUMENT_MODEL')); // Get platform
        currentData.push(currentElement.getChild('STUDY_REF').getAttributes().map(function(x){return x.getValue();}).join()); // Get accessions from EXPERIMENT node
        
        var currentElement = currentElement.getChild('DESIGN').getChild('LIBRARY_DESCRIPTOR');
        currentData.push(currentElement.getChildText('LIBRARY_STRATEGY'));
        currentData.push(currentElement.getChild('LIBRARY_LAYOUT').getChildren()[0].getName());
        currentData.push(currentElement.getChildText('LIBRARY_SOURCE'));
        currentData.push(currentElement.getChildText('LIBRARY_SELECTION'));
        currentData.push(currentElement.getChildText('LIBRARY_CONSTRUCTION_PROTOCOL'));
        
        // Getting run information, which may have a multiple mapping to the rest of the metadata.
        var currentData = entries[j]
           .getChild('RUN_SET')
           .getChildren()
           .map(function(x) {return [x.getChild('IDENTIFIERS').getChildText('PRIMARY_ID')].concat(currentData).concat([x.getAttribute('total_bases').getValue()])});
        metadata = metadata.concat(currentData);
        // Adds parsed metadata to the dictionary, overwriting the XML.
      }
        SRA_METADATA[key] = metadata;
  }
  return SRA_METADATA;
      
}

function outputTable(metadata,header,save_pref_param) {
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
    var range = cSheet.getRange(1,1,metadata[key].length+1,metadata[key][0].length);
    metadata[key].unshift(header);
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
    cSheet.setName('Reannotator_Results');
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
  this.key = key;
  this.keyBool = key!=null; // True if key exists.
  this.url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"; // hard coded
  this.retmax = retmax.toString();
  function key_to_rate(key) {if(key==null){return (1/2.75)*1000;} else {return (1/9.75)*1000;}}
  this.api_rate = key_to_rate(key);
  this.request = function(url) {
  // Ping the requested URL with GET
    return UrlFetchApp.fetch(url);
  };
  this.esearch = function(term,db) {
    var url = this.url+'esearch.fcgi?db='+db+'&term='+term+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(this.request(url));
  };
  this.elink = function(id,dbfrom,db) {
    var url = this.url+'elink.fcgi?db='+db+'&dbfrom='+dbfrom+'&id='+id+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(this.request(url));
  };
  this.esummary = function(id,db) {
    var url = this.url+'esummary.fcgi?db='+db+'&id='+id+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(this.request(url));
  };
  this.efetch = function(id,db) {
    // returns an xml
    var url = this.url+'efetch.fcgi?db='+db+'&id='+id+'&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return this.request(url);
  };
}

// API-key get set
function setKey_() {
  key=showAPIKEYPrompt();
  PropertiesService.getUserProperties().setProperty('api-key',key);
}
function showAPIKEYPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please enter your ncbi API key.',
    'Key:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  // Should test that API key works
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    // User clicked "Cancel".
    ui.alert('API key not set.');
  }
  return text
  }
function getKey_() { 
  return PropertiesService.getUserProperties().getProperty('api-key')
}
