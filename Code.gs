// Eutils path: esearch(GSE) -> esummary(GSE UID) -> esearch(bioproject) -> elink(bioproject UID) -> efetch(sra UID's)

// Add-on logic
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Enter API-key', functionName: 'setKey'},
    {name: 'Begin Run Collection', functionName: 'openSidebar'}
  ];
  spreadsheet.addMenu('SRA Run Collector', menuItems);
}
function onInstall() {
  onOpen();
}

function openSidebar() {
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
  var blankEntries = [];
  var regex = new RegExp("^GSE|^SRP|^PRJNA");
  for (var i = 0; i < response.accessions.length; i++) {
    // Check if PRJNA, SRP, or GSE. Also validate whether all are correctly formatted. Each has its own process.
     if (response.accessions[i]=='') {
      blankEntries.push(i);
      } else if(regex.exec(response.accessions[i])==null){
           throw response.accessions[i]+' is not a valid accession.';
      }
  }
  for (var i=0; i < blankEntries.length; i++) {
        response.accessions.splice(blankEntries[i]-i,1);
        }
  return response
}

function getMetadataRequest(response) {
  // Pass through function.
  //var response = {accessions:['PRJNA433861'],parameters:{save_pref:false}};
  //
  return getMetadata_(response);
}

function getMetadata_(response) {
  // This function runs through the metadata collection process, utilizing various other functions and the eutils object.
  var RETMAX = 400 // Global parameter. This limit will impact total I/O'
  // Creating Eutils instance
  var e = new Eutils_(getKey_(),RETMAX);
  // Creating dictionary for UIDs (values) for samples associated with a given accessions (keys)
  var SRA_UIDs = {};
  // Creating a regex pattern for determing what type of accession has been entered.
  var regex = new RegExp("^GSE|^SRP|^PRJNA");
  for (var i = 0; i < response.accessions.length; i++) {
    // Check if PRJNA, SRP, or GSE. Also validate whether all are correctly formatted. Each has its own process.
    Utilities.sleep(e.api_rate); 
    if(regex.exec(response.accessions[i])=='PRJNA'){
       // PRJNA PRJNA PRJNA PRJNA
       var BIOP_UID =  e.esearch(response.accessions[i],'sra').esearchresult.idlist;
       // First check for PRJNA route is whether esearch yielded no results.
       if(BIOP_UID==null){throw response.accessions[i]+' was not found.';}
       SRA_UIDs[response.accessions[i]] = BIOP_UID;
     
     } else if(regex.exec(response.accessions[i])=='SRP'){
        // SRP SRP SRP SRP SRP
        var SRP_UID = e.esearch(response.accessions[i],'sra');
        // First check for SRP is whether esearch yielded hits, which indicates associated runs. WIP may fail if more than 1:1 mapping for SRP
        if(SRP_UID.esearchresult.idlist.length<1) {throw response.accessions[i]+' has no associated run information.';}
        SRA_UIDs[response.accessions[i]] = SRP_UID.esearchresult.idlist;
      
    } else if(regex.exec(response.accessions[i])=='GSE'){
        // GSE GSE GSE GSE GSE
        var GDS_UID = e.esearch(response.accessions[i],'gds').esearchresult.idlist[0]; // extracts first entry from returned ID's, which is always the GEO study.
        if(GDS_UID==null){throw response.accessions[i]+' was not found.';}
        Utilities.sleep(e.api_rate);
        var SRA_ELINK = e.elink(GDS_UID,'gds','sra');
        if(SRA_ELINK.linksets[0].linksetdbs==null) {throw response.accessions[i]+' has no associated run information.';}
        SRA_UIDs[response.accessions[i]] = SRA_ELINK.linksets[0].linksetdbs[0].links;
    }
  }
  // This avoids repetitive function calls in the above loop
  for (var key in SRA_UIDs) {
    Utilities.sleep(e.api_rate);
    SRA_UIDs[key] = XmlService.parse(e.efetch(SRA_UIDs[key].join(),'sra').getContentText());
  }
  var header = ["SRR","Submitted Accession","Sample Accession","Title","Sample Attributes","Platform","Accessions","Library Strategy","Library Layout","Library Source","Library Selection","Library Construction Protocol","Bases"];
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
        currentData.push(currentElement.getChild('SAMPLE_ATTRIBUTES')
                         .getChildren()
                         .map(function(x) {return x.getChildText('TAG')+': '+x.getChildText('VALUE')})
                         .join(" || ")); // Creates a string of sample attributes.
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
           .map(function(x) {return [x.getChild('IDENTIFIERS').getChildText('PRIMARY_ID')].concat(currentData).concat([x.getAttribute('total_bases').getValue()])});
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
  this.esearch = function(term,db) {
    var url = this.url+'esearch.fcgi?db='+db+'&term='+term+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(handler.request(url));
  };
  this.elink = function(id,dbfrom,db) {
    var url = this.url+'elink.fcgi?db='+db+'&dbfrom='+dbfrom+'&id='+id+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(handler.request(url));
  };
  this.esummary = function(id,db) {
    var url = this.url+'esummary.fcgi?db='+db+'&id='+id+'&retmode=json&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return JSON.parse(handler.request(url));
  };
  this.efetch = function(id,db) {
    // returns an xml
    var url = this.url+'efetch.fcgi?db='+db+'&id='+id+'&retmax='+this.retmax;
    if(this.keyBool) {url=url+'&api_key='+this.key;}
    return handler.request(url);
  };
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
