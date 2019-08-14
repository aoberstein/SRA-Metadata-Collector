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
