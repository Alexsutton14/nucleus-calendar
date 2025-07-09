/*

Library ScriptID:
1WgP91v450WvBiCQ5tYh3jEnjCvFlisEz0hi88ab-7K2_w_h2hZ7jUVdJ

*/

function doGet(e){
  return NucleusLibrary.serverGet(e);
}

function enrolFormSubmit(formObject){
  NucleusLibrary.enrolFormSubmit(formObject);
}

function resetServerUserProps(){
  NucleusLibrary.resetServerUserProps();
}

function clearCurrentTrigger(){
  NucleusLibrary.clearCurrentTrigger();
}

function FetchAndUpdateCalendar(){
  NucleusLibrary.FetchAndUpdateCalendar();
}

function DailyChecks(){
  NucleusLibrary.DailyChecks();
}

function getScriptURL() {
  return ScriptApp.getService().getUrl();
}
