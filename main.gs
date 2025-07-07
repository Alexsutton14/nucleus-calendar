/*

Library ID:
1WgP91v450WvBiCQ5tYh3jEnjCvFlisEz0hi88ab-7K2_w_h2hZ7jUVdJ

*/

function doGet(e){
  return Nucleus.serverGet(e);
}

function enrolFormSubmit(formObject){
  Nucleus.enrolFormSubmit(formObject);
}

function resetServerUserProps(){
  Nucleus.resetServerUserProps();
}

function clearCurrentTrigger(){
  Nucleus.clearCurrentTrigger();
}

function FetchAndUpdateCalendar(){
  Nucleus.FetchAndUpdateCalendar();
}

function checkForExpiringCookie(){
  Nucleus.checkForExpiringCookie();
}

function getScriptURL() {
  return ScriptApp.getService().getUrl();
}
