<!--
//---------------------GLOBAL VARIABLES
var layerRef="null", layerStyleRef="null", styleSwitch="null";

//---------------------INITIALIZATION
if (navigator.appName == "Netscape") {
	appCode = "ns";
	layerStyleRef="layer.";
	layerRef="document.layers";
	styleSwitch="";
}
else {
  appCode = "ie";
	layerStyleRef="layer.style.";
	layerRef="document.all";
	styleSwitch=".style";
}

//---------------------SHOW LAYER
function showLayer(layerName) {
	eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="visible"');
}

//---------------------HIDE LAYER
function hideLayer(layerName) {
	eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="hidden"');
}

//---------------------WRITE TO LAYER
function writeToLayer(layerName, text) {
  if (appCode == "ns") {
    //***** for this to work in netscape the layer must have the position style value set *****
    eval(layerRef+'["'+layerName+'"].document.write("'+text+'")');
    eval(layerRef+'["'+layerName+'"].document.close()');
  }
  else {
    eval(layerRef+'["'+layerName+'"].innerHTML="'+text+'"');
  }
}

//---------------------READ FROM LAYER (IE Only)
function readFromLayer(layerName) {
  return eval(layerRef+'["'+layerName+'"].innerHTML');
}

//---------------------DISPLAY LAYER (IE only)
function displayLayer(layerName, yesOrNo) {
  if (yesOrNo)
	  eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.display=""');
	else
	  eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.display="none"');
}

//---------------------TOGGLE DISPLAY (IE only)
function toggleDisplay(layerName) {
  if (eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.display') == "")
    displayLayer(layerName, false);
  else
    displayLayer(layerName, true);
}
//-->