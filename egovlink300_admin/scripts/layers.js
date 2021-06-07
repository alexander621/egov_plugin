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
function displayLayer(layerName, yesOrNo) 
{
	if (yesOrNo)
	{
		eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.display=""');
		if (document.getElementById(layerName + "img"))
		{
			document.getElementById(layerName + "img").innerHTML = "&ndash;";
		}
	}
	else
	{
		eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.display="none"');
		if (document.getElementById(layerName + "img"))
		{
			document.getElementById(layerName + "img").innerHTML = "+";
		}
	}
}

//---------------------TOGGLE DISPLAY (IE only)
function toggleDisplay(layerName) 
{
	if (eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.display') == "")
		displayLayer(layerName, false);
	else
		displayLayer(layerName, true);
}

	// The following div hide and show code by Steve Loar 6/26/2008
	// This version works in IE and FireFox
		function ChangeLayerDisplay( layerName, bShow ) 
		{
			if (bShow)
			{
				document.getElementById(layerName).style.display = "";
				if (document.getElementById(layerName + "img"))
				{
					document.getElementById(layerName + "img").innerHTML = "&ndash;";
				}
			}
			else
			{
				document.getElementById(layerName).style.display = "none";
				if (document.getElementById(layerName + "img"))
				{
					document.getElementById(layerName + "img").innerHTML = "+";
				}
			}
		}

		//---------------------TOGGLE DISPLAY (IE and FireFox)
		function toggleDisplayShow( layerName ) 
		{
			if (document.getElementById(layerName).style.display != "none")
				ChangeLayerDisplay( layerName, false );
			else
				ChangeLayerDisplay( layerName, true );
		}

//-->