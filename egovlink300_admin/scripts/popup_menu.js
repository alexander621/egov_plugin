<!--
var lastItem = null;

function checkIn(name) {
	if (lastItem != null) {
		hideMenu(lastItem);
		//window.setTimeout("hideMenu('" + lastItem + "')", 200);
	}
  if (name == "") {
		lastItem = null;
	}
	else {
		showMenu(name);
		//window.setTimeout("showMenu('" + name + "')", 200);
		lastItem = name;
	}
}

function checkOut(name) {
	obj = eval("document.all." + name);
  if (event.toElement.id != name && !obj.contains(event.toElement))
		checkIn("");
}

function showMenu(name) {
	eval("document.all." + name + ".style.visibility = 'visible'");
}

function hideMenu(name) {
	eval("document.all." + name + ".style.visibility = 'hidden'");
}
//-->