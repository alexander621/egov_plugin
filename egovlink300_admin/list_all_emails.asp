<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
  Dim sError

 'Set Timezone information into session
  session("iUserOffset") = request.cookies("tz")

 'Override of value from common.asp
  sLevel = ""

%>
<html>
<head>
  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="global.css" />
  <link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />

  <script language="javascript" src="scripts/modules.js"></script>
  <script language="javascript" src="scripts/ajaxLib.js"></script>

<script language="javascript">
<!--

//Set timezone in cookie to retrieve later
var d=new Date();
if(d.getTimezoneOffset) {
		 var iMinutes = d.getTimezoneOffset();
		 document.cookie = "tz=" + iMinutes;
}

//function doPicker(sFormField) {
//		w = (screen.width - 350)/2;
//		h = (screen.height - 350)/2;
//		eval('window.open("sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
//}

function doPicker(sFormField) {
  w = 600;
  h = 500;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";

  //lcl_showFolderStart = "&folderStart=published_documents";
  //lcl_showFolderStart = "&folderStart=unpublished_documents";
  lcl_showFolderStart = "&folderStart=CITY_ROOT";


  pickerURL  = "sitelinker/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += lcl_showFolderStart;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
		    var caretPos = textEl.caretPos;
  			 caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? text + ' ' : text;
  } else {
   			textEl.value = textEl.value + text;
	 }
}

function openWin2(url, name) {
		popupWin = window.open(url, name, "resizable,width=500,height=450");
}

function SetLocation( iUserId ) {
		//alert(document.LocationForm.locationid.options[document.LocationForm.locationid.selectedIndex].value);
		//doAjax('../includes/setuserlocation.asp', 'locationid=' + document.LocationForm.locationid.options[document.LocationForm.locationid.selectedIndex].value, 'LocationSet', 'get', '0');
		doAjax('includes/setuserlocation.asp', 'locationid=' + document.LocationForm.locationid.options[document.LocationForm.locationid.selectedIndex].value, '', 'get', '0');
}

function LocationSet( iLocationId ) {
		// Nothing happens here
		//alert( iLocationId );
}

//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content" style="width:auto;">
 	<div id="centercontent">

<table id="bodytable" border="0" cellpadding="0" cellspacing="0" class="start">
  <tr valign="top">
    	<td>
	<h2 id="title">All Citizen Emails</h2>

	<%

	sSQL = "SELECT DISTINCT useremail FROM egov_users WHERE orgid = '" & session("orgid") & "' and isdeleted = 0 and useremail IS NOT NULL and useremail <> ''"
'	sSQL = "SELECT top 2 useremail FROM egov_users WHERE orgid = '" & session("orgid") & "' and isdeleted = 0 and useremail IS NOT NULL and useremail <> ''"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1

	if not oRs.EOF then response.write oRs.RecordCount & " Email Addresses<br /><br />"

	response.write "<button onclick=""CopyToClipboard('getallemails')"" >COPY</button><br /><br />"
	response.write "<div id=""getallemails"">"
	Do While Not oRs.EOF
		response.write oRs("useremail") & "<br />"
		oRs.MoveNext
	loop

	oRs.Close
	Set oRs = Nothing
	response.write "</div>"
	%>
	<script>
	function CopyToClipboard(containerid) {
		if (document.selection) { 
    		var range = document.body.createTextRange();
    		range.moveToElementText(document.getElementById(containerid));
    		range.select().createTextRange();
    		document.execCommand("Copy"); 
		
		} else if (window.getSelection) {
    		var range = document.createRange();
     		range.selectNode(document.getElementById(containerid));
     		window.getSelection().addRange(range);
     		document.execCommand("Copy");
     		//alert("text copied") 
		}}

	</script>
      </td>
  </tr>
</table>

  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="admin_footer.asp"-->  

</body>
</html>
