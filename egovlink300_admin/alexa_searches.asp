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
<!--#include file="va_searches.asp"-->

  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
Sub ShowNewsItems()
	Dim sSql, oItems

	sSQL = "SELECT itemdate, itemtitle, itemtext, ISNULL(itemlinkurl,'') AS itemlinkurl "
 sSQL = sSQL & " FROM egov_news_items I, organizations O "
	sSQL = sSQL & " WHERE O.isadminnewssource = 1 "
 sSQL = sSQL & " AND O.orgid = I.orgid "
 sSQL = sSQL & " AND I.itemdisplay = 1 "
	sSQL = sSQL & " AND (publicationstart is null OR publicationstart <= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
	sSQL = sSQL & " AND (publicationend is null OR publicationend >= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
	sSQL = sSQL & " ORDER BY itemorder"

	set oItems = Server.CreateObject("ADODB.Recordset")
	oItems.Open sSQL, Application("DSN"), 3, 1

	if not oItems.eof then
  		do while not oItems.eof
    			response.write "<strong>" & oItems("itemdate") &  " &mdash; " & oItems("itemtitle") &  "</strong><br />" & oItems("itemtext") & vbcrlf

    			if oItems("itemlinkurl") <> "" then
      				response.write "<br /><a href=""" & oItems("itemlinkurl") & """ target=""_topwin""><strong>More &gt;&gt;</strong></a>" & vbcrlf
       end if

    			response.write "<br /><br />" & vbcrlf

    			oItems.movenext
    loop
	else
  		response.write "There are no updates at the present time. Please check back again for future updates." & vbcrlf
	end if

	oItems.Close
	set oItems = nothing 

End Sub 

'------------------------------------------------------------------------------
sub checkForDelegateAssignments(iUserID)

  if iUserID <> "" then
     sSQL = "SELECT userid, firstname, lastname "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE delegateid = " & iUserID
     sSQL = sSQL & " ORDER BY lastname, firstname "

    	set oDelegates = Server.CreateObject("ADODB.Recordset")
    	oDelegates.Open sSQL, Application("DSN"), 3, 1

     if not oDelegates.eof then
        iRowCount = 0

        response.write "<fieldset>" & vbcrlf
        response.write "  <legend style=""margin-bottom:10px;"">You have been assigned as a delegate for:&nbsp;</legend>" & vbcrlf

        do while not oDelegates.eof
           iRowCount = iRowCount + 1

           if iRowCount > 1 then
              response.write "<br />" & vbcrlf
           end if

           response.write oDelegates("firstname") & " " & oDelegates("lastname")

           oDelegates.movenext
        loop

        response.write "</fieldset><br />" & vbcrlf

     end if

     oDelegates.close
     set oDelegates = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub checkForDelegateAssigned(iUserID)

  if iUserID <> "" then
     sSQL = "SELECT userid, firstname, lastname "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE userid = (select delegateid "
     sSQL = sSQL &                 " from users "
     sSQL = sSQL &                 " where userid = " & iUserID & ") "
     sSQL = sSQL & " ORDER BY lastname, firstname "

    	set oDelegates = Server.CreateObject("ADODB.Recordset")
    	oDelegates.Open sSQL, Application("DSN"), 3, 1

     if not oDelegates.eof then
        iRowCount = 0

        response.write "<fieldset>" & vbcrlf
        response.write "  <legend style=""margin-bottom:10px;"">You have assigned the following as your delegate:&nbsp;</legend>" & vbcrlf

        do while not oDelegates.eof
           iRowCount = iRowCount + 1

           if iRowCount > 1 then
              response.write "<br />" & vbcrlf
           end if

           response.write oDelegates("firstname") & " " & oDelegates("lastname")

           oDelegates.movenext
        loop

        response.write "</fieldset><br />" & vbcrlf

     end if

     oDelegates.close
     set oDelegates = nothing

  end if

end sub
%>
