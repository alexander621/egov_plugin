<!-- #include file="../includes/common.asp" //-->
<!-- #include file="picker_global_functions.asp" //-->
<%
  sLocationName = trim(GetVirtualName(Session("OrgID")))
  lcl_message   = request("message")
  lcl_name      = request("name")

 'Determine which sections to display
  lcl_displayDocuments  = "N"
  lcl_displayActionLine = "N"
  lcl_displayPayments   = "N"
  lcl_displayURL        = "N"

  if request("displayDocuments") = "Y" then
     lcl_displayDocuments = "Y"
  end if

  if request("displayActionLine") = "Y" then
     lcl_displayActionLine = "Y"
  end if

  if request("displayPayments") = "Y" then
     lcl_displayPayments = "Y"
  end if

  if request("displayURL") = "Y" then
     lcl_displayURL = "Y"
  end if

 'This will return the link within an anchor tag (HTML).
  lcl_returnAsHTMLLink    = "Y"
  lcl_includeClickCounter = "N"
  lcl_returnOnlyFileName  = "N"

  if request("returnAsHTMLLink") <> "" then
     lcl_returnAsHTMLLink = request("returnAsHTMLLink")
  end if

 'ONLY if the "returnAsHTMLLink" <> "Y" can we return only the filename.
  if request("returnOnlyFileName") = "Y" then
     lcl_returnOnlyFileName = "Y"
  end if

 'Now check to see if "click-counter" is to be included in the HTML anchor tags ONLY.
  if lcl_returnAsHTMLLink = "Y" AND request("includeClickCounter") <> "" then
     lcl_includeClickCounter = request("includeClickCounter")
  end if

 'Setup additional features
  lcl_displayLinkText = "Y"

  if request("displayLinkText") <> "" then
     lcl_displayLinkText = request("displayLinkText")
  end if

 'Get the starting folder
  lcl_folderStart = getStartingFolder(request("folderStart"))
  'if request("folderStart") <> "" then
  '   lcl_folderStart = request("folderStart")
  'else
  '   lcl_folderStart = "published_documents"
  'end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Choose File...}</title>
  <style type="text/css">
  <!--
    td, input, select {font-family:MS Sans Serif,Tahoma,Arial; font-size:11px;}
  //-->
  </style>

<script language="javascript">
<!--
function MakeActive(id) {
  document.all.exdoc.style.display  = "none";
  document.all.newdoc.style.display = "none";
  document.all.newpay.style.display = "none";
  document.all.newurl.style.display = "none";

  eval("document.all." + id + ".style.display = 'block'");

  if(id == "newurl") {
     document.getElementById("Url").readOnly = false;
  }
}
    
function saveSelection() {
  var objParent=window.opener;
  var path=document.frmFilePath.FilePath.value;
  document.frmFilePath.FileLink.value = "<a href='" + document.all.currentpath.value + "/" + path + "'>" + path + "</a>";
  objParent.addItem.itemID.value=document.all.currentpath.value + "/" + path;
  objParent.addItem.link.value=path;
  if(objParent.addItem.title.value=='')objParent.addItem.title.value=path;
  window.close();
}

function buildlink(sFormField) {
  var sLocation = '<%=sLocationName%>';
  var sLink     = '';
		var path      = document.frmFilePath.FilePath.value;
  var sLinkName = '';
//  var sLinkName = document.frmAddArticle.ALinkName.value;

		if (document.frmFilePath.FilePath.value == "") {
   			alert("Please select a file.");
    		return;
		}

  if (document.frmFilePath.LinkName) {
      if(document.frmFilePath.LinkName.value != '') {
         sLinkName = document.frmFilePath.LinkName.value;

      }
  }

  if (sLinkName == '') {
      sLinkName = path;
  }

  lcl_currentpath  = document.getElementById("currentpath");
  lcl_originalpath = document.getElementById("original_currentpath");

  lcl_cleanpath = lcl_currentpath.value;
  lcl_cleanpath = lcl_cleanpath.replace(/\/public_documents300/,"");

//http://www.egovlink.com/public_documents300/bristol/published_documents/Human%20Resources/Job%20openings/Notice%20of%20Juvenile%20Judge%20Vacancy.pdf

  lcl_url = "<%=Application("newLinkPicker_url")%>" + lcl_currentpath.value + "/" + path;
  lcl_url = lcl_url.replace("/custom/pub","");

  //Check the format of the return link
  if("<%=lcl_returnAsHTMLLink%>" == "Y") {
     //sLink = "<a target='_egovlink' href='" + document.all.currentpath.value + "/" + path + "'>" + sLinkName + "</a>";

     if("<%=lcl_includeClickCounter%>" == "Y") {
        sID         = "<%=month(now) & day(now) & year(now) & hour(now) & minute(now) & second(now)%>";
        sFieldID    = " id='" + sID + "'";
        lcl_onclick = " onclick='countClick(\"" + sID + "\")'";
     }else{
        sFieldID    = "";
        lcl_onclick = "";
     }

     //lcl_cleanpath = lcl_currentpath.value;
     //lcl_cleanpath = lcl_cleanpath.replace(/\/public_documents300/,"");

     //Replace the "spaces" with "%20"
  	  while (lcl_url.indexOf(" ") > -1) {
          		lcl_url = lcl_url.replace(" ","%20");
  	  }

     sLink = "<a" + sFieldID + lcl_onclick + " href='" + lcl_url + "' target='_egovlink'>" + sLinkName + "</a>";
  }else{
     //sLink = document.all.currentpath.value + "/" + path;
     //Determine if we are only returning the folder(s) + filename
     if("<%=lcl_returnOnlyFileName%>" == "Y") {
        sLink = lcl_cleanpath + "/" + path;
        sLink = sLink.replace("/custom/pub","");
        sLink = sLink.replace("/<%=sLocationName%>","");
        sLink = sLink.replace("/unpublished_documents/","");
     }else{
        sLink = lcl_url;
     }
  }
  //sLink = sLink.replace(lcl_originalpath.value,"");
  var oFormField = 'window.opener.document.' + sFormField; 

  //This handles the ckeditor
  if(sFormField == 'frmlocation.sHTMLBody') {
     window.opener.updateEditor_or_Field('EDITOR', sLink);
  } else {
     window.opener.insertAtCaret(eval(oFormField),sLink);
  }

<% if session("orgid") = 178 then %>
    window.opener.sendemail.checked = true;
<% end if %>
  window.close();
}

function buildactionlink(sFormField){
  var sLocation = "<%=sLocationName%>";
  var iFormID = document.frmAddArticle.iFormID.value;
  var sLinkName = document.frmAddArticle.ALinkName.value;
  if (sLinkName =='') {
    		sLinkName = document.frmAddArticle.AFormName.value;
  }

  //Check the format of the return link
  if("<%=lcl_returnAsHTMLLink%>" == "Y") {
     //Replace the "spaces" with "%20"
  	  while (sLocation.indexOf(" ") > -1) {
          		sLocation = sLocation.replace(" ","%20");
  	  }

     sLink = "<a target='_egovlink' href='http://www.egovlink.com/" + sLocation + "/action.asp?actionid=" + iFormID + "'>" + sLinkName + "</a>";
  }else{
     sLink = "http://www.egovlink.com/" + sLocation + "/action.asp?actionid=" + iFormID;
  }

  var oFormField = 'window.opener.document.' + sFormField; 

  //This handles the ckeditor
  if(sFormField == 'frmlocation.sHTMLBody') {
     window.opener.updateEditor_or_Field('EDITOR', sLink);
  } else {
     window.opener.insertAtCaret(eval(oFormField),sLink);
  }

<% if session("orgid") = 178 then %>
    window.opener.sendemail.checked = true;
<% end if %>
  window.close();
}

function buildpaymentlink(sFormField){
  var sLocation = "<%=sLocationName%>";
  var iFormID = document.frmPaymentLink.iFormID.value;
  var sLinkName = document.frmPaymentLink.ALinkName.value;
  if (sLinkName =='') {
	    	sLinkName = document.frmPaymentLink.AFormName.value;
  }

  //Check the format of the return link
  if("<%=lcl_returnAsHTMLLink%>" == "Y") {
     //Replace the "spaces" with "%20"
  	  while (sLocation.indexOf(" ") > -1) {
          		sLocation = sLocation.replace(" ","%20");
  	  }

     sLink = "<a target='_egovlink' href='http://www.egovlink.com/" + sLocation + "/payment.asp?paymenttype=" + iFormID + "'>" + sLinkName + "</a>";
  }else{
     sLink = "http://www.egovlink.com/" + sLocation + "/payment.asp?paymenttype=" + iFormID;
  }

  var oFormField = 'window.opener.document.' + sFormField; 

  //This handles the ckeditor
  if(sFormField == 'frmlocation.sHTMLBody') {
     window.opener.updateEditor_or_Field('EDITOR', sLink);
  } else {
     window.opener.insertAtCaret(eval(oFormField),sLink);
  }

<% if session("orgid") = 178 then %>
    window.opener.sendemail.checked = true;
<% end if %>
  window.close();
}

function buildwebpagelink(sFormField){
  var sURL = document.frmURL.UrlType.value + document.frmURL.Url.value;
  var sLinkName = document.frmURL.UrlName.value;
  if (sLinkName =='') {
		    sLinkName = document.frmURL.Url.value;
  }

     if("<%=lcl_includeClickCounter%>" == "Y") {
        sID         = "<%=month(now) & day(now) & year(now) & hour(now) & minute(now) & second(now)%>";
        sFieldID    = " id='" + sID + "'";
        lcl_onclick = " onclick='countClick(\"" + sID + "\")'";
     }else{
        sFieldID    = "";
        lcl_onclick = "";
     }

  //Check the format of the return link
  if("<%=lcl_returnAsHTMLLink%>" == "Y") {
     //Replace the "spaces" with "%20"
  	  while (sURL.indexOf(" ") > -1) {
          		sURL = sURL.replace(" ","%20");
  	  }

     sLink = "<a" + sFieldID + lcl_onclick + " target='_egovlink' href='" + sURL + "'>" + sLinkName + "</a>";
  }else{
     sLink = sURL;
  }

  var oFormField = 'window.opener.document.' + sFormField; 

  //This handles the ckeditor
  if(sFormField == 'frmlocation.sHTMLBody') {
     window.opener.updateEditor_or_Field('EDITOR', sLink);
  } else {
     window.opener.insertAtCaret(eval(oFormField),sLink);
  }

<% if session("orgid") = 178 then %>
    window.opener.sendemail.checked = true;
<% end if %>
  window.close();
}

function myFunction() {
  alert(document.frmAddArticle.currentfolderpath.value);
}

function doPicker() {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("../picker_new/default.asp", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function setupOriginalCurrentPath() {
  lcl_currentpath  = document.getElementById("currentpath");
  lcl_originalpath = document.getElementById("original_currentpath");

  if(lcl_originalpath.value == "") {
     lcl_originalpath.value = lcl_currentpath.value;
  }
}
//-->
</script>
</head>
<body bgcolor="#d4d0c8" leftmargin="2" topmargin="0" onload="setupOriginalCurrentPath()">
  <input type="hidden" name="currentpath" id="currentpath" size="80" />
  <input type="hidden" name="original_currentpath" id="original_currentpath" size="80" />
<table border="0" cellpadding="3" cellspacing="0">
  <tr>
      <td>&nbsp;</td>
      <td>
          <input type="text" name="currentfolder" style="height:20px; width:250px;" readonly>&nbsp;
          <a href="#" style="color:#0000ff" onclick="explorer.window.history.back();" name="anchorBack">
            <img src="images/up.gif" alt="Back" border="0" align="absmiddle" />
          </a>
      </td>
  </tr>
  <tr>
      <td rowspan="3" valign="top">
        <%
          lcl_menu_url = "menu.asp"
          lcl_menu_url = lcl_menu_url & "?folderStart="       & lcl_folderStart
          lcl_menu_url = lcl_menu_url & "&displayDocuments="  & lcl_displayDocuments
          lcl_menu_url = lcl_menu_url & "&displayActionLine=" & lcl_displayActionLine
          lcl_menu_url = lcl_menu_url & "&displayPayments="   & lcl_displayPayments
          lcl_menu_url = lcl_menu_url & "&displayURL="        & lcl_displayURL
          lcl_menu_url = lcl_menu_url & "&displayLinkText="   & lcl_displayLinkText

          response.write "<iframe name=""menu"" width=""100"" height=""265"" src="""& lcl_menu_url & """></iframe>" & vbcrlf
        %>
      </td>
  </tr>
  <tr>
      <td valign="top">
          <%
            lcl_loadtree_url = "loadtree.asp"
            lcl_loadtree_url = lcl_loadtree_url & "?path=/public_documents300/custom/pub/" & sLocationName & lcl_folderStart
            lcl_loadtree_url = lcl_loadtree_url & "&displayDocuments="  & lcl_displayDocuments
            lcl_loadtree_url = lcl_loadtree_url & "&displayActionLine=" & lcl_displayActionLine
            lcl_loadtree_url = lcl_loadtree_url & "&displayPayments="   & lcl_displayPayments
            lcl_loadtree_url = lcl_loadtree_url & "&displayURL="        & lcl_displayURL
            lcl_loadtree_url = lcl_loadtree_url & "&displayLinkText="   & lcl_displayLinkText

            response.write "<iframe id=""explorer"" name=""explorer"" width=""400"" height=""250"" src=""" & lcl_loadtree_url & """></iframe>" & vbcrlf

            buildDiv "exdoc",  "frmFilePath",    lcl_displayLinkText, "Link Text", "LinkName",  "Y", "File Name", "FilePath",  lcl_message, lcl_name
            buildDiv "newdoc", "frmAddArticle",  lcl_displayLinkText, "Link Text", "ALinkName", "Y", "Form Name", "AFormName", lcl_message, lcl_name
            buildDiv "newpay", "frmPaymentLink", lcl_displayLinkText, "Link Text", "ALinkName", "Y", "Form Name", "AFormName", lcl_message, lcl_name
            buildDiv "newurl", "frmURL",         lcl_displayLinkText, "Name",      "UrlName",   "Y", "URL",       "Url",       lcl_message, lcl_name
          %>
      </td>
  </tr>
</table>
</body>
</html>
<%
'------------------------------------------------------------------------------
sub buildDiv(iDivID, iFormName, iShowLinkTextField, iLinkLabel, iLinkFieldName, _
             iShowFileNameField, iFileLabel, iFileFieldName, iMessage, iName)

 'Determine if the <DIV> is initially displayed or not.
  if UCASE(iDivID) <> "EXDOC" then
     lcl_display_div = " style=""display:none"""
  else
     lcl_display_div = ""
  end if

 'Set the width of the input fields based on the iFormName
  if UCASE(iFormName) = "FRMURL" then
     lcl_field_width = "188"
  else
     lcl_field_width = "250"
  end if

 'Setup the Add Link button
  if UCASE(iDivID) = "EXDOC" then
    lcl_button_label = "Select"
    lcl_onclick      = "buildlink"
  elseif UCASE(iDivID) = "NEWDOC" then
    lcl_button_label = "Add Link"
    lcl_onclick      = "buildactionlink"
  elseif UCASE(iDivID) = "NEWPAY" then
    lcl_button_label = "Add Link"
    lcl_onclick      = "buildpaymentlink"
  elseif UCASE(iDivID) = "NEWURL" then
    lcl_button_label = "Add Link"
    lcl_onclick      = "buildwebpagelink"
  end if

 'Determine which fields are displayed.
  lcl_showLinkTextField  = "N"
  lcl_showFieldNameField = "N"

  if iShowLinkTextField <> "" then
     lcl_showLinkTextField = iShowLinkTextField
  end if

  if iShowFileNameField <> "" then
     lcl_showFileNameField = iShowFileNameField
  end if

  response.write "<div id=""" & iDivID & """" & lcl_display_div & ">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""400"">" & vbcrlf
  response.write "  <form name=""" & iFormName & """>" & vbcrlf

  if iFormName = "frmAddArticle" OR iFormName = "frmPaymentLink" then
     response.write "    <input type=""hidden"" name=""currentfolderpath"" id=""currentfolderpath"" />" & vbcrlf
     response.write "	   <input type=""hidden"" name=""iFormID"" name=""iFormID"" />" & vbcrlf
  end if

  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" style=""padding-top:5px;"">" & vbcrlf
  response.write "          <table>" & vbcrlf

 'BEGIN: Link Text ------------------------------------------------------------
  if lcl_showLinkTextField = "Y" then
     response.write "            <tr>" & vbcrlf
     response.write "                <td nowrap=""nowrap"">" & iLinkLabel & ":&nbsp;&nbsp;</td>" & vbcrlf
     response.write "                <td><input type=""text"" name=""" & iLinkFieldName & """ id=""" & iLinkFieldName & """ style=""width:250px; height:20px;"" /></td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Link Text --------------------------------------------------------------

 'BEGIN: File Name ------------------------------------------------------------
  if lcl_showFileNameField = "Y" then
     response.write "            <tr>" & vbcrlf
     response.write "                <td nowrap=""nowrap"">" & iFileLabel & ":&nbsp;&nbsp;</td>" & vbcrlf
     response.write "                <td>" & vbcrlf

     if iFormName = "frmURL" then
        response.write "                    <select name=""UrlType"" id=""UrlType"">" & vbcrlf
        response.write "                      <option value=""http://"">http://</option>" & vbcrlf
        response.write "                      <option value=""https://"">https://</option>" & vbcrlf
        response.write "                      <option value=""mailto:"">mailto:</option>" & vbcrlf
        response.write "                      <option value=""ftp://"">ftp://</option>" & vbcrlf
        response.write "                    </select>" & vbcrlf
     end if

     response.write "                    <input type=""text"" name=""" & iFileFieldName & """ id=""" & iFileFieldName & """ style=""width:" & lcl_field_width & "px; height:20px;"" readonly />" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: File Name --------------------------------------------------------------

  response.write "          </table>" & vbcrlf
  response.write "          <br />" & vbcrlf

  if iMessage <> "" then
     response.write iMessage & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "      <td valign=""top"" align=""right"" style=""padding-top:5px;"">" & vbcrlf
  response.write "          <input type=""button"" value=""" & lcl_button_label & """ style=""width:80px; height:22px;"" onClick=""javascript:" & lcl_onclick & "('" & iName & "');"" /><br />" & vbcrlf
  response.write "          <img src=""images/spacer.gif"" width=""1"" height=""5"" /><br />" & vbcrlf
  response.write "          <input type=""button"" value=""Cancel"" style=""width:80px; height:22px;"" onclick=""window.close();"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</form>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub
%>
