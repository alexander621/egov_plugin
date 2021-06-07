<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="formletters_global_functions.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"form letters") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
	<title>E-Gov Administration Consule {Form Letter Management}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

 <script language="javascript" src="../scripts/modules.js"></script>
 <script language="javascript" src="../scripts/tooltip_new.js"></script>

	<script language="javascript">
	<!--
		function confirm_delete(iletterid) {
    sname = document.getElementById("title_" + iletterid).innerHTML;

 			//input_box=confirm("Are you sure you want to delete '" + sname + "'? \nAll parameters will be lost.");
    input_box=confirm("Are you sure you want to delete '" + sname + "'?");

	 		if(input_box==true) { 
    			//Delete has been verified
				   location.href='delete_letter.asp?iletterid='+ iletterid;
 			}
		}

		function previewMe(iD) {
    w = 600;
    h = 500;
    t = (screen.availHeight/2)-(h/2);
    l = (screen.availWidth/2)-(w/2);
    eval('window.open("preview_letter.asp?FLid=' + iD + '", "_previewtemplate' + iD +'", "width='+w+',height='+h+',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=yes,resizable=yes,status=yes,menubar=0")');

			 //var newFile = "preview_letter.asp?FLid=" + iD + "";
 			//newWin = window.open(newFile,'popupName','width=600,height=500,toolbars=no,left=50,top=50,scrollbars=yes,resizable=yes,status=yes')
	 		//newWin.focus();
		}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
	<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <font size=""+1""><strong>Manage Form Letters</strong></font><br />" & vbcrlf

 'blnCanEditForms = HasPermission("CanEditActionForms") 
 	blnCanEditForms = True 

	 if blnCanEditForms then
     response.write "<p>" & vbcrlf
     response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td width=""60%"">" & vbcrlf
     'response.write "          <input type=""button"" name=""newLetterButton"" id=""newLetterButton"" value=""Create a New Form Letter"" class=""button"" onclick=""location.href='new_letter.asp';"" />" & vbcrlf
     response.write "          <input type=""button"" name=""newLetterButton"" id=""newLetterButton"" value=""Create a New Form Letter"" class=""button"" onclick=""location.href='manage_letter.asp';"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td width=""40%"" align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

                     subListFLs()
     response.write "</p>" & vbcrlf
  else
     response.write "<p>" & vbcrlf
     response.write "You do not have permission to access the <strong>E-Gov Form Letter section</strong>.  " & vbcrlf
     response.write "Please contact your E-Govlink administrator to inquire about gaining access to the <strong>E-Gov Forms Creator section</strong>." & vbcrlf
     response.write "</p>" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"-->  

<%

  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub subListFLs()
	iorgID = session("orgid")

	if iorgID = "" then
  		iorgID = -1
	end if

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM formletters "
 sSQL = sSQL & " WHERE orgid = " & iorgID
 sSQL = sSQL & " ORDER BY sequence"

	set oLetterList = Server.CreateObject("ADODB.Recordset")
	oLetterList.Open sSQL, Application("DSN"), 3, 1
	
	if not oLetterList.eof then
  		response.write "<div class=""shadow"">" & vbcrlf
		  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"">" & vbcrlf
  		response.write "  <tr>" & vbcrlf
    'response.write "      <th>&nbsp;</th>" & vbcrlf
    response.write "      <th align=""left"">Title</th>" & vbcrlf
    response.write "      <th colspan=""2"">Seq</th>" & vbcrlf
    response.write "      <th colspan=""2"">&nbsp;</th>" & vbcrlf
    response.write "  </tr>" & vbcrlf

    lcl_bgcolor = "#ffffff"

  		do while not oLetterList.eof
       lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
			
    			if IsNull(oLetterList("FLtitle")) OR oLetterList("FLtitle")="" then
     					iTitle = ""
    			elseif len(oLetterList("FLtitle")) > 50 then
      				iTitle = left(oLetterList("FLtitle"),40) & "..."
    			else
     					iTitle = oLetterList("FLtitle")
    			end if

       'lcl_edit_url  = "location.href='manage_letter.asp?iorgid=" & iorgid & "&iletterid=" & oLetterList("FLid") & "';"
       'lcl_move_up   = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move UP');"" onmouseout=""tooltip.hide();"" onclick=""location.href='order_letter.asp?direction=UP&iletterid=" & oLetterList("FLid") & "&iorgid=" & iorgid & "'"" /><br />" & vbcrlf
       'lcl_move_down = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move DOWN');"" onmouseout=""tooltip.hide();"" onclick=""location.href='order_letter.asp?direction=DOWN&iletterid=" & oLetterList("FLid") & "&iorgid=" & iorgid & "'"" />" & vbcrlf
       lcl_edit_url  = "location.href='manage_letter.asp?iletterid=" & oLetterList("FLid") & "';"
       lcl_move_up   = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move UP');"" onmouseout=""tooltip.hide();"" onclick=""location.href='order_letter.asp?direction=UP&iletterid=" & oLetterList("FLid") & "'"" /><br />" & vbcrlf
       lcl_move_down = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move DOWN');"" onmouseout=""tooltip.hide();"" onclick=""location.href='order_letter.asp?direction=DOWN&iletterid=" & oLetterList("FLid") & "'"" />" & vbcrlf

       response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
			    response.write "      <td class=""formlist"" width=""55%"" align=""left"" onclick=""" & lcl_edit_url & """><span id=""title_" & oLetterList("FLid") & """>" & iTitle & "</span></td>" & vbcrlf
    			response.write "      <td class=""formlist"" nowrap=""nowrap"" onclick=""" & lcl_edit_url & """>(" & oLetterList("sequence") & ")</td>" & vbcrlf
    			response.write "      <td class=""formlist"" nowrap=""nowrap"">" & lcl_move_up & lcl_move_down & "</td>" & vbcrlf
    			response.write "      <td class=""formlist"" nowrap=""nowrap""><input type=""button"" name=""viewTemplateButton"" id=""viewTemplateButton"" value=""View Form Letter Template"" class=""button"" onclick=""previewMe(" &oLetterList("FLid") & ")"" /></td>" & vbcrlf
    			response.write "      <td class=""formlist"" nowrap=""nowrap""><input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirm_delete('" & oLetterList("FLid") & "');"" />" & vbcrlf
       'response.write "          <a href=""javascript:previewMe(" &oLetterList("FLid") & ")"">View Form Letter Template</a> | " & vbcrlf
    			'response.write "          <a href=""order_letter.asp?direction=UP&iletterid=" & oLetterList("FLid") & "&iorgid=" & iorgid & """>Move Up</a> | " & vbcrlf
    			'response.write "          <a href=""order_letter.asp?direction=DOWN&iletterid=" & oLetterList("FLid") & "&iorgid=" & iorgid & """>Move Down</a> | " & vbcrlf
			 			'response.write "          <a href=""javascript:confirm_delete('" & oLetterList("FLid") & "','" & UCASE(iTitle) & "');"">Delete</a>" & vbcrlf
       response.write "      </td>" & vbcrlf
    			response.write "  </tr>" & vbcrlf

       oLetterList.MoveNext
    loop

  		response.write "</table>" & vbcrlf
  		response.write "</div>" & vbcrlf

 else

		  response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;""><i>No Form Letters</i></p>" & vbcrlf

	end if

	set oLetterList = nothing

end sub
%>