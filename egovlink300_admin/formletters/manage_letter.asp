<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="formletters_global_functions.asp" //-->
<%
 Dim blnAllMergeFields

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "form letters") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if this is a new form letter or if we are maintaining one.
 if request("iletterid") <> "" then
    iLetterID                 = request("iletterid")
    lcl_button_label          = "UPDATE"
    lcl_showMergeFieldsButton = True
 else
    iLetterID                 = 0
    lcl_button_label          = "ADD"
    lcl_showMergeFieldsButton = False
 end if

 if request.servervariables("REQUEST_METHOD") = "POST" then
   	select case request("TASK")

		    case "new_question"
      	 		'Add new question to form
	        		subAddQuestion session("orgid"),iFormID

    	 case else
   			    'Default action

  	 end select
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
 	<title>E-Gov Administration Console {Maintain Form Letter}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript">
<!--
function doPicker(sFormField) {
		w = (screen.width - 350)/2;
		h = (screen.height - 350)/2;
		eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function fnCheckSubject() {
  if (document.NewEvent.Subject.value != '') {
  				return true;
		}else{
  				return false;
		}
}

function previewMe() {
		var FormData = document.NewEvent;
		if (FormData.FLtitle.value == "") {
  				alert("Please enter a Form Letter Title");
		  		FormData.FLtitle.focus();
				  return false;
		}

		if (FormData.FLbody.value == "") {
  				alert("Please enter the Form Letter body");
		  		FormData.FLbody.focus();
				  return false;
		}

 	alert("This is a PREVIEW only, Form Letter has not been saved.");
		var newFile = "preview_letter.asp?FLbody=" + FormData.FLbody.value + "&FLtitle=" + FormData.FLtitle.value + "";
		newWin = window.open(newFile,'popupName','width=600,height=500,toolbars=no,left=50,top=50,scrollbars=yes,resizable=yes,status=yes')
		newWin.focus();
		return false;
}

function addHTMLTag(p_tag) {
  var lcl_body = document.getElementById("FLbody").value;

  if(p_tag=="BOLD") {
     lcl_body = lcl_body + " <STRONG></STRONG>";
  }else if(p_tag=="ITALICS") {
     lcl_body = lcl_body + " <EM></EM>";
  }else if(p_tag=="H1") {
     lcl_body = lcl_body + " <H1></H1>";
  }else if(p_tag=="H2") {
     lcl_body = lcl_body + " <H2></H2>";
  }else if(p_tag=="H3") {
     lcl_body = lcl_body + " <H3></H3>";
  }else if(p_tag=="LINK") {
     lcl_body = lcl_body + " <A HREF=\"url goes here\"></A>";
  }else if(p_tag=="IMG") {
     lcl_body = lcl_body + " <IMG SRC=\"image filename goes here\" WIDTH=\"0\" HEIGHT=\"0\" />";
  }else if(p_tag=="FONT") {
     lcl_body = lcl_body + " <FONT style=\"font-size: 10pt;\"></FONT>";
  }else if(p_tag=="BR") {
     lcl_body = lcl_body + "<BR />";
  }else if(p_tag=="P") {
     lcl_body = lcl_body + "<P>text goes here</P>";
  }else if(p_tag=="P_LEFT") {
     lcl_body = lcl_body + " <P align=\"LEFT\"></P>";
  }else if(p_tag=="P_CENTER") {
     lcl_body = lcl_body + " <P align=\"CENTER\"></P>";
  }else if(p_tag=="P_RIGHT") {
     lcl_body = lcl_body + " <P align=\"RIGHT\"></P>";
  }

  document.getElementById("FLbody").value = lcl_body;
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

<style>
		div.correctionsbox          {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
		div.correctionsboxnotfound  {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
		td.correctionslabel         {font-weight:bold;}
		th.corrections              {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
		input.correctionstextbox    {border: solid 1px #336699;width:400px;}
		.savemsg                    {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
</style>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">

<table border="0" cellspacing="0" cellpadding="0" width="600">
  <tr valign="top">
      <td>
          <font size="+1"><strong>Edit Form Letter</strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Return to Form Letter List" class="button" onclick="location.href='list_letter.asp';" />
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
</table>
<p>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
      <td width="40%" align="right" nowrap="nowrap"><strong>Common HTML formatting tags:</strong></td>
      <td><% displayButton "Bold" %></td>
      <td><% displayButton "Italics" %></td>
      <td><% displayButton "FONT" %></td>
      <td><% displayButton "H1" %></td>
      <td><% displayButton "H2" %></td>
      <td><% displayButton "H3" %></td>
      <td><% displayButton "LINK" %></td>
      <td><% displayButton "IMG" %></td>
      <td><% displayButton "BR" %></td>
      <td><% displayButton " P " %></td>
      <td align="center">
          Alignment:<br>
          <select name="p_format_alignment" onchange="addHTMLTag(this.value);">
            <option value=""></option>
            <option value="P_LEFT">LEFT</option>
            <option value="P_CENTER">CENTER</option>
            <option value="P_RIGHT">RIGHT</option>
          </select>
      </td>
      <td nowrap="nowrap">[<a href="http://www.w3schools.com/tags/default.asp" target="_blank">Additional TAGs</a>]</td>
  </tr>
</table>
<div class="shadow">
<table class="tablelist" cellspacing="0" cellpadding="2">
<form name="NewEvent" action="save_letter.asp" method="post">
		<tr><th class="corrections" colspan="2" align="left">Edit Form Letter</th></tr>
		<tr>
    		<td valign="top" style="padding:10px;">
 					<!--BEGIN: Form Letter -->
    						<input type="hidden" name="iLetterID" value="<%=iLetterID%>" />
    						<p>
   							<%GetFLs(iLetterID)%>
      				<p>
    						<input type="submit" value="<%=lcl_button_label%> Form Letter" class="button" />

        <%
          if lcl_showMergeFieldsButton then
             response.write "<input type=""button"" value=""Manage Merge Fields"" class=""button"" onclick=""location.href='manage_letter_to_forms.asp?iletterid=" & request("iletterid") & "';"" />" & vbcrlf
          end if
        %>
 					<!--END: Form Letter -->
  				</td>
  				<td valign="top" style="padding:10px;">
		       	<div style="padding : 5px; width : 250px; height : 500px; overflow : auto; ">
     					<!--BEGIN: Form Letter Dynamic Fields-->
     					<strong>Instructions</strong><br /><br />
      				Any of the fields below may be copied/pasted into the body of your form letters. 
     					Please copy exactly as they appear including the brackets and astericks.
<%
 'DISPLAY AVAILABLE MERGE FIELDS
		response.write "<p>"
		response.write "<strong>General Purpose Fields:</strong><br>"
		response.write "[*TodaysDate*]"
		response.write "</p>"

	'CONTACT FIELDS
		sSQL = "SELECT TOP 1 userfname,userlname,userbusinessname,useremail,userhomephone,userfax,useraddress,usercity,userstate,userzip "
  sSQL = sSQL & " FROM egov_users" 

		Set oUser = Server.CreateObject("ADODB.Recordset")
		oUser.Open sSQL, Application("DSN"), 3, 1

		response.write "<p>"
		response.write "<strong>Contact Fields:</strong><br />"
	 For Each Field In oUser.Fields
						if Field.Name <> "userpassword" then
  							response.write "[*" & Field.Name & "*]<br />"
						end if
		Next
		Set oUser = Nothing
		response.write "</p>"

	'ADDITIONAL INFORMATION FIELD
 	response.write "<p>"
		response.write "<strong>Additional Comment Field:</strong><br />"
		response.write "[*Admin_Additional_Comments*]"
		response.write "</p>"

	'TRACKING NUMBER FIELD
		response.write "<p>"
		response.write "<strong>Tracking Number Field:</strong><br />"
		response.write "[*Tracking Number*]"
		response.write "</p>"

	'CODE SECTIONS FIELDS
		response.write "<p>"
		response.write "<strong>Code Section Field:</strong><br />"
		response.write "[*Code_Sections*]"
		response.write "</p>"

	'ISSUE LOCATION FIELDS
		response.write "<p>"
		response.write "<strong>Issue Location Fields:</strong><br>"

	'GET ISSUE LOCATION FIELDS
		Call subListIssueLocationFields()
		response.write "</p>"

	'FORM FIELDS
		response.write "<p>"
		response.write "<strong>Form Fields:</strong><br />"

	'GET ADDITIONAL FIELD VALUES
		Call subListMergeFields(iLetterID,blnAllMergeFields)
		response.write "</p>"
%>
<!--END: Form Letter Dynamic Fields -->
		       	</div>
      </td>
		</tr>
</form>
</table>
		</div>
</div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
function GetFLs(iLetterID)

 lcl_fl_title      = ""
 lcl_fl_body       = ""
 blnAllMergeFields = False

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM FormLetters "
	sSQL = sSQL & " WHERE FLid=" & iLetterID
	sSQL = sSQL & " AND orgid="  & session("orgid")

	set oForm = Server.CreateObject("ADODB.Recordset")
	oForm.Open sSQL, Application("DSN"), 3, 1

	if not oForm.eof then
    lcl_fl_title = oForm("FLtitle")
    lcl_fl_body  = oForm("FLbody")

    if oForm("containsHTML") then
       lcl_checked_containsHTML = " checked=""checked"""
    else
       lcl_checked_containsHTML = ""
    end if

  		blnAllMergeFields = oForm("blnAllMergeFields")

 end if

 response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
 response.write "  <tr><td colspan=""2""><strong>Title:</strong></td></tr>" & vbcrlf
 response.write "  <tr><td colspan=""2""><input type=""text"" class=""question"" name=""FLtitle"" value=""" & lcl_fl_title & """ size=""72"" maxlength=""199"" /></td></tr>" & vbcrlf
 response.write "  <tr valign=""bottom"">" & vbcrlf
 response.write "      <td><strong>Body:</strong></td>" & vbcrlf
 response.write "      <td align=""right""><input type=""checkbox"" name=""containsHTML"" id=""containsHTML"" value=""Y""" & lcl_checked_containsHTML & " />&nbsp;Form Letter Body contains HTML</td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"">" & vbcrlf
 response.write "          <textarea name=""FLbody"" id=""FLbody"" class=""none"" rows=""40"" cols=""100"" style=""width: 550px; font-size: 10px; font-family: Verdana,Tahoma,Arial;"">" & lcl_fl_body & "</textarea>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf

	set oForm = nothing

end function

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
Function IsRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
  		sReturnValue = " <font color=red>*</font> "
	Else
  		sReturnValue = ""
	End If

	IsRequired = sReturnValue
End Function

'------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  strDB = Replace( strDB, "'", "\'" )
  strDB = Replace( strDB, chr(34), "\'" )
  strDB = Replace( strDB, ";", "\;" )
  strDB = Replace( strDB, "-", "\-" )
  JSsafe = strDB
End Function

'------------------------------------------------------------------------------
Sub subListMergeFields(iLetterID,blnAll)

'CONNECT TO DATABASE RETRIEVE VALUES FOR FORM FIELDS ASSOCIATED WITH THIS REQUEST
	If blnAll Then
	  	sSQL ="SELECT DISTINCT pdfformname, sequence "
    sSQL = sSQL & " FROM egov_action_form_questions "
    sSQL = sSQL & " WHERE pdfformname IS NOT NULL "
    sSQL = sSQL & " AND pdfformname <> '' "
    'sSQL = sSQL & " AND orgid='" & session("orgid") & "'"
    sSQL = sSQL & " ORDER BY sequence "
	Else
  		sSQL ="SELECT DISTINCT pdfformname, sequence "
    sSQL = sSQL & " FROM egov_letter_to_form "
    sSQL = sSQL & " INNER JOIN egov_action_form_questions ON egov_letter_to_form.formid = egov_action_form_questions.formid "
    sSQL = sSQL & " WHERE pdfformname IS NOT NULL "
    sSQL = sSQL & " AND pdfformname <> '' "
    'sSQL = sSQL & " AND orgid='" & session("orgid") & "'"
    sSQL = sSQL & " AND letterid='" & iLetterID & "'"
    sSQL = sSQL & " ORDER BY sequence "
	End If

	Set oDynamicFields = Server.CreateObject("ADODB.Recordset")
	oDynamicFields.Open sSQL, Application("DSN"), 3, 1

	If NOT oDynamicFields.EOF Then		

  	'LOOP THRU FORM FIELDS
		  Do While NOT oDynamicFields.EOF
  		  	If oDynamicFields("pdfformname") <> "" Then
		      		response.write "[*" & oDynamicFields("pdfformname") & "*]<br />"	
    			End If
    			oDynamicFields.MoveNext
  		Loop
	End If

End Sub

'------------------------------------------------------------------------------
Sub subListIssueLocationFields()

'CONNECT TO DATABASE RETRIEVE VALUES FOR FORM FIELDS ASSOCIATED WITH THIS REQUEST
	sSQL = "Select TOP 1 streetnumber,streetaddress,city,state,zip,comments, legaldescription,listedowner, parcelidnumber "
 sSQL = sSQL & " FROM egov_action_response_issue_location"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 3, 1

	For Each Field In oLocation.Fields
   		Response.Write "[*" & Field.Name & "*]<br />"
	Next

	Set oLocation = Nothing

End Sub

'------------------------------------------------------------------------------
sub displayButton(p_name)

  if p_name <> "" then
     response.write "<input type=""button"" name=""" & p_name & "Button"" id=""" & p_name & "Button"" value=""" & p_name & """ class=""button"" onclick=""addHTMLTag('" & ucase(trim(p_name)) & "');"" />" & vbcrlf
  end if

end sub
%>
