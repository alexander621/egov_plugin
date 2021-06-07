<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
	lenErr = ""
	genErr = ""

 'Check to see if the feature is offline
  if isFeatureOffline("action line") = "Y" then
     response.redirect "outage_feature_offline.asp"
  end if

  sLevel     = "../"     'Override of value from common.asp
  lcl_hidden = "hidden"  'Show/Hide hidden fields.  HIDDEN = Hide, TEXT = Show

  if not UserHasPermission( Session("UserId"), "form creator" ) then
 	   response.redirect sLevel & "permissiondenied.asp"
  end if

  iFormID = request("iformid")
  iOrgID  = request("iorgid")
  sTask   = ""

  if request("task") <> "" then
     sTask = request("task")
     sTask = dbsafe(sTask)
     sTask = ucase(sTask)
  end if

  if request.servervariables("REQUEST_METHOD") = "POST" then
  	  subSaveValues iFormID, sTask
  end if

 'PROCESS ENABLE REQUEST FROM FORM LIST PAGE
  if sTask = "ENABLEFORM" then
  	   subSaveValues iFormID, sTask
  end if 

 'PROCESS INTERNAL ONLY REQUEST
  if sTask = "INTERNAL" then
 	   subSaveValues iFormID, sTask
  end if

 'BEGIN: Add Question Builder Information --------------------------------------
  sDisplayFormQuestion = ""
  sFormQuestion        = ""

  'if iFieldType <> "0" then
     sFormQuestion = buildQuestion(sTask, iFormID)

     sDisplayFormQuestion = "<form name=""frmSaveField"" id=""frmSaveField"" action=""edit_form.asp"" method=""post"" >" & vbcrlf
     sDisplayFormQuestion = sDisplayFormQuestion & "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iFormID & """ />" & vbcrlf
     sDisplayFormQuestion = sDisplayFormQuestion & "	 <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
     sDisplayFormQuestion = sDisplayFormQuestion & "	 <input type=""hidden"" name=""task"" id=""task"" value=""" & sTask & """ />" & vbcrlf
     sDisplayFormQuestion = sDisplayFormQuestion & sFormQuestion & vbrlf
     sDisplayFormQuestion = sDisplayFormQuestion & "</form>" & vbcrlf
  'End If 
'END: Add Question Builder Information ----------------------------------------

%>
<html>
<head>
	<title>E-Gov Administration Console { Maintain Form }</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<style type="text/css">
  .subLabel {
      color: #c00000;
  }

  #formQuestionDiv {
      border: 1pt solid #808080;
      border-radius: 5px;
      margin-top: 5px;
      padding: 10px;
  }

  #formQuestionDiv textarea {
      width: 100%;
      height: 50px;
  }
 
</style>

 	<script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

<script type="text/javascript">

$(document).ready(function() {
    setMaxLength();

    $('#returnButton').click(function() {
        location.href='manage_form.asp?iformid=<%=iFormID%>';
    });

    $('#saveButton').click(function() {
        $('#frmSaveField').submit();
    });
});
</script>

</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "    <table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td valign=""top"">" & vbcrlf
  response.write "              <h3>Forms - Edit Form</h3>" & vbcrlf
  response.write "              <input type=""button"" name=""returnButton"" id=""returnButton"" value=""Back"" class=""button"" />" & vbcrlf
  response.write "              <div id=""formQuestionDiv"">" & sDisplayFormQuestion & "</div>" & vbcrlf
  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf
  response.write "	 </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
function buildQuestion(iTask, iFormID)
  dim lcl_return, sTask, sFormID, sColumnName, sSQL, sSQLCustom, sMask
  dim arrNameList, sLabel, sSubLabel, sSaveButtonColspan
  dim sStreetNumberInputType, sStreetAddressInputType

  lcl_return         = ""
  sTask              = ""
  sFormID            = 0
  sColumnName        = ""
  sMask              = ""
  arrNameList        = ""
  sLabel             = "Form Name"
  sSubLabel          = "Change/Update the form name below."
  sSaveButtonColspan = "1"

  if iTask <> "" then
     sTask = dbsafe(iTask)
     sTask = ucase(sTask)
  end if

  if iFormID <> "" then
     sFormID = clng(iFormID)
  end if

 '-----------------------------------------------------------------------------
  select case sTask
    case "NAME"
        sColumnName = "action_form_name"
    case "INTRO"
        sColumnName = "action_form_description"
        sLabel      = "Intro Text"
        sSubLabel   = "Change/update the intro text below."
    case "FOOTER"
        sColumnName = "action_form_footer"
        sLabel      = "Footer Text"
        sSubLabel   = "Change/update the footer text below."
    case "ENOTE"
        sColumnName = "action_form_emergency_text"
        sLabel      = "Emergency Text"
        sSubLabel   = "Change/update the emergency text below. (max 1024 chars) " & lenErr
    case "CUSTOMFOILEMAILEDITS"
        sColumnName = "customFOILEmailEdits"
        sLabel      = "Custom FOIL Email Edits"
        sSubLabel   = "Change/Update the custom FOIL email edits."
   '---------------------------------------------------------------------------
    case "CONTACT"
        sColumnName        = "action_form_contact_mask"
        sLabel             = "Contact Options"
        sSubLabel          = "Change/Update the contact options below."
        sSaveButtonColspan = "2"
   '---------------------------------------------------------------------------
    case "ISSUENAME"
        sColumnName        = "''"
        sSQLCustom         = ", issuelocationname, issuelocationdesc, issuequestion "
        sLabel             = "Name"
        sSubLabel          = "Change/Update the issue location name below (Maximum 512 characters)."
        sSaveButtonColspan = "2"
    case else
        sColumnName = "''"
  end select
 '-----------------------------------------------------------------------------
  
 	sSQL = "SELECT " & sColumnName & " as dbColumnName "
  sSQL = sSQL & sSQLCustom
  sSQL = sSQL & " FROM egov_action_request_forms "
  sSQL = sSQL & " WHERE action_form_id = " & sFormID

 	Set oForm = Server.CreateObject("ADODB.Recordset")
 	oForm.Open sSQL, Application("DSN"), 3, 1
	
 	if not oForm.eof then
     lcl_return = "<table>" & vbcrlf
     lcl_return = lcl_return & "  <tr>" & vbcrlf
     lcl_return = lcl_return & "      <td colspan=""" & sSaveButtonColspan & """>" & vbcrlf
     lcl_return = lcl_return & "          <strong>" & sLabel & "</strong>: " & vbcrlf
     lcl_return = lcl_return & "          <span class=""subLabel"">" & sSubLabel & genErr & "</span>" & vbcrlf
     lcl_return = lcl_return & "      </td>" & vbcrlf
     lcl_return = lcl_return & "  </tr>" & vbcrlf

     if sTask = "CONTACT" then
        sMask       = trim(oForm("dbColumnName"))
        arrNameList = array("", "First Name", "Last Name", "Business Name", "Email", "Daytime Phone", "Fax", "Street", "City", "State","Zip")

        for iList = 1 to 10
           lcl_return = lcl_return & "  <tr>" & vbcrlf
           lcl_return = lcl_return & "      <td align=""right""><strong>" & arrNameList(iList) & "</strong></td>" & vbcrlf
           lcl_return = lcl_return & "      <td>" & vbcrlf
           lcl_return = lcl_return & "          <input type=""checkbox"" name=""maskshow" & iList & """ id=""maskshow" & iList & """ onclick=""if(document.frmSaveField.maskshow" & iList &".checked==false){document.frmSaveField.maskrequired" & iList &".checked=false};""" & IsDisplay(sMask,iList) & " /> Show on Form." & vbcrlf
           lcl_return = lcl_return & "          <input type=""checkbox"" name=""maskrequired" & iList & """ id=""maskrequired" & iList & """ onclick=""document.frmSaveField.maskshow" & iList &".checked=true;""" & IsRequired(sMask,iList) & " />Is Required." & vbcrlf
           lcl_return = lcl_return & "      </td>" & vbcrlf
           lcl_return = lcl_return & "  </tr>" & vbcrlf
        next

     elseif sTask = "ISSUENAME" then
        sIssueName = oForm("issuelocationname")
        sIssueDesc = oForm("issuelocationdesc")
        sIssueQues = oForm("issuequestion")

        if trim(sIssueName) = "" OR isnull(sIssueName) then
           sIssueName = "Issue/Problem Location"
        end if

        if trim(sIssueDesc) = "" OR isnull(sIssueDesc) then
           sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". Provide any additional information on problem location in the box below."
        end if

        if trim(sIssueQuestion) = "" OR isnull(sIssueQuestion) then
           sIssueQuestion = "Provide any additional information on problem location in the box below."
        end if

      		lcl_return = lcl_return & "  <tr>" & vbcrlf
      		lcl_return = lcl_return & "      <td>" & vbcrlf
      		lcl_return = lcl_return & "          <textarea class=""PROMPT"" name=""IssueName"" id=""IssueName"" maxlength=""512"">" & sIssueName & "</textarea>" & vbcrlf
      		lcl_return = lcl_return & "      </td>" & vbcrlf
      		lcl_return = lcl_return & "  </tr>" & vbcrlf

      		lcl_return = lcl_return & "  <tr><td>&nbsp;</td></tr>" & vbcrlf
        lcl_return = lcl_return & "  <tr>" & vbcrlf
        lcl_return = lcl_return & "      <td>" & vbcrlf
        lcl_return = lcl_return & "          <strong>Description</strong>: " & vbcrlf
        lcl_return = lcl_return & "          <span class=""subLabel"">Change/update the issue location description below.</span>" & vbcrlf
        lcl_return = lcl_return & "      </td>" & vbcrlf
        lcl_return = lcl_return & "  </tr>" & vbcrlf
      		lcl_return = lcl_return & "  <tr>" & vbcrlf
      		lcl_return = lcl_return & "      <td>" & vbcrlf
      		lcl_return = lcl_return & "          <textarea class=""PROMPT"" name=""IssueDesc"" id=""IssueDesc"">" & sIssueDesc & "</textarea>" & vbcrlf
      		lcl_return = lcl_return & "      </td>" & vbcrlf
      		lcl_return = lcl_return & "  </tr>" & vbcrlf

      		lcl_return = lcl_return & "  <tr><td>&nbsp;</td></tr>" & vbcrlf
        lcl_return = lcl_return & "  <tr>" & vbcrlf
        lcl_return = lcl_return & "      <td>" & vbcrlf
        lcl_return = lcl_return & "          <strong>Question</strong>: " & vbcrlf
        lcl_return = lcl_return & "          <span class=""subLabel"">Change/Update the issue location question below.</span>" & vbcrlf
        lcl_return = lcl_return & "      </td>" & vbcrlf
        lcl_return = lcl_return & "  </tr>" & vbcrlf
      		lcl_return = lcl_return & "  <tr>" & vbcrlf
      		lcl_return = lcl_return & "      <td>" & vbcrlf
      		lcl_return = lcl_return & "          <textarea class=""PROMPT"" name=""IssueQues"" id=""IssueQues"">" & sIssueQues & "</textarea>" & vbcrlf
      		lcl_return = lcl_return & "      </td>" & vbcrlf
      		lcl_return = lcl_return & "  </tr>" & vbcrlf

     else  'sTask = "NAME", "INTRO", "FOOTER", "ENOTE"
			 fieldVal = oForm("dbColumnName")
			 if request(sTask) <> "" then fieldVal = request(sTask)

      		lcl_return = lcl_return & "  <tr>" & vbcrlf
      		lcl_return = lcl_return & "      <td>" & vbcrlf
      		lcl_return = lcl_return & "          <textarea class=""PROMPT"" name=""" & sTask & """>" & fieldVal & "</textarea>" & vbcrlf
      		lcl_return = lcl_return & "      </td>" & vbcrlf
      		lcl_return = lcl_return & "  </tr>" & vbcrlf
     end if

   		lcl_return = lcl_return & "  <tr>" & vbcrlf
     lcl_return = lcl_return & "      <td colspan=""" & sSaveButtonColspan & """>" & vbcrlf
     lcl_return = lcl_return & "          <input type=""button"" name=""saveButton"" id=""saveButton"" class=""button"" value=""Save Changes"" />" & vbcrlf
   	 lcl_return = lcl_return & "      </td>" & vbcrlf
     lcl_return = lcl_return & "  </tr>" & vbcrlf
     lcl_return = lcl_return & "</table>" & vbcrlf

 	end if

 	set oForm = nothing 

  buildQuestion = lcl_return

end function

'------------------------------------------------------------------------------
sub subSaveValues(iFieldID,sTask)

 sSQL = "UPDATE egov_action_request_forms SET "

	select case UCASE(sTask)

	Case "NAME"
  sSQL = sSQL & "action_form_name='"           & DBsafe(request("NAME"))          & "' "
	
	Case "ENOTE"
  sSQL = sSQL & "action_form_emergency_text='" & DBsafe(request("ENOTE"))        & "' "
	
	Case "INTRO"
  sSQL = sSQL & "action_form_description='"    & DBsafe(request("intro"))        & "' "

	Case "FOOTER"
  sSQL = sSQL & "action_form_footer='"         & DBsafe(request("footer"))       & "' "

	Case "CONTACT"
  sSQL = sSQL & "action_form_contact_mask='"   & DBsafe(fnGetMask())              & "' "

 Case "CUSTOMFOILEMAILEDITS"
  sSQL = sSQL & "customFOILEmailEdits='"       & dbsafe(request("customFOILEmailEdits")) & "' "

	Case "EMERGENCYNOTE"
  sSQL = sSQL & "action_form_emergency_note='" & DBsafe(request("emergencynote")) & "' "

	Case "PUBLICSEARCH"
  sSQL = sSQL & "publicsearchrequests='" & DBsafe(request("publicsearchnote")) & "' "
	
	Case "ENABLEFORM"
  sSQL = sSQL & "action_form_enabled='"        & request("blnenabled")            & "', "

  if request("blnenabled") then
     sSQL = sSQL & "showInALSearch = 1"
  else
     sSQL = sSQL & "showinALSearch = 0"
  end if
		
	Case "INTERNAL"
  sSQL = sSQL & "action_form_internal='"       & request("blnInternal")           & "' "
	
	Case "ENABLEISSUE"
  sSQL = sSQL & "action_form_display_issue='"  & request("ENABLEISSUE")           & "' "

	Case "ISSUE"
  sSQL = sSQL & "issuestreetnumberinputtype = '" & request("selNumberInputType")  & "', "
  sSQL = sSQL & "issuestreetaddressinputtype='"  & request("selAddressInputType") & "', "
  sSQL = sSQL & "action_form_issue_mask='"       & DBsafe(fnGetIssueMask())       & "' "

 Case "ISSUELOCADDINFO"
  sSQL = sSQL & "hideIssueLocAddInfo = " & request("ISSUELOCADDINFO")

	Case "ISSUENAME"
  sSQL = sSQL & "issuelocationname='" & DBsafe(request("IssueName")) & "', "
  sSQL = sSQL & "issuelocationdesc='" & DBsafe(request("Issuedesc")) & "', "
  sSQL = sSQL & "issuequestion='"     & DBsafe(request("IssueQues")) & "' "

	Case "ENABLEFEE"
  sSQL = sSQL & "action_form_display_fees='" & request("ENABLEFEE") & "' "

	Case "RESOLVEDSTATUS"
  if request("RESOLVEDSTATUS") = "Y" then
     lcl_status = "N"
  else
     lcl_status = "Y"
  end if

  sSQL = sSQL & "action_form_resolved_status='" & lcl_status & "' "

 Case "ENABLEMOBILEOPTIONSTAKEPIC"
    sSQL = sSQL & "display_mobileoptions_takepic = '" & request("ENABLEMOBILEOPTIONSTAKEPIC") & "' "

 Case "ENABLEMOBILEOPTIONSGEOLOC"
    sSQL = sSQL & "display_mobileoptions_geoloc = '" & request("ENABLEMOBILEOPTIONSGEOLOC") & "' "
 Case "ENABLESHOWMAPINPUT"
    sSQL = sSQL & "showmapinput = '" & request("ENABLESHOWMAPINPUT") & "' "

	Case "SAVEAPPSETTINGS"
	if request("formobile") = "ON" then
  		sSQL = sSQL & "formobile='1', "
	else
  		sSQL = sSQL & "formobile='0', "
	end if
  sSQL = sSQL & "mobilename='" & DBsafe(request("mobilename")) & "', "
  sSQL = sSQL & "mobilehelptext='" & DBsafe(request("mobilehelptext")) & "' "
	Case "SAVECUSTOMURL"
  		sSQL = sSQL & "redirectURL='" & DBsafe(request("redirectURL")) & "' "


	Case Else

	End Select

 sSQL = sSQL & " WHERE action_form_id='" & iFieldID & "'"

 response.write sSQL

	errNum = 0
	on error resume next
	set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.Open sSQL, Application("DSN") , 3, 1
	set oSave = nothing
	errNum = err.number
	on error goto 0

	if errNum = 0 then
		' REDIRECT TO MANANGE FORM PAGE
		If UCASE(sTask) <> "ENABLEFORM" AND UCASE(sTask) <> "INTERNAL" Then
  			response.redirect("manage_form.asp?iformid=" & iFormID)
		Else
		  	response.redirect("list_forms.asp")
		End If
	else
		if UCASE(sTask) = "ENOTE" then
			lenErr = "<p>You've exceeded the maximum length.  Please shorten your message.</p>"
		else
			genErr = "<p>Sorry, your data could not be saved.</p>"
		end if
	End If

End Sub

'------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
' FUNCTION ISDISPLAY(SMASK,IFIELD)
'------------------------------------------------------------------------------
Function IsDisplay(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "1" or sValue = "2" Then
  		sReturnValue = " checked=""checked"""
	Else
		  sReturnValue = ""
	End If

	IsDisplay = sReturnValue
End Function

'------------------------------------------------------------------------------
' FUNCTION ISREQUIRED(SMASK,IFIELD)
'------------------------------------------------------------------------------
Function IsRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
  		sReturnValue = " checked=""checked"""
	Else
	  	sReturnValue = ""
	End If

	IsRequired = sReturnValue
End Function

'------------------------------------------------------------------------------
' FUNCTION FNGETMASK()
'------------------------------------------------------------------------------
Function fnGetMask()

	sReturnValue = "0000000000"
	sMask = ""

	For iList = 1 to 10
		sParm = "0"
		If request.form("maskshow"&iList) <> "" Then
  			sParm = "1"
		End If
		If request.form("maskrequired"&iList) <> "" Then
		  	sParm = "2"
		End If

		sMask = sMask & sParm
		
	Next 

	If Len(sMask) = 10 Then
  		sReturnValue = sMask
	End If

	fnGetMask = sReturnValue

End Function

'------------------------------------------------------------------------------
' FUNCTION FNGETISSUEMASK()
'------------------------------------------------------------------------------
Function fnGetIssueMask()

	sReturnValue = "0000000000"
	sMask = ""

	For iList = 1 to 10
		sParm = "0"
		If request.form("issueshow"&iList) <> "" Then
			sParm = "1"
		End If
		If request.form("issuerequired"&iList) <> "" Then
			sParm = "2"
		End If

		sMask = sMask & sParm
		
	Next 

	If Len(sMask) = 10 Then
		sReturnValue = sMask
	End If

	fnGetIssueMask = sReturnValue

End Function

'------------------------------------------------------------------------------
'Sub subEditName(iFormID)

'	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id='" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then

'		response.write "<table>"
'		response.write "<tr><td><b>Form Name</b><br>Change/update the form name below.</td></tr>"
'		response.write "<tr><td><textarea class=""PROMPT"" name=""Name"">" & oForm("action_form_name") & "</textarea></td></tr>"
'		response.write "<tr><td><input type=submit class=""button"" value=""Save Changes""></td></tr>"
'		response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub


'------------------------------------------------------------------------------
' SUB SUBEDITENOTE(IFORMID)
'------------------------------------------------------------------------------
'Sub subEditEnote(iFormID)

'	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id='" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then

'		response.write "<table>"
'		response.write "<tr><td><b>Emergency Text</b><br>Change/update the form emergency text below.</td></tr>"
'		response.write "<tr><td><textarea class=""PROMPT"" name=""enote"">" & oForm("action_form_emergency_text") & "</textarea></td></tr>"
'		response.write "<tr><td><input type=submit class=""button"" value=""Save Changes""></td></tr>"
'		response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub


'------------------------------------------------------------------------------
' SUB SUBINTRONAME(IFORMID)
'------------------------------------------------------------------------------
'Sub subIntroName(iFormID)

'	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id='" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then

'		response.write "<table>"
'		response.write "<tr><td><b>Intro Text</b><br>Change/update the intro text below.</td></tr>"
'		response.write "<tr><td><textarea class=""PROMPT"" name=""INTRO"">" & oForm("action_form_description") & "</textarea></td></tr>"
'		response.write "<tr><td><input type=submit class=""button"" value=""Save Changes""></td></tr>"
'		response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub


'------------------------------------------------------------------------------
' SUB SUBFOOTERNAME(IFORMID)
'------------------------------------------------------------------------------
'Sub subFooterName(iFormID)

'	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id='" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then

'		response.write "<table>"
'		response.write "<tr><td><b>Footer Text</b><br>Change/update the footer text below.</td></tr>"
'		response.write "<tr><td><textarea class=""PROMPT"" name=""FOOTER"">" & oForm("action_form_footer") & "</textarea></td></tr>"
'		response.write "<tr><td><input type=submit class=""button"" value=""Save Changes""></td></tr>"
'		response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub

'------------------------------------------------------------------------------
'Sub subContactOptions(iFormID)

'	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id='" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then
		' GET CONTACT MASK
'		sMASK=trim(oForm("action_form_contact_mask"))
'		arrNameList = Array("","First Name", "Last Name", "Business Name", "Email", "Daytime Phone", "Fax", "Street", "City", "State","Zip")
'		response.write "<table>"
'		response.write "  <tr><td colspan=""2""><b>Contact Options</b><br>Change/update the contact options below.</td></tr>"

		' ENUMERATE LIST FOR NAMES AND VALUES
'		For iList = 1 to 10
'   			response.write "  <tr>" & vbcrlf
'      response.write "      <td align=""right""><strong>" & arrNameList(iList) & "</strong></td>" & vbcrlf
'      response.write "      <td><input " & IsDisplay(sMASK,iList) & " name=""maskshow" & iList & """ type=""checkbox"" onClick=""if(document.frmSaveField.maskshow" & iList &".checked==false){document.frmSaveField.maskrequired" & iList &".checked=false};""> Show on Form."
'      response.write           "<input name=""maskrequired" & iList & """ " & IsRequired(sMASK,iList) & " type=""checkbox"" onClick=""document.frmSaveField.maskshow" & iList &".checked=true;"">Is Required." & vbcrlf
'      response.write "      </td>" & vbcrlf
'      response.write "  </tr>" & vbcrlf
'		Next 
		
'		response.write "  <tr><td align=""right""><input type=""submit"" class=""buton"" value=""Save Changes"" /></td></tr>" & vbcrlf
'		response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub


'------------------------------------------------------------------------------
' SUB SUBISSUEOPTIONS(IFORMID)
'------------------------------------------------------------------------------
'Sub subIssueOptions(iFormID)

'	sSQL = "SELECT * FROM egov_action_request_forms WHERE action_form_id='" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then
		' GET CONTACT MASK
'		sIssueMask = oForm("action_form_issue_mask")
'		iStreetNumberInputType = oForm("issuestreetnumberinputtype")
'		iStreetAddressInputType = oForm("issuestreetaddressinputtype")

'		If sIssueMask = "" or IsNull(sIssueMask) Then
'			sIssueMask = "121111"
'		End If

'		arrNameList = Array("","Number", "Street", "City", "State", "Zip", "Additional Information")
'		response.write "<table>"
'		response.write "<tr><td colspan=2><b>Issue/problem Options</b><br>Change/update the issue/problem options below.</td></tr>"

		' ENUMERATE LIST FOR NAMES AND VALUES
'		For iList = 1 to UBOUND(arrNameList)
'			response.write "<tr><td align=right><b>" & arrNameList(iList) & "<b></td><TD> <input " & IsDisplay(sIssueMask,iList) & " name=issueshow" & iList & "  type=checkbox onClick=""if(document.frmSaveField.issueshow" & iList &".checked==false){document.frmSaveField.issuerequired" & iList &".checked=false};"">Show on Form. <input name=issuerequired" & iList & " " & IsRequired(sIssueMask,iList) & " type=checkbox onClick=""document.frmSaveField.issueshow" & iList &".checked=true;"">Is Required.</td>"
'			response.write "<td>"

			' DISPLAY INPUT OPTION STREET NUMBER
'			If iList = 1 Then
				
'				Select Case iStreetNumberInputType

'					Case "1"
'						sSelectText = "SELECTED"
'						sSelectSelect = ""
'						sSelectBoth = ""

'					Case "2"
'						sSelectText = ""
'						sSelectSelect = "SELECTED"
'						sSelectBoth = ""

'					Case "3"
'						sSelectText = ""
'						sSelectSelect = ""
'						sSelectBoth = "SELECTED"
			
'					Case Else
'						sSelectText = "SELECTED"
'						sSelectSelect = ""
'						sSelectBoth = ""

'				End Select
				
				
'				response.write "<select name=selNumberInputType>"
'				response.write "<option value=1 " & sSelectText & ">TEXT"
'				response.write "<option value=2 " & sSelectSelect & ">SELECT"
'				response.write "<option value=3 " & sSelectBoth & ">BOTH"
'				response.write "</select>"
'			End If

			' DISPLAY INPUT OPTION STREET ADDRESS
'			If iList = 2 Then

'					Select Case iStreetAddressInputType

'					Case "1"
'						sSelectText = "SELECTED"
'						sSelectSelect = ""
'						sSelectBoth = ""

'					Case "2"
'						sSelectText = ""
'						sSelectSelect = "SELECTED"
'						sSelectBoth = ""

'					Case "3"
'						sSelectText = ""
'						sSelectSelect = ""
'						sSelectBoth = "SELECTED"
			
'					Case Else
'						sSelectText = "SELECTED"
'						sSelectSelect = ""
'						sSelectBoth = ""

'				End Select
					
	'			response.write "<select name=selAddressInputType>"
	'			response.write "<option value=1 " & sSelectText & ">TEXT"
	'			response.write "<option value=2 " & sSelectSelect & ">SELECT"
	'			response.write "<option value=3 " & sSelectBoth & ">BOTH"
	'			response.write "</select>"

	'		End If

	'		response.write "</td>"
	'		response.write "</tr>"
	'	Next 
		
	'	response.write "<tr><td align=right><input type=submit class=""button"" value=""Save Changes""></td></tr>"
	'	response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub

'------------------------------------------------------------------------------
' SUB SUBEDITISSUENAME(IFORMID)
'------------------------------------------------------------------------------
'Sub subEditIssueName(iFormID)

'	sSQL = "SELECT issuelocationname, issuelocationdesc, issuequestion "
'	sSQL = sSQL & " FROM egov_action_request_forms "
'	sSQL = sSQL & " WHERE action_form_id = '" & iFormID & "'"

'	Set oForm = Server.CreateObject("ADODB.Recordset")
'	oForm.Open sSQL, Application("DSN") , 3, 1
	
'	If NOT oForm.EOF Then
'		sIssueName = oForm("issuelocationname")
'		If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
'			 sIssueName = "Issue/Problem Location:"
'		End If

'		sIssueDesc = oForm("issuelocationdesc")
'		If IsNull(sIssueDesc) Then
'			sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". Provide any additional information on problem location in the box below."
'		End If

'        sIssueQues = oForm("issuequestion")
'		if trim(sIssueQues) = "" or IsNull(sIssueQues) then
'		   sIssueQues = "Provide any additional information on problem location in the box below."
'        end if

'		response.write "<table>"
'		response.write "  <tr><td><b>Name</b><br>Change/update the issue location name below (Maximum 512 characters).</td></tr>"
'		response.write "  <tr><td><textarea class=""PROMPT"" name=""IssueName"" onkeyup=""CheckMaxLength(this, 512);"">" & sIssueName & "</textarea></td></tr>"
'		response.write "  <tr><td>&nbsp;</td></tr>"

'		response.write "  <tr><td><b>Description</b><br>Change/update the issue location description below.</td></tr>"
'		response.write "  <tr><td><textarea class=""PROMPT"" name=""IssueDesc"" style=""height: 200px;"">" & sIssueDesc & "</textarea></td></tr>"
'		response.write "  <tr><td>&nbsp;</td></tr>"

'		response.write "  <tr><td><b>Question</b><br>Change/update the issue location question below (Maximum 512 characters).</td></tr>"
'		response.write "  <tr><td><input type=""type"" name=""IssueQues"" size=""72"" maxlength=""512"" class=""PROMPT"" value=""" & sIssueQues & """></td></tr>"
'		response.write "  <tr><td>&nbsp;</td></tr>"

'		response.write "  <tr><td><input type=submit class=""button"" value=""Save Changes""></td></tr>"
'		response.write "</table>"

'	End If

'	Set oForm = Nothing 

'End Sub
%>
