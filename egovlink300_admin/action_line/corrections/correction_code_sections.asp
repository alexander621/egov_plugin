<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../action_line_global_functions.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_CODE_SECTIONS.ASP
' AUTHOR: DAVID BOYER
' CREATED: 09/04/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0	09/04/07      DAVID BOYER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

 Dim sError
 sLevel = "../../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "action_line_code_sections" ) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Set timezone information into session
 session("iUserOffset") = request.cookies("tz")

 irequestid = request("irequestid")
 sStatus    = request("status")
 sSubStatus = request("substatus")

if request("sAction") = "Save Changes" then

  'Retrieve all of the current code sections to build the "log file"
   sSqlc = "SELECT cs.code_name "
   sSqlc = sSqlc & " FROM egov_actionline_code_sections cs, egov_submitted_request_code_sections scs "
   sSqlc = sSqlc & " WHERE cs.action_code_id = scs.submitted_action_code_id "
   sSqlc = sSqlc & " AND scs.submitted_request_id = " & irequestid
   sSqlc = sSqlc & " ORDER BY upper(cs.code_name), upper(cs.description) "

   Set oCurrent = Server.CreateObject("ADODB.Recordset")
   oCurrent.Open sSqlc, Application("DSN"),3,1

   if not oCurrent.eof then
      lcl_original_values = ""
	  while not oCurrent.eof
        if lcl_original_values = "" then
           lcl_original_values = oCurrent("code_name")
        else
           lcl_original_values = lcl_original_values & ", " & oCurrent("code_name")
		end if

		oCurrent.movenext
	  wend
   end if

  'Used for log file
   lcl_new_values = ""

   for e = 1 to request.form("total_codesections")

      'Loop through each code section to determine if a record already exists for this form
	   sSqle = "SELECT count(submitted_request_code_id) AS lcl_exists "
	   sSqle = sSqle & " FROM egov_submitted_request_code_sections "
	   sSqle = sSqle & " WHERE submitted_action_code_id = " & request("frmanswer_codeid_"&e)
       sSqle = sSqle & " AND submitted_request_id = "       & irequestid

       Set oExists = Server.CreateObject("ADODB.Recordset")
       oExists.Open sSqle, Application("DSN"),3,1

       if oExists("lcl_exists") > 0 then
          lcl_exists = "Y"
       else
          lcl_exists = "N"
	   end if
	   
	   if request.form("frmanswer_"&e) = "yes" then
		  
         'Create the record for this Code Section (in loop) if a record doesn't exist for this form.
		  if lcl_exists = "N" then
  		     sSqlMax = "SELECT IsNull(max(submitted_request_code_id),0) + 1 AS max_code_id "
		     sSqlMax = sSqlMax & " FROM egov_submitted_request_code_sections "

             Set oMaxCode = Server.CreateObject("ADODB.Recordset")
             oMaxCode.Open sSqlMax, Application("DSN"),3,1

             lcl_max_id = oMaxCode("max_code_id")

		     sSqli = "INSERT INTO egov_submitted_request_code_sections (submitted_request_code_id, submitted_request_id, submitted_action_code_id) VALUES ("
             sSqli = sSqli & lcl_max_id                     & ", "
		     sSqli = sSqli & irequestid                     & ", "
		     sSqli = sSqli & request("frmanswer_codeid_"&e) & ") "

		     Set oInsert = Server.CreateObject("ADODB.Recordset")
             oInsert.Open sSqli, Application("DSN"),3,1
          end if

         'Update log file
		  if lcl_new_values = "" then
             lcl_new_values = fnGetFieldPrompt(request("frmanswer_codeid_"&e))
		  else
             lcl_new_values = lcl_new_values & ", " & fnGetFieldPrompt(request("frmanswer_codeid_"&e))
		  end if
	   else
          if lcl_exists = "Y" then
  		     delQuery = "DELETE FROM egov_submitted_request_code_sections "
             delQuery = delQuery & " WHERE submitted_action_code_id = " & request("frmanswer_codeid_"&e)
             delQuery = delQuery & " AND submitted_request_id = "       & irequestid

             Set oDelete = Server.CreateObject("ADODB.Recordset")
             oDelete.Open delQuery, Application("DSN"),3,1
          end if

         'Update log file
		  if lcl_new_values = "" then
             lcl_new_values = ""
		  else
             lcl_new_values = lcl_new_values
		  end if

	   end if
   next

   set oExists  = nothing
   set oMaxCode = nothing
   set oInsert  = nothing
   set oDelete  = nothing
   set oCurrent = nothing

   sLogEntry = ""
   If trim(lcl_original_values) <> trim(lcl_new_values) Then
      sLogEntry = chr(34) & lcl_original_values & chr(34) & " changed to " & chr(34) & lcl_new_values & chr(34)
   End If

  'Record in log the save activity
   if sLogEntry <> "" then
      sLogEntry = "Edit Code Section: "  & sLogEntry
      AddCommentTaskComment sLogEntry, sExternalMsg, sStatus, request("irequestid"), session("userid"), session("orgid"), sSubStatus, "", ""
   end if

   response.redirect "../action_respond.asp?control=" & request("irequestid") & "&r=save&status="&request("status")
end if
%>
<html>
<head>
  <title><%=langBSHome%></title>
  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />
<script language="javascript" src="../../scripts/modules.js"></script>
<script language="javascript"> 
<!--
//Set timezone in cookie to retrieve later
var d=new Date();
if(d.getTimezoneOffset) {
   var iMinutes = d.getTimezoneOffset();
   document.cookie = "tz=" + iMinutes;
}
//-->
</script>
<style>
  div.correctionsbox           {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
  div.correctionsboxnotfound   {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
  td.correctionslabel          {font-weight:bold;}
  th.corrections               {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
  th.correctionsinternal       {background-color:#e0e0e0;font-size:12px;padding:5px;color:#000000; }
  input.correctionstextbox     {border: solid 1px #336699;width:400px;}
  textarea.correctionstextarea {border: solid 1px #336699;width:600px;height:100px;}
  .savemsg                     {font-size:12px;padding:5px;color:#ff0000;font-weight:bold; }
</style>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
  <div id="centercontent">

<h3>Edit Code Sections</h3>
<input type="button" name="returnButton" id="returnButton" value="Return to Request" class="button" onclick="location.href='../action_respond.asp?control=<%=request("irequestid")%>';" />
<%
  'DISPLAY TO USER THAT VALUES WERE SAVED
   if request("r") = "save" then
      response.write "<p><span class=""savemsg"">Saved " & Now() & ".</span></p>"
   end if

  'GET FORM INFORMATION
   subDisplayCodeSections request("irequestid")
%>
  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../../admin_footer.asp"-->  
</body>
</html>
<%
'------------------------------------------------------------------------------
Sub subDisplayCodeSections(irequestid)

 sSQL = " SELECT action_code_id, code_name, description, active_flag, orgid "
	sSQL = sSQL & " FROM egov_actionline_code_sections "
	sSQL = sSQL & " WHERE orgid = " & session("orgid")
	sSQL = sSQL & " AND active_flag = 'Y' "
	sSQL = sSQL & " ORDER BY upper(code_name), upper(description) "

	Set oCodes = Server.CreateObject("ADODB.Recordset")
	oCodes.Open sSQL, Application("DSN"), 3, 1

    If NOT oCodes.EOF Then
       response.write "<form name=""codesections"" action=""correction_code_sections.asp"" method=""POST"">" & vbcrlf
       response.write "<input type=""hidden"" name=""p_action"" value=""modify_codesection"" />" & vbcrlf
       response.write "<div class=""shadow"">" & vbcrlf
       response.write "<table class=""tablelist"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
       response.write "  <tr><th class=""corrections"" align=""left"">&nbsp;Code Sections</th></tr>" & vbcrlf

      'DISPLAY INSTRUCTIONS
       response.write "  <tr><td><p class=""instructions"">Please update the code sections and press <b>Save</b> when finished making changes.</p></td></tr>" & vbcrlf

      'SAVE AND CANCEL ROW
       response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""left"">" & vbcrlf
                                 displayButtons request("irequestid")
   	   response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

       iCodeCount = 0
   	   do while not oCodes.eof

         'ENUMERATE CODE SECTIONS
          iCodeCount = iCodeCount + 1

'       response.write "<input type=""hidden"" name=""ef:fmcode" & iCodeCount & "-select/req"" value=""" &  Left(oCodes("submitted_request_field_prompt"),75) & "..."">"
'       response.write "<input type=""hidden"" name=""fmname" & iQuestionCount & """ value=""" &  oCodes("submitted_request_field_prompt") & """>"

'       response.write "<tr><td class=""question"">" & oCodes("code_name")& "</td></tr>"
'       arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))

'       For alist = 0 to ubound(arrAnswers)
          response.write "  <tr>" & vbcrlf
          response.write "      <td>" & vbcrlf
          response.write "          <input type=""checkbox"" name=""frmanswer_" & iCodeCount & """ value=""yes"" class=""formcheckbox"" " & IsQuestionValueMatch(irequestid, oCodes("action_code_id"),"CHECKED") & " />&nbsp;" & oCodes("code_name") & vbcrlf
          response.write "          <input type=""hidden"" name=""frmanswer_codeid_" & iCodeCount & """ value=""" & oCodes("action_code_id") & """ />" & vbcrlf
          response.write "      </td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
'       Next

'       response.write "<tr><td>&nbsp;</td></tr>"
          oCodes.MoveNext
       loop

      'DISPLAY SAVE AND CANCEL BUTTONS
       response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
       response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""left"">" & vbcrlf
                                 displayButtons request("irequestid")
	      response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "</table>" & vbcrlf
       response.write "</div>" & vbcrlf
       'response.write "<input type=""hidden"" name=""formtype"" value=""code_sections"">"
       response.write "<input type=""hidden"" name=""irequestid"" value=""" & request("irequestid") & """ />" & vbcrlf
  	    response.write "<input type=""hidden"" name=""status"" value=""" & sStatus & """ />" & vbcrlf
  	    response.write "<input type=""hidden"" name=""substatus"" value=""" & sSubStatus & """ />" & vbcrlf
  	    response.write "<input type=""hidden"" name=""total_codesections"" value=""" & iCodeCount & """ />" & vbcrlf
       response.write "</form>" & vbcrlf
    else

       response.write "No Code Sections Available" & vbcrlf
      'FORM INFORMATION NON BLOB NOT FOUND
'       response.write "<p><div class=""correctionsboxnotfound"">Note: The original request question\answer formatting is not available for this request.  The form was submitted on a previous release of the E-Gov Link software.</div></P>"
       response.write "<form name=""frmblob"" action=""correction_request_form_cgi.asp"" method=""POST"">" & vbcrlf
'       response.write "<div class=""shadow"">"
'       response.write "<table class=""tablelist"" cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">"
'       response.write "  <tr><th class=""corrections"" align=""left"" colspan=""2"">&nbsp;Request Form</th></tr>"

	  'DISPLAY INSTRUCTIONS
'       response.write "  <tr><td colspan=""2""><p class=""instructions"">Please update the request form information and press <b>Save</b> when finished making changes.</p></td></tr>"

      'SAVE AND CANCEL ROW
'       response.write "  <tr><td class=""correctionslabel"" align=""left"" colspan=""2""><input type=""submit"" value=""Save"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';""></td></tr>"
'       response.write "  <tr><td>&nbsp;</td></tr>"

'       Call SubGenericEdit(irequestid)

'       response.write "  <tr><td>&nbsp;</td></tr>"
		
      'SAVE AND CANCEL ROW
'       response.write "  <tr><td class=""correctionslabel"" align=""left"" colspan=""2""><input type=""submit"" value=""Save"">&nbsp;&nbsp;<input type=""button"" value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';""></td></tr>"
'       response.write "</table>"
'       response.write "</div>"
'       response.write "<input type=""hidden"" name=""formtype"" value=""blob"">"
       response.write "<input type=""hidden"" name=""irequestid"" value=""" & request("irequestid") & """ />" & vbcrlf
	      response.write "<input type=""hidden"" name=""status"" value=""" & sStatus & """ />" & vbcrlf
	      response.write "<input type=""hidden"" name=""substatusid"" value=""" & sSubStatus & """ />" & vbcrlf
	      response.write "<input type=""hidden"" name=""total_codesections"" value=""" & iCodeCount & """ >" & vbcrlf
       response.write "</form>"
    end if

	set oCodes = nothing

end sub
	
'------------------------------------------------------------------------------
Function GetQuestionValue(ifieldid)

	sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE submitted_request_field_id='" & ifieldid & "'"
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1

	If NOT oQuestions.EOF Then
		GetQuestionValue = oQuestions("submitted_request_field_response")
	End If

	Set oQuestions = Nothing

End Function

'------------------------------------------------------------------------------
Function IsQuestionValueMatch(ifieldid,icodeid,sTrueValue)

	sReturnValue = ""

	sSQL = "SELECT * "
	sSQL = sSQL & " FROM egov_submitted_request_code_sections "
	sSQL = sSQL & " WHERE submitted_request_id = '"   & ifieldid & "' "
	sSQL = sSQL & " AND submitted_action_code_id = '" & icodeid  & "' "

	Set oCodes = Server.CreateObject("ADODB.Recordset")
	oCodes.Open sSQL, Application("DSN"), 3, 1

	If NOT oCodes.EOF Then
'		If TRIM(oCodes("submitted_request_field_response")) = TRIM(sValue) Then
			sReturnValue = sTrueValue 
'		End If
	End If

	Set oCodes = Nothing

	IsQuestionValueMatch = sReturnValue 

End Function

'------------------------------------------------------------------------------
Sub SubGenericEdit(irequestid)

	sSQL = "SELECT comment FROM egov_actionline_requests WHERE action_autoid='" & irequestid & "'"
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oQuestions.EOF Then

		sBlob = oQuestions("comment")

		If Trim(sBlob) <> "" Then
			' CHANGE TO QUESTION\ANSWER FORM

			' SPLIT QUESTIONS INTO ARRAY
			arrQues = split(lcase(sblob),"<p>")

			For iQues = 1 to UBOUND(arrQues)
				
				'SPLIT QUESTIONS AND ANSWERS 
				'QUESTION
				sQues   = LEFT(arrQues(iQues),INSTR(arrQues(iQues),"<br>")-1)
				sAnswer = replace(arrQues(iQues),sQues,"")
				sQues   = replace(lcase(sQues),"<p>","")
				sQues   = replace(lcase(sQues),"<b>","")
				sQues   = replace(lcase(sQues),"</b>","")

				'ANSWER
				sAnswer = replace(lcase(sAnswer),"<br>","")
				sAnswer = replace(lcase(sAnswer),"</p>","")

				'FORM INFORMATION
				response.write "<input name=""question" & iQues & """ type=""hidden"" value=""" & sQues & """>"
				
				'DISPLAY
				response.write "<tr><td class=""correctionslabel"">" & UCASE(sQues) & "</td></tr>"
				response.write "<tr><td><textarea name=""answer" & iQues & """ class=""correctionstextarea"">" & sAnswer & "</textarea></td></tr>"

			Next

		End If

		oQuestions.Close

	End If

	Set oQuestions = Nothing

End Sub

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
Function fnGetFieldPrompt(iField)

	  sReturnValue = "UNKNOWN" 

	  sSQL = "SELECT code_name FROM egov_actionline_code_sections WHERE action_code_id = " & iField
	  Set oFieldName = Server.CreateObject("ADODB.Recordset")
      oFieldName.Open sSQL, Application("DSN") , 1,3
	  
	  If NOT oFieldName.EOF Then
		sReturnValue = oFieldName("code_name") 
		oFieldName.Close
	  End If

	  Set oFieldName = Nothing
	
	  fnGetFieldPrompt = sReturnValue

End Function

'------------------------------------------------------------------------------
sub displayButtons(iRequestID)

  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='../action_respond.asp?control=" & iRequestID & "';"" />" & vbcrlf
  response.write "<input type=""submit"" name=""sAction"" class=""button"" value=""Save Changes"" />" & vbcrlf

end sub
%>