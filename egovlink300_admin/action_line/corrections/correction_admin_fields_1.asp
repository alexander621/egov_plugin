<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">


<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->


<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_REQUEST_FORM.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/5/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0	02/05/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' INITIALIZE AND DECLARE VARIABLES
Dim sError
sLevel = "../../" ' OVERRIDE OF VALUE FROM COMMON.ASP

' SET TIMEZONE INFORMATION INTO SESSION
Session("iUserOffset") = request.cookies("tz")
%>



<html>

<head>

  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="Javascript" src="../../scripts/modules.js"></script>

  <script language="Javascript" > 
  <!--

	//Set timezone in cookie to retrieve later
	var d=new Date();
	if (d.getTimezoneOffset)
	{
		var iMinutes = d.getTimezoneOffset();
		document.cookie = "tz=" + iMinutes;
	}

  //-->
  </script>

  <STYLE>
		div.correctionsbox {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
		div.correctionsboxnotfound  {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
		td.correctionslabel {font-weight:bold;}
		th.corrections {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
		th.correctionsinternal{background-color:#e0e0e0;font-size:12px;padding:5px;color:#000000; }
		input.correctionstextbox {border: solid 1px #336699;width:400px;}
		textarea.correctionstextarea {border: solid 1px #336699;width:600px;height:100px;}
		.savemsg {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
  </STYLE>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >


<% ShowHeader sLevel %>


<!--#Include file="../../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		

		<h3>Edit Request Form</h3>
		<img align=absmiddle src="../../../admin/images/arrow_2back.gif"> <a href="../action_respond.asp?control=<%=request("irequestid")%>">Return to Request</a> 

		<%
			' DISPLAY TO USER THAT VALUES WERE SAVED
			If request("r") = "save" Then 
				response.write "<P><span class=""savemsg"">Saved " & Now() & ".</span></P>"
			End If
			

			' ADMINISTRATIVE FIELDS
			If UserHasPermission( Session("UserId"), "internalfields" ) Then
				
				If fnAdminSubmitted(request("irequestid")) Then
					
					' EDIT ADMIN QUESTIONS
					Call subDisplayAdminQuestions(request("irequestid"),1)

				Else
				
					' ADD ADMIN QUESTIONS
					response.write "<form name=""adminfields"" action=""correction_admin_fields_cgi.asp"" method=""POST"">"
					response.write "<div class=shadow >"

					response.write "<table  class=tablelist cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">"
					response.write "<tr><th class=corrections  align=left colspan=2>&nbsp;Internal Use Only - Administrative Fields</th></tr>"
					
					' DISPLAY INSTRUCTIONS
					response.write "<tr><td colspan=2><P class=instructions>Please update the administrative fields and press <b>Save</b> when finished making changes.</p></td></tr>"
					
					' SAVE AND CANCEL ROW
					response.write "<tr><td class=correctionslabel align=""left"" colspan=2 ><input type=submit value=""Save"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';""></td></tr>"
					response.write "<tr><td>&nbsp;</td></tr>"

					' DISPLAY ADMIN QUESTIONS - NEW
					Call subDisplayAdminQuestionsNew(request("irequestid"))

					response.write "<tr><td>&nbsp;</td></tr>"

					' SAVE AND CANCEL ROW
					response.write "<tr><td class=correctionslabel align=""left"" colspan=2 ><input type=submit value=""Save"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';""></td></tr>"
					response.write "<tr><td>&nbsp;</td></tr>"

					response.write "</table>"
					response.write "</div>"
					response.write "<input type=hidden name=""formtype"" value=""adminfields"">"
					response.write "<input type=hidden name=""irequestid"" value=""" & request("irequestid") & """>"
					response.write "<input type=hidden value=""" & request("status") & """ name=""status"">"
					response.write "</form>"

				End If
			End If
		%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../../admin_footer.asp"-->  

</body>
</html>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETQUESTIONVALUE(IFIELDID)
'------------------------------------------------------------------------------------------------------------
Function GetQuestionValue(ifieldid)

	sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE submitted_request_field_id='" & ifieldid & "'"
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1

	If NOT oQuestions.EOF Then
		GetQuestionValue = oQuestions("submitted_request_field_response")
	End If

	Set oQuestions = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION ISQUESTIONVALUEMATCH(IFIELDID,SVALUE,STRUEVALUE)
'------------------------------------------------------------------------------------------------------------
Function IsQuestionValueMatch(ifieldid,sValue,sTrueValue)

	sReturnValue = ""

	sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE submitted_request_field_id='" & ifieldid & "' AND submitted_request_field_response = '" & sValue & "'"
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1

	If NOT oQuestions.EOF Then
		If TRIM(oQuestions("submitted_request_field_response")) = TRIM(sValue) Then
			sReturnValue = sTrueValue 
		End If
	End If

	Set oQuestions = Nothing

	IsQuestionValueMatch = sReturnValue 

End Function


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYADMINQUESTIONS(irequestID)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayAdminQuestionsNew(irequestID)

	sSQL = "SELECT * FROM  egov_actionline_requests INNER JOIN egov_action_form_questions ON egov_actionline_requests.category_id=egov_action_form_questions.formid WHERE action_autoid='" & irequestID & "' AND egov_action_form_questions.isinternalonly = 1  ORDER BY sequence"
	
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oQuestions.EOF Then
	
	
	Do While NOT oQuestions.EOF 
		
		' ENUMERATE QUESTIONS
		iQuestionCount = iQuestionCount + 1
		
		' DETERMINE IF REQUIRED
		sIsrequired = oQuestions("isrequired")
		If sIsrequired = True Then
			sIsrequired = " <font color=red>*</font> "
		Else
			sIsrequired = ""
		End If

		' TRACKING CURRENT FORM CONFIGURATION FOR EDITTING LATER
		response.write "<input value=""" & oQuestions("fieldtype") & """ name=""fieldtype"" type=hidden>"
		response.write "<input value=""" & oQuestions("answerlist") & """ name=""answerlist"" type=hidden>"
		response.write "<input value=""" & oQuestions("isrequired") & """ name=""isrequired"" type=hidden>"
		response.write "<input value=""" & oQuestions("sequence") & """ name=""sequence"" type=hidden>"
		response.write "<input value=""" & oQuestions("pdfformname") & """ name=""pdfformname"" type=hidden>"
		

		Select Case oQuestions("fieldtype")

			Case "2"
			' BUILD RADIO QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input value=""" & arrAnswers(alist) & """ name=fmquestion" & iQuestionCount & " class=formradio type=radio>" & arrAnswers(alist) & "</td></tr>"
			Next

			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "4"
			' BUILD SELECT QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			response.write "<tr><td><select class=formselect name=fmquestion" & iQuestionCount & " >"
			For alist = 0 to ubound(arrAnswers)
				response.write "<option value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>" 
			Next
			response.write "</select></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "6"
			' BUILD CHECKBOX QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input value=""" & arrAnswers(alist) & """ name=fmquestion" & iQuestionCount & " class=formcheckbox type=checkbox>" & arrAnswers(alist) & "</td></tr>"
			Next

			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "8"
			' BUILD TEXT QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><input name=fmquestion" & iQuestionCount & " value="""" type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "10"
			' BUILD TEXTAREA QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><textarea name=fmquestion" & iQuestionCount & " class=""formtextarea"" ></textarea></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case Else

		End Select 

		oQuestions.MoveNext
	Loop

	Else

		response.write "<tr><td><p style=""color:red;"">There are no administrative fields assigned to this form.  Administrative fields can be added by users with the <b>Form Creator</b> security feature enabled.</form></p></td></tr>"

	End If

	Set oQuestions = Nothing 

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYQUESTIONS(IREQUESTID,BLNISADMIN)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayAdminQuestions(irequestid,blnIsAdmin)

	' INTERNAL ONLY FIELDS OR REGULAR FIELDS
	If blnIsAdmin Then 
		sSQL = "SELECT * FROM egov_submitted_request_fields WHERE submitted_request_id='" & irequestid & "' AND (submitted_request_field_isinternal = 1)  ORDER BY submitted_request_field_sequence"
	Else
		sSQL = "SELECT * FROM egov_submitted_request_fields WHERE submitted_request_id='" & irequestid & "' AND (submitted_request_field_isinternal = 0 OR submitted_request_field_isinternal IS NULL)  ORDER BY submitted_request_field_sequence"
	End If


	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oQuestions.EOF Then
	
		response.write "<form name=""frmAdminFieldsEdit"" action=""correction_admin_fields_cgi.asp"" method=""POST"">"
		response.write "<div class=shadow >"
				
		response.write "<table  class=tablelist cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">"
		response.write "<tr><th class=corrections  align=left colspan=2>&nbsp;Internal Use Only - Administrative Fields</th></tr>"
		
		
		' DISPLAY INSTRUCTIONS
		response.write "<tr><td colspan=2><P class=instructions>Please update the administrative fields and press <b>Save</b> when finished making changes.</p></td></tr>"
		
		' SAVE AND CANCEL ROW
		response.write "<tr><td class=correctionslabel align=""left""><input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';""></td><td align=""right""><input type=submit value=""Save"">&nbsp;&nbsp;</td></tr>"
		response.write "<tr><td>&nbsp;</td></tr>"
	
	Do While NOT oQuestions.EOF 
		
		' ENUMERATE QUESTIONS
		iQuestionCount = iQuestionCount + 1
		
		' DETERMINE IF REQUIRED
		sIsrequired = oQuestions("submitted_request_field_isrequired")
		If sIsrequired = True Then
			sIsrequired = " <font color=red>*</font> "
		Else
			sIsrequired = ""
		End If
		
		response.write "<input value=""" & oQuestions("submitted_request_field_pdf_name") & """ name=""pdfformname"" type=hidden>"
		
		Select Case oQuestions("submitted_request_field_type_id")

			Case "2"
			' BUILD RADIO QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired & oQuestions("submitted_request_field_prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),arrAnswers(alist)," CHECKED ") & " value=""" & arrAnswers(alist) & """ name=frmanswer" & oQuestions("submitted_request_field_id") & " class=formradio type=radio>" & arrAnswers(alist) & "</td></tr>"
			Next

			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "4"
			' BUILD SELECT QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))
			
			response.write "<tr><td><select class=formselect name=frmanswer" &  oQuestions("submitted_request_field_id") & " >"
			For alist = 0 to ubound(arrAnswers)
				response.write "<option " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),arrAnswers(alist)," SELECTED ") & " value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>" 
			Next
			response.write "</select></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "6"
			' BUILD CHECKBOX QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),arrAnswers(alist)," CHECKED ") & " value=""" & arrAnswers(alist) & """ name=frmanswer" & oQuestions("submitted_request_field_id") & " class=formcheckbox type=checkbox>" & arrAnswers(alist) & "</td></tr>"
			Next

			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "8"
			' BUILD TEXT QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">"
			End If

			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>"
			response.write "<tr><td><input name=frmanswer" & oQuestions("submitted_request_field_id") & " value=""" & GetQuestionValue( oQuestions("submitted_request_field_id")) & """ type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "10"
			' BUILD TEXTAREA QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">"
			End If

			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>"
			response.write "<tr><td><textarea name=frmanswer" &  oQuestions("submitted_request_field_id") & " class=""formtextarea"" >" & GetQuestionValue( oQuestions("submitted_request_field_id")) & "</textarea></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case Else

		End Select 

		oQuestions.MoveNext
	Loop


		' DISPLAY SAVE AND CANCEL BUTTONS
		response.write "<tr><td>&nbsp;</td></tr>"
		response.write "<tr><td class=correctionslabel align=""left""><input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';""></td><td align=""right""><input type=submit value=""Save"">&nbsp;&nbsp;</td></tr>"

		response.write "</table>"
		response.write "</div>"
		response.write "<input type=hidden name=""formtype"" value=""nonblob"">"
		response.write "<input type=hidden name=""irequestid"" value=""" & request("irequestid") & """>"
		response.write "<input type=hidden value=""" & request("status") & """ name=""status"">"
		response.write "</form>"

	End If

	Set oQuestions = Nothing 

End Sub

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNADMINSUBMITTED(IREQUESTID)
'------------------------------------------------------------------------------------------------------------
Function fnAdminSubmitted(irequestid)

	blnReturnValue = False
	
	sSQL = "SELECT * FROM egov_submitted_request_fields WHERE submitted_request_id='" & irequestid & "' AND (submitted_request_field_isinternal = 1)  ORDER BY submitted_request_field_sequence"

	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oQuestions.EOF Then
		blnReturnValue = True
	End If

	Set oQUestions = Nothing

	fnAdminSubmitted = blnReturnValue 


End Function
%>


