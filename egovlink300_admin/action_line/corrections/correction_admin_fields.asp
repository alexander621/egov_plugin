<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<%
  lcl_hidden = "hidden"  'Show/Hide field: HIDDEN = Hide, TEXT = Show
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: CORRECTION_REQUEST_FORM.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/5/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0	02/05/07	 JOHN STULLENBERGER - INITIAL VERSION
' 2.0 11/07/07  David Boyer - Fixed bug with radio/checkboxes that when initially inserting records these fields do not pass data
'                             and therefore not set up correctly in the database and when viewed in system they appeared to be mis-aligned.
' 2.1  08/20/08  David Boyer - Added javascript field length check to textarea fields.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Initialize and declare variables
 Dim sError
 sLevel = "../../"  'Override of value from common.asp

'Set timezone information into session
 session("iUserOffset") = request.cookies("tz")

'Set the field lengths for the custom/internal fields
 lcl_text_field_length     = 1024
 lcl_textarea_field_length = 4000
%>
<html>
<head>
  <title>E-GovLink Administration Consule {Edit Administrative Fields}</title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="javascript" src="../../scripts/modules.js"></script>
 	<script language="javascript" src="../../scripts/textareamaxlength.js"></script>

<script language="javascript" > 
<!--

	//Set timezone in cookie to retrieve later
	var d=new Date();
	if(d.getTimezoneOffset) {
 	 	var iMinutes = d.getTimezoneOffset();
  		document.cookie = "tz=" + iMinutes;
	}
//-->
function validateCheckbox(p_field) {
  //get the total options
  var lcl_total_options = document.getElementById("total_options_"+p_field).innerHTML;
  var lcl_total_checked = 0;

  for (i = 1; i <= lcl_total_options-1; i++) {
       if(document.getElementById(p_field+"_"+i).checked==true) {
          lcl_total_checked = lcl_total_checked + 1;
       }
  }
  if(lcl_total_checked == 0) {
     document.getElementById(p_field+"_"+lcl_total_options).checked=true;
  }else{
     document.getElementById(p_field+"_"+lcl_total_options).checked=false;
  }
}
</script>

<style>
  div.correctionsbox           {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
  div.correctionsboxnotfound   {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
  td.correctionslabel          {font-weight:bold;}
  th.corrections               {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
  th.correctionsinternal       {background-color:#e0e0e0;font-size:12px;padding:5px;color:#000000; }
  input.correctionstextbox     {border: solid 1px #336699;width:400px;}
  textarea.correctionstextarea {border: solid 1px #336699;width:600px;height:100px;}
  .savemsg                     {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
</style>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();">
<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
  <div id="centercontent">
		  <h3>Edit Request Form</h3>
  		<!-- <img align="absmiddle" src="../../../admin/images/arrow_2back.gif"> <a href="../action_respond.asp?control=<%'request("irequestid")%>">Return to Request</a> -->
    <input type="button" name="returnButton" id="returnButton" class="button" value="Return to Request" onclick="location.href='<%=sLevel%>/action_line/action_respond.asp?control=<%=request("irequestid")%>';" />
<%
	 	'DISPLAY TO USER THAT VALUES WERE SAVED
  		If request("r") = "save" Then 
    			response.write "<p><span class=""savemsg"">Saved " & Now() & ".</span></p>"
 			End If

			'ADMINISTRATIVE FIELDS
 			If UserHasPermission( Session("UserId"), "internalfields" ) Then
   				If fnAdminSubmitted(request("irequestid")) Then
    					'EDIT ADMIN QUESTIONS
  		   			Call subDisplayAdminQuestions(request("irequestid"),1)
    			Else

				    	'ADD ADMIN QUESTIONS
     					response.write "<form name=""adminfields"" action=""correction_admin_fields_cgi.asp"" method=""POST"">" & vbcrlf
     					response.write "  <input type=""" & lcl_hidden & """ name=""formtype"" value=""adminfields"">" & vbcrlf
     					response.write "  <input type=""" & lcl_hidden & """ name=""status"" value="""     & request("status")     & """>" & vbcrlf
     	    response.write "  <input type=""" & lcl_hidden & """ name=""substatus"" value="""  & request("substatus")  & """>" & vbcrlf
     					response.write "  <input type=""" & lcl_hidden & """ name=""irequestid"" value=""" & request("irequestid") & """>" & vbcrlf
										response.write "<div class=""shadow"">" & vbcrlf
										response.write "<table class=""tablelist"" cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">" & vbcrlf
										response.write "  <tr><th class=""corrections"" align=""left"" colspan=""2"">&nbsp;Internal Use Only - Administrative Fields</th></tr>" & vbcrlf
					
									'DISPLAY INSTRUCTIONS
										response.write "  <tr><td colspan=""2""><p class=""instructions"">Please update the administrative fields and press <b>Save</b> when finished making changes.</p></td></tr>" & vbcrlf
					
									'Display Buttons
										response.write "  <tr>" & vbcrlf
          response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
                                    displayButtons request("irequestid")
          response.write "      </td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
										response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

   					 'DISPLAY ADMIN QUESTIONS - NEW
     					Call subDisplayAdminQuestionsNew(request("irequestid"))

					     response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

    					'Display Buttons
										response.write "  <tr>" & vbcrlf
          response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
                                    displayButtons request("irequestid")
          response.write "      </td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
     					response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

     					response.write "</table>" & vbcrlf
     					response.write "</div>" & vbcrlf
     					response.write "</form>" & vbcrlf
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
'----------------------------------------------------------------------------------------
Function GetQuestionValue(ifieldid)
 	sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE submitted_request_field_id='" & ifieldid & "'"
	 Set oQuestions = Server.CreateObject("ADODB.Recordset")
	 oQuestions.Open sSQL, Application("DSN"), 3, 1

 	If NOT oQuestions.EOF Then
 	  	GetQuestionValue = oQuestions("submitted_request_field_response")
  End If

 	Set oQuestions = Nothing
End Function

'----------------------------------------------------------------------------------------
Function IsQuestionValueMatch(ifieldid,sValue,sTrueValue)
 	sReturnValue = ""

	 sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_LIKLIKEubmitted_request_field_responses "
  sSQL = sSQL & " WHERE submitted_request_field_id='" & ifieldid & "' "
  sSQL = sSQL & " AND submitted_request_field_response LIKE '" & sValue & "'"
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

'----------------------------------------------------------------------------------------
Sub subDisplayAdminQuestionsNew(irequestID)
 	sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_actionline_requests "
  sSQL = sSQL & " INNER JOIN egov_action_form_questions ON egov_actionline_requests.category_id=egov_action_form_questions.formid "
  sSQL = sSQL & " WHERE action_autoid='" & irequestID & "' "
  sSQL = sSQL & " AND egov_action_form_questions.isinternalonly = 1 "
  sSQL = sSQL & " ORDER BY sequence "

	 Set oQuestions = Server.CreateObject("ADODB.Recordset")
 	oQuestions.Open sSQL, Application("DSN"), 3, 1

 	If NOT oQuestions.EOF Then
     Do While NOT oQuestions.EOF 
     		'ENUMERATE QUESTIONS
      		iQuestionCount = iQuestionCount + 1

     		'DETERMINE IF REQUIRED
      		sIsrequired = oQuestions("isrequired")
      		If sIsrequired = True Then
         		sIsrequired = " <font color=""red"">*</font> "
      		Else
        			sIsrequired = ""
      		End If

     		'TRACKING CURRENT FORM CONFIGURATION FOR EDITTING LATER
      		response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("fieldtype") & """ name=""fieldtype"">" & vbcrlf
    				response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("answerlist") & """ name=""answerlist" & iQuestionCount & """>" & vbcrlf
    				response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("isrequired") & """ name=""isrequired"">" & vbcrlf
    				response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("sequence") & """ name=""sequence"">" & vbcrlf
    				response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("pdfformname") & """ name=""pdfformname"">" & vbcrlf
		
		      select Case oQuestions("fieldtype")

    		  		Case "2"
           	 		'BUILD RADIO QUESTION
            	 		If sIsrequired <> "" Then
               				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">" & vbcrlf
             			End If
             			response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>" & vbcrlf
             			response.write "<tr><td class=""question"">" & sIsrequired & oQuestions("prompt")& "</td></tr>" & vbcrlf
          						arrAnswers = split(oQuestions("answerlist"),chr(10))

          						For alist = 0 to ubound(arrAnswers)
          			   				response.write "<tr><td><input value=""" & arrAnswers(alist) & """ name=""fmquestion" & iQuestionCount & """ class=""formradio"" type=""radio"">" & arrAnswers(alist) & "</td></tr>" & vbcrlf
          						Next

                response.write "<tr style=""display: none"">" & vbcrlf
                response.write "    <td><input type=""radio"" name=""fmquestion" & iQuestionCount & """ value=""default_novalue"" CHECKED></td>" & vbcrlf
                response.write "</tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

         	Case "4"
          					'BUILD SELECT QUESTION
          						If sIsrequired <> "" Then
          			  				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">" & vbcrlf
          						End If
          						response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>" & vbcrlf
          						response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>" & vbcrlf
          						arrAnswers = split(oQuestions("answerlist"),chr(10))

          						response.write "<tr><td><select class=""formselect"" name=""fmquestion" & iQuestionCount & """>" & vbcrlf
          						For alist = 0 to ubound(arrAnswers)
          			   				response.write "<option value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>"  & vbcrlf
          						Next
          						response.write "</select></td></tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

          Case "6"
          					'BUILD CHECKBOX QUESTION
          						If sIsrequired <> "" Then
          			  				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">" & vbcrlf
          						End If
          						response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>" & vbcrlf
          						response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>" & vbcrlf
          						arrAnswers = split(oQuestions("answerlist"),chr(10))

                i = 0
          						For alist = 0 to ubound(arrAnswers)
                    i = i + 1
          			   				response.write "<tr><td><input value=""" & arrAnswers(alist) & """ name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & "_" & i & """ class=""formcheckbox"" type=""checkbox"" onclick=""validateCheckbox('fmquestion" & iQuestionCount & "')"">" & arrAnswers(alist) & "</td></tr>" & vbcrlf
          						Next
                i = i + 1
                response.write "<tr style=""display: none"">" & vbcrlf
                response.write "    <td>" & vbcrlf
                response.write "        <input type=""checkbox"" name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & "_" & i & """ value=""default_novalue"" CHECKED onclick=""validateCheckbox('fmquestion" & iQuestionCount & "')"">" & vbcrlf
                response.write "        <span id=""total_options_fmquestion" & iQuestionCount & """>" & i & "</span>" & vbcrlf
                response.write "    </td></tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

   			    Case "8"
          					'BUILD TEXT QUESTION
          						If sIsrequired <> "" Then
          			  				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">" & vbcrlf
          						End If
          						response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>" & vbcrlf
          						response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>" & vbcrlf
          						response.write "<tr><td><input name=""fmquestion" & iQuestionCount & """ value="""" type=""text"" style=""width:300px;"" maxlength=""" & lcl_text_field_length & """></td></tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

 			      Case "10"
          					'BUILD TEXTAREA QUESTION
          						If sIsrequired <> "" Then
          			  				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">" & vbcrlf
          						End If
          						response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>" & vbcrlf
          						response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("prompt")& "</td></tr>" & vbcrlf
          						response.write "<tr><td><textarea name=""fmquestion" & iQuestionCount & """ class=""formtextarea"" maxlength=""" & lcl_textarea_field_length & """ onchange=""checkMaxLength();""></textarea></td></tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

   			    Case Else

		      End Select 

		      oQuestions.MoveNext
     Loop

  Else

     response.write "<tr><td><p style=""color:red;"">There are no administrative fields assigned to this form.  Administrative fields can be added by users with the <b>Form Creator</b> security feature enabled.</form></p></td></tr>" & vbcrlf

 	End If

	 Set oQuestions = Nothing 

End Sub

'----------------------------------------------------------------------------------------
Sub subDisplayAdminQuestions(irequestid,blnIsAdmin)
 'INTERNAL ONLY FIELDS OR REGULAR FIELDS
	 If blnIsAdmin Then 
		   sSQL = "SELECT * FROM egov_submitted_request_fields "
   		sSQL = sSQL & " WHERE submitted_request_id='" & irequestid & "' "
   		sSQL = sSQL & " AND (submitted_request_field_isinternal = 1) "
   		sSQL = sSQL & " ORDER BY submitted_request_field_sequence"
 	Else
	   	sSQL = "SELECT * FROM egov_submitted_request_fields "
   		sSQL = sSQL & " WHERE submitted_request_id='" & irequestid & "' "
   		sSQL = sSQL & " AND (submitted_request_field_isinternal = 0 "
   		sSQL = sSQL & " OR submitted_request_field_isinternal IS NULL) "
   		sSQL = sSQL & " ORDER BY submitted_request_field_sequence"
 	End If

 	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	 oQuestions.Open sSQL, Application("DSN"), 3, 1

 	If NOT oQuestions.EOF Then
   		response.write "<form name=""frmAdminFieldsEdit"" action=""correction_admin_fields_cgi.asp"" method=""POST"">" & vbcrlf
   		response.write "<div class=""shadow"">" & vbcrlf
   		response.write "<table class=""tablelist"" cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">" & vbcrlf
   		response.write "<tr><th class=""corrections"" align=""left"" colspan=""2"">&nbsp;Internal Use Only - Administrative Fields</th></tr>" & vbcrlf

  		'DISPLAY INSTRUCTIONS
   		response.write "<tr><td colspan=""2""><P class=""instructions"">Please update the administrative fields and press <b>Save</b> when finished making changes.</p></td></tr>" & vbcrlf

   	'SAVE AND CANCEL ROW
   		'response.write "<tr>" & vbcrlf
     'response.write "    <td class=""correctionslabel"" align=""left"">" & vbcrlf
     'response.write "        <input type=""button"" value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';"">" & vbcrlf
     'response.write "    </td>" & vbcrlf
     'response.write "    <td align=""right"">" & vbcrlf
     'response.write "        <input type=""submit"" value=""Save"">" & vbcrlf
     'response.write "    </td>" & vbcrlf


					response.write "  <tr>" & vbcrlf
     response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
                               displayButtons request("irequestid")
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     'response.write "</tr>" & vbcrlf
   		response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf
	
   	 Do While NOT oQuestions.EOF 
     		'ENUMERATE QUESTIONS
      		iQuestionCount = iQuestionCount + 1

     		'DETERMINE IF REQUIRED
       	sIsrequired = oQuestions("submitted_request_field_isrequired")
      		If sIsrequired = True Then
        			sIsrequired = " <font color=""red"">*</font> "
      		Else
        			sIsrequired = ""
      		End If

      		response.write "<input value=""" & oQuestions("submitted_request_field_pdf_name") & """ name=""pdfformname"" type=""" & lcl_hidden & """>" & vbcrlf

       	Select Case oQuestions("submitted_request_field_type_id")

        		Case "2"
            			'BUILD RADIO QUESTION
             			If sIsrequired <> "" Then
               				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">" & vbcrlf
             			End If
             			response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>" & vbcrlf
          						response.write "<tr><td class=""question"">" & sIsrequired & oQuestions("submitted_request_field_prompt")& "</td></tr>" & vbcrlf
          						arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))

          						For alist = 0 to ubound(arrAnswers)
          			   				response.write "<tr><td><input " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),arrAnswers(alist)," CHECKED ") & " value=""" & arrAnswers(alist) & """ name=""frmanswer" & oQuestions("submitted_request_field_id") & """ class=""formradio"" type=""radio"">" & arrAnswers(alist) & "</td></tr>" & vbcrlf
          						Next

                response.write "<tr style=""display: none"">" & vbcrlf
                response.write "    <td><input type=""radio"" name=""frmanswer" & oQuestions("submitted_request_field_id") & """ value=""default_novalue"" " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),"default_novalue"," CHECKED ") & "></td>" & vbcrlf
                response.write "</tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

        		Case "4"
            			'BUILD SELECT QUESTION
              		If sIsrequired <> "" Then
               				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">" & vbcrlf
             			End If
          						response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>" & vbcrlf
          						response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>" & vbcrlf
          						arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))

          						response.write "<tr><td><select class=""formselect"" name=""frmanswer" &  oQuestions("submitted_request_field_id") & """>" & vbcrlf
          						For alist = 0 to ubound(arrAnswers)
          			   				response.write "<option " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),arrAnswers(alist)," SELECTED ") & " value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>"  & vbcrlf
          						Next
          						response.write "</select></td></tr>" & vbcrlf
          						response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

          Case "6"
            			'BUILD CHECKBOX QUESTION
            		 	If sIsrequired <> "" Then
              				'response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
           			    	response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">" & vbcrlf
	             		End If
           					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>" & vbcrlf
           					response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>" & vbcrlf
           					arrAnswers = split(oQuestions("submitted_request_field_answerlist"),chr(10))

                i = 0
           					For alist = 0 to ubound(arrAnswers)
                    i = i + 1
           		   				response.write "<tr><td><input " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),arrAnswers(alist)," CHECKED ") & " type=""checkbox"" name=""frmanswer" & oQuestions("submitted_request_field_id") & """ id=""frmanswer" & oQuestions("submitted_request_field_id") & "_" & i & """ value=""" & arrAnswers(alist) & """ class=""formcheckbox"" onclick=""validateCheckbox('frmanswer" & oQuestions("submitted_request_field_id") & "')"">" & arrAnswers(alist) & "</td></tr>" & vbcrlf
           					Next
                i = i + 1
                response.write "<tr style=""display: none"">" & vbcrlf
                response.write "    <td>" & vbcrlf
                response.write "        <input " & IsQuestionValueMatch(oQuestions("submitted_request_field_id"),"default_novalue"," CHECKED ") & " type=""checkbox"" name=""frmanswer" & oQuestions("submitted_request_field_id") & """ id=""frmanswer" & oQuestions("submitted_request_field_id") & "_" & i & """ value=""default_novalue"" onclick=""validateCheckbox('frmanswer" & oQuestions("submitted_request_field_id") & "')"">" & vbcrlf
                response.write "        <span id=""total_options_frmanswer" & oQuestions("submitted_request_field_id") & """>" & i & "</span>" & vbcrlf
                response.write "    </td></tr>" & vbcrlf
           					response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

			       Case "8"
			            'BUILD TEXT QUESTION
             			If sIsrequired <> "" Then
               				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">" & vbcrlf
             			End If

              		response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>" & vbcrlf
           					response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>" & vbcrlf
           					response.write "<tr><td><input name=""frmanswer" & oQuestions("submitted_request_field_id") & """ value=""" & GetQuestionValue( oQuestions("submitted_request_field_id")) & """ type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>" & vbcrlf
           					response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

			       Case "10"
             		'BUILD TEXTAREA QUESTION
           					If sIsrequired <> "" Then
           		  				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" &  Left(oQuestions("submitted_request_field_prompt"),75) & "..."">" & vbcrlf
           					End If

           					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("submitted_request_field_prompt") & """>" & vbcrlf
           					response.write "<tr><td class=""question"">" & sIsrequired  & oQuestions("submitted_request_field_prompt")& "</td></tr>" & vbcrlf
           					response.write "<tr><td><textarea name=""frmanswer" &  oQuestions("submitted_request_field_id") & """ class=""formtextarea"" >" & GetQuestionValue( oQuestions("submitted_request_field_id")) & "</textarea></td></tr>" & vbcrlf
           					response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

          Case Else

       	End Select 

      		oQuestions.MoveNext
   	 Loop

  		'DISPLAY SAVE AND CANCEL BUTTONS
   		response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf
 				'response.write "<tr>" & vbcrlf
     'response.write "    <td class=""correctionslabel"" align=""left"">" & vbcrlf
     'response.write "        <input type=""button"" value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';"">" & vbcrlf
     'response.write "    </td>" & vbcrlf
     'response.write "    <td align=""right""><input type=""submit"" value=""Save""></td>" & vbcrlf
     'response.write "</tr>" & vbcrlf
					response.write "  <tr>" & vbcrlf
     response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
                               displayButtons request("irequestid")
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
 				response.write "</table>" & vbcrlf
 				response.write "</div>" & vbcrlf
 				response.write "<input type=""" & lcl_hidden & """ name=""formtype"" value=""nonblob"">" & vbcrlf
 				response.write "<input type=""" & lcl_hidden & """ name=""irequestid"" value=""" & request("irequestid") & """>" & vbcrlf
 				response.write "<input type=""" & lcl_hidden & """ value=""" & request("status") & """ name=""status"">" & vbcrlf
 				response.write "<input type=""" & lcl_hidden & """ value=""" & request("substatus") & """ name=""substatus"">" & vbcrlf
 				response.write "</form>" & vbcrlf

 	End If

  Set oQuestions = Nothing 

End Sub

'------------------------------------------------------------------------------
Function fnAdminSubmitted(irequestid)
 	blnReturnValue = False

 	sSQL = "SELECT * FROM egov_submitted_request_fields "
 	sSQL = sSQL & " WHERE submitted_request_id='" & irequestid & "' "
	 sSQL = sSQL & " AND (submitted_request_field_isinternal = 1) "
 	sSQL = sSQL & " ORDER BY submitted_request_field_sequence"

 	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	 oQuestions.Open sSQL, Application("DSN"), 3, 1

 	If NOT oQuestions.EOF Then
	   	blnReturnValue = True
 	End If

	 Set oQuestions = Nothing

 	fnAdminSubmitted = blnReturnValue 
End Function

'------------------------------------------------------------------------------
sub displayButtons(iRequestID)

  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onClick=""location.href='../action_respond.asp?control=" & iRequestID & "';"" />" & vbcrlf
  response.write "<input type=""submit"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" />" & vbcrlf

end sub
%>
