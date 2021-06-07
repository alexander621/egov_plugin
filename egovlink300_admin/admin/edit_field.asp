<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 Dim sPDFFormName
 sLevel     = "../"     ' Override of value from common.asp
 lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

 if not userhaspermission(session("userid"), "form creator" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_pushcontent_communitycalendar = orghasfeature("pushcontent_communitycalendar")

'Check for user permissions
 lcl_userhaspermission_requestmergeforms = userhaspermission(session("userid"), "requestmergeforms")
 lcl_userhaspermission_form_letters      = userhaspermission(session("userid"), "form letters")

'Process form posts and identify fields/questions and form
 iFormID  = request("iformid")
 iOrgID   = request("iorgid")
 iFieldID = request("ifieldid")

 if request.servervariables("REQUEST_METHOD") = "POST" then

  	'Save changes to question
    subSaveQuestion iFieldID

   	response.redirect("manage_form.asp?iformid=" & iFormID )
 end if
%>
<html>
<head>
 	<title>E-GovLink Administration Consule {Forms Management - Edit Field} </title>
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript">

  function saveChanges() {
     document.getElementById("frmSaveField").submit();
  }

</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
 	<% ShowHeader sLevel %>
 	<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "    <table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td valign=""top"">" & vbcrlf
  response.write "           			<p>" & vbcrlf
  response.write "                 <font class=""label"">Forms - Edit Question </font> " & vbcrlf
  response.write "                 <small>[<a class=""edit"" href=""manage_form.asp?iformid=" & iformid & """>Return to Form</a>]</small>" & vbcrlf
  response.write "                 <hr size=""1"" width=""600px;"" style=""text-align:left; color:#000000;"">" & vbcrlf
  response.write "              </p>" & vbcrlf

                                displayButtons "TOP"

  response.write "           			<div class=""group"">" & vbcrlf

		if request.servervariables("REQUEST_METHOD") = "POST" then
			  response.write "<small><strong><font color=""#0000ff"">Question Updated - " & Now() & "</strong></font></small>" & vbcrlf
  end if

  response.write "                <div class=""orgadminboxf"">" & vbcrlf

 'BEGIN: Add question builder information -------------------------------------
  if iFieldType <> "0" then
     response.write "<form name=""frmSaveField"" id=""frmSaveField"" action=""edit_field.asp"" method=""post"">" & vbcrlf
     response.write "  <input type=""" & lcl_hidden & """ name=""iformid"" value=""" & iformid & """ />" & vbcrlf
     response.write "  <input type=""" & lcl_hidden & """ name=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
     response.write "  <input type=""" & lcl_hidden & """ name=""ifieldid"" value=""" & iFieldID & """ />" & vbcrlf
     response.write "  <input type=""" & lcl_hidden & """ name=""task"" value=""ADD"" />" & vbcrlf

 				subCreateEditQuestion iFormID, iFieldID

     response.write "</form>" & vbcrlf
  end if
 'END: Add question builder information ---------------------------------------

  response.write "                </div>" & vbcrlf
  response.write "              </div>" & vbcrlf

                                displayButtons "BOTTOM"

  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  %>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
<script>
$(document).ready(function() {
    var onclick_max = 1024;
    var answerlist_max = 2048;
    //var onclick_max = 1024 - $("#onclick").val().length;
    //$('#onclick_feedback').html(onclick_max + ' characters remaining');
    //$('#answerlist_feedback').html(answerlist_max + ' characters remaining');
    calcOnClick(onclick_max);
   calcAnswerList(answerlist_max);

    $('#onclick').keyup(function() {
	    calcOnClick(onclick_max);
    });
    $('#answerlist').keyup(function() {
	    calcAnswerList(answerlist_max);

    });

});

function calcOnClick(max)
{
        var text_length = $('#onclick').val().length;
        var text_remaining = max - text_length;

	if (text_remaining < 0)
	{
		$('#onclick').val($('#onclick').val().substring(0,max));
		text_remaining = 0;
	}

        $('#onclick_feedback').html(text_remaining + ' characters remaining');
}
function calcAnswerList(max)
{
        var text_length = $('#answerlist').val().length;
        var text_remaining = max - text_length;

	if (text_remaining < 0)
	{
		$('#answerlist').val($('#answerlist').val().substring(0,max));
		text_remaining = 0;
	}

        $('#answerlist_feedback').html(text_remaining + ' characters remaining');
}
</script>
  <%
%>
<!--#include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub subCreateEditQuestion(p_formid, p_fieldid)

	sSQL = "SELECT * FROM egov_action_form_questions WHERE questionid='" & p_fieldid & "'"

	set oEditQues = Server.CreateObject("ADODB.Recordset")
	oEditQues.Open sSQL, Application("DSN") , 3, 1
	
	if not oEditQues.eof then
		  sPrompt      = oEditQues("prompt")
		  sAnswerList  = oEditQues("answerlist")
		  sIsrequired  = oEditQues("isrequired")
		  sPDFFormName = oEditQues("pdfformname")
    sPushFieldID = oEditQues("pushfieldid")

		  if sIsrequired = True then
  		  	sIsrequired = " CHECKED "
  		else
		    	sIsrequired = ""
  		end if

		  iFieldType = oEditQues("fieldtype")
 else
		  response.write "Error: Question Not Found! Return to the form and try again."
		  exit sub
	end if

	set oEditQues = nothing 

'	select case iFieldType
'		 case "2"
'			  subCreateChooseSingleAnswerRadio sPrompt,sIsrequired,sAnswerList

' 		case "4"
'  			subCreateChooseSingleAnswerSelect sPrompt,sIsrequired,sAnswerList

'		 case "6"
'		  	subCreateChooseMultipleAnswersCheckbox sPrompt,sIsrequired,sAnswerList

'		 case "8"
'  			subCreateSingleLinewPrompt sPrompt,sIsrequired
		
'		 case "10"
'		  	subCreateEssay sPrompt,sIsrequired
'		 case else

' end select


'Build question with specified values
	response.write "<table>" & vbcrlf
	response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
	response.write "  <tr><td><textarea class=""PROMPT"" name=""FieldPrompt"" id=""onclick"">" & sPrompt & "</textarea><div id=""onclick_feedback""></div></td></tr>" & vbcrlf
	response.write "  <tr><td><input " & sIsrequired & " type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong>" & vbcrlf

 if iFieldType = 2 OR iFieldType = 4 OR iFieldType = 6 then
   	response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
   	response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist"" id=""answerlist"">" & sAnswerList & "</textarea><div id=""answerlist_feedback""></div></td></tr>" & vbcrlf
 end if

' if sAnswerList <> "" then
'   	response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
'   	response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist"">" & sAnswerList & "</textarea></td></tr>" & vbcrlf
' end if

'Check for PDF Merge Field
 if lcl_userhaspermission_requestmergeforms or lcl_userhaspermission_form_letters then
  		response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
	  	response.write "  <tr style=""background:#c0c0c0;""><td><strong>(Optional)</td></tr>" & vbcrlf
	  	response.write "  <tr><td><strong>Merge Field Name</strong><br />Enter the name of the merge field for this prompt used when merging request data in order to create form letters, permits, work orders, etc.  Leave blank if not used.</td></tr>" & vbcrlf
		  response.write "  <tr><td><input value=""" & sPDFFormName & """ type=""text"" class=""PROMPT"" name=""PDFName"" style=""width:300px;"" maxlength=""255"" /></td></tr>" & vbcrlf
	end if

'Check for "Push" fields
'  1. This sub-procedure checks the egov_actionline_pushfields table and returns any/all records that have the "push_feature_permission" turned-on for the org.
'  2. In order for this dropdown list to be displayed (currently) the form itself (formid) must match the formid entered on the Org Properties screen for 
'     the "Calendar Request Form #" field (Organizations.OrgRequestCalForm)
'  3. In the future, the sub-routine to check to see if the form is a "push" form will need to have each form entered as they become a "push" form or the "check" removed itself.
 if checkIsPushForm(session("orgid"), p_formid) then
    displayPushFields sPushFieldID
 else
    response.write "<input type=""hidden"" name=""pushfieldid"" id=""pushfieldid"" value=""" & iPushFieldID & """ />" & vbcrlf
 end if

	response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub subSaveQuestion(iFieldID)
	
 	if request("isrequired") <> "" then
	   	blnisrequired = 1 
 	else
	   	blnisrequired = NULL
 	end if

  if request("fieldprompt") <> "" then
     lcl_fieldprompt = request("fieldprompt")
     lcl_fieldprompt = fnRemoveNewLine(lcl_fieldprompt)
     'lcl_fieldprompt = replace(lcl_fieldprompt,"""","&quot;")
     lcl_fieldprompt = dbsafe(lcl_fieldprompt)
     lcl_fieldprompt = "'" & lcl_fieldprompt & "'"
  else
     lcl_fieldprompt = ""
  end if

  if request("answerlist") <> "" then
     lcl_answerlist = request("answerlist")
     'lcl_answerlist = replace(lcl_answerlist,"""","&quot;")
     lcl_answerlist = dbsafe(lcl_answerlist)
     lcl_answerlist = "'" & lcl_answerlist & "'"
  else
     lcl_answerlist = "NULL"
  end if

  if request("pushfieldid") <> "" then
     lcl_pushfieldid = request("pushfieldid")
  else
     lcl_pushfieldid = 0
  end if

	'Update Question/Field
 	sSQL = "UPDATE egov_action_form_questions SET "
  sSQL = sSQL & "prompt = "       & lcl_fieldprompt            & ", "
  sSQL = sSQL & "isrequired = '"  & blnisrequired              & "', "
  sSQL = sSQL & "answerlist = "   & lcl_answerlist             & ", "
  sSQL = sSQL & "pdfformname = '" & dbsafe(request("pdfname")) & "', "
  sSQL = sSQL & "pushfieldid = "  & lcl_pushfieldid
  sSQL = sSQL & " WHERE questionid = '" & iFieldID & "'"

 	set oSaveQues = Server.CreateObject("ADODB.Recordset")
	 oSaveQues.Open sSQL, Application("DSN") , 3, 1
	 set oSaveQues = nothing

end sub

'------------------------------------------------------------------------------
sub displayPushFields(iPushFieldID)

  lcl_total_pushfields = 0

  sSQL = "SELECT pushfieldid, push_table, push_column, push_column_datatype, push_column_label, push_to_feature, push_feature_permission"
  sSQL = sSQL & " FROM egov_actionline_pushfields "
  sSQL = sSQL & " ORDER BY push_to_feature, push_column_label "

 	set oGetPushFields = Server.CreateObject("ADODB.Recordset")
	 oGetPushFields.Open sSQL, Application("DSN") , 3, 1

  if not oGetPushFields.eof then
     do while not oGetPushFields.eof

        if orghasfeature(oGetPushFields("push_feature_permission")) then
           lcl_total_pushfields = lcl_total_pushfields + 1
           lcl_featurename      = getFeatureName(oGetPushFields("push_to_feature"))

           if lcl_total_pushfields = 1 then
             	response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
        	    	response.write "  <tr style=""background:#c0c0c0;""><td><strong>""Push"" Fields:</td></tr>" & vbcrlf
        	    	response.write "  <tr><td>Select the field that this field will be ""pushed"" to.</td></tr>" & vbcrlf
        	     response.write "  <tr>" & vbcrlf
        	     response.write "      <td>" & vbcrlf
        	     response.write "          <strong>Push Field: </strong>" & vbcrlf
              response.write "          <select name=""pushfieldid"" id=""pushfieldid"">" & vbcrlf
              response.write "            <option value=""""></option>" & vbcrlf
           end if

           if iPushFieldID = oGetPushFields("pushfieldid") then
              lcl_selected_pushfieldid = " selected=""selected"""
           else
              lcl_selected_pushfieldid = ""
           end if

       		  response.write "            <option value=""" & oGetPushFields("pushfieldid") & """" & lcl_selected_pushfieldid & ">" & lcl_featurename & " - " & oGetPushFields("push_column_label") & "</option>" & vbcrlf
        else
           lcl_total_pushfields = lcl_total_pushfields
        end if

        oGetPushFields.movenext

     loop

     if lcl_total_pushfields > 0 then
        response.write "          </select>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

  end if

  oGetPushFields.close
	 set oGetPushFields = nothing

 'If there are no "push" fields then build the hidden field to track the value behind-the-scenes.
  if lcl_total_pushfields < 1 then
     response.write "<input type=""hidden"" name=""pushfieldid"" id=""pushfieldid"" value=""" & iPushFieldID & """ />" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )

  if not VarType( strDB ) = vbString then DBsafe = strDB : exit function
  DBsafe = Replace( strDB, "'", "''" )

end function

'------------------------------------------------------------------------------
'sub subCreateEssay()
'sub subCreateEssay(sPrompt,sIsrequired)

	'response.write "<table>" & vbcrlf
	'response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""FieldPrompt"">" & sPrompt & "</textarea></td></tr>" & vbcrlf
	'response.write "  <tr><td><input " & sIsrequired & " type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong></td></tr>" & vbcrlf

 'subDrawPDFFormFieldName

	'response.write "  <tr><td><input type=""submit"" value=""Save Changes"" class=""button"" /></td></tr>" & vbcrlf
	'response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateChooseSingleAnswerRadio(sAnswerList)
'sub subCreateChooseSingleAnswerRadio(sPrompt,sIsrequired,sAnswerList)
	
	'response.write "<table>" & vbcrlf
	'response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""fieldprompt"">" & sPrompt & "</textarea></td></tr>" & vbcrlf
	'response.write "  <tr><td><input " & sIsrequired & " type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong>" & vbcrlf
	'response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist"">" & sAnswerList & "</textarea></td></tr>" & vbcrlf

	'subDrawPDFFormFieldName

	'response.write "<tr><td><input type=""submit"" value=""Save Changes"" /></td></tr>" & vbcrlf
	'response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateChooseSingleAnswerSelect(sAnswerList)
'sub subCreateChooseSingleAnswerSelect(sPrompt,sIsrequired,sAnswerList)

	'response.write "<table>" & vbcrlf
	'response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""FieldPrompt"">" & sPrompt & "</textarea></td></tr>" & vbcrlf
	'response.write "  <tr><td><input " & sIsrequired & " type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong>" & vbcrlf
	'response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist"">" & sAnswerList & "</textarea></td></tr>" & vbcrlf

	'subDrawPDFFormFieldName

	'response.write "  <tr><td><input type=""submit"" value=""Save Changes"" /></td></tr>" & vbcrlf
	'response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateChooseMultipleAnswersCheckbox(sAnswerList)
'sub subCreateChooseMultipleAnswersCheckbox(sPrompt,sIsrequired,sAnswerList)

	'response.write "<table>" & vbcrlf
	'response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""FieldPrompt"">" & sPrompt & "</textarea></td></tr>" & vbcrlf
	'response.write "  <tr><td><input " & sIsrequired & " type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong>" & vbcrlf
	'response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist"">" & sAnswerList & "</textarea></td></tr>" & vbcrlf

	'subDrawPDFFormFieldName

	'response.write "  <tr><td><input type=""submit"" value=""Save Changes"" /></td></tr>" & vbcrlf
	'response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateSingleLinewPrompt()
'sub subCreateSingleLinewPrompt(sPrompt,sIsrequired)
	'response.write "<table>" & vbcrlf
	'response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
	'response.write "  <tr><td><textarea class=""PROMPT"" name=""FieldPrompt"">" & sPrompt & "</textarea></td></tr>" & vbcrlf
	'response.write "  <tr><td><input " & sIsrequired & "type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong>" & vbcrlf

	'subDrawPDFFormFieldName

	'response.write "  <tr><td><input type=""submit"" value=""Save Changes"" /></td></tr>" & vbcrlf
	'response.write "</table>" & vbcrlf
'end sub

'------------------------------------------------------------------------------
'sub subDrawPDFFormFieldName()

	'User Security Check
	 'if lcl_userhaspermission_requestmergeforms or lcl_userhaspermission_form_letters then
  '		response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
	 ' 	response.write "  <tr style=""background:#c0c0c0;""><td><strong>(Optional)</td></tr>" & vbcrlf
	 ' 	response.write "  <tr><td><strong>Merge Field Name</strong><br />Enter the name of the merge field for this prompt used when merging request data in order to create form letters, permits, work orders, etc.  Leave blank if not used.</td></tr>" & vbcrlf
		'  response.write "  <tr><td><input value=""" & sPDFFormName & """ type=""text"" class=""PROMPT"" name=""PDFName"" style=""width:300px;"" maxlength=""255"" /></td></tr>" & vbcrlf
	 'end if
'end sub

'------------------------------------------------------------------------------
function fnRemoveNewLine(sValue)

	sReturnValue = sValue

	' REMOVE LINE BREAKS
    Dim rx
    Set rx = New RegExp
    rx.IgnoreCase = True
    rx.Global = True
    rx.Multiline = True
    rx.Pattern = "\r\n"  ' Set pattern.
    sReturnValue = rx.Replace(sValue, " ")
	
	fnRemoveNewLine = sReturnValue

end function

'------------------------------------------------------------------------------
sub displayButtons(iPosition)

  response.write "<p><input type=""button"" name=""saveButton" & iPosition & """ id=""saveButton" & iPosition & """ class=""button"" value=""Save Changes"" onclick=""saveChanges()"" /></p>" & vbcrlf

end sub
%>
