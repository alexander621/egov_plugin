<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"form creator") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Process form posts and identify field/question and form
 iFormID     = request("iformid")
 iOrgID      = request("iorgid")
 iIsInternal = request("isinternal")
 iIsRequired = request("isrequired")
 iTask       = ""

 if iIsInternal <> "1" then
	   iIsInternal = 0
 end if

 if request("task") <> "" then
    iTask = UCASE(request("task"))
 end if

 iFieldType = 0

 if request.servervariables("REQUEST_METHOD") = "POST" then
   	iFieldType = request("FieldType")

   	if iTask = "ADD" then
     		subAddQuestion iIsRequired, iFormID

     		iFieldType = 0  'Clear select box

       if request("blnaddanother") = "NO" then
       			response.redirect("manage_form.asp?iformid=" & iformid)  'Return to editing form
       end if
    end if
 end if
%>
<html>
<head>
	 <title>E-GovLink Administration Consule {Forms Management}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">

          <p>
             <font class="label">Forms - New Question </font>
             <input type="button" name="returnButton" id="returnButton" value="Return to Form" class="button" onclick="location.href='manage_form.asp?iformid=<%=iformid%>';" />
          </p>
          <!-- <small>[<a class="edit" href="manage_form.asp?iformid=<%'iformid%>">Return to Form</a>]</small><hr size="1" width="600px;" style="text-align:left; color:#000000;"></P> -->

          <div class="group">
          <%
            if iTask = "ADD" then
              	response.write "<small style=""color:#0000ff; font-weight:bold;"">Question Added - " & Now() & "</small>"
            end if
          %>

          <!--BEGIN: FORM FIELD SELECTION INFORMATION -->
          <div class="orgadminboxf">

          <!--BEGIN: INTRO TEXT-->
          <p>
          <strong>Instructions.</strong>
          <ul>
             	<li>Select the type of question.
             	<li>Enter the necessary question prompt/choices.
              <li>Select <strong>Save and Add Another Question</strong> or <strong>Save and Return to Form</strong>.
          </ul>
          </p>
          <p>Click the <strong>Return to Form</strong> link above to return to the form edit page without saving new question.</p>
          <!--END: INTRO TEXT-->

          <form name="frmSelection" action="add_field.asp" method="POST">
            <input type="hidden" name="iformid" value="<%=iformid%>" />
            <input type="hidden" name="iorgid" value="<%=iorgid%>" />
            <input type="hidden" name="isinternal" value="<%=iIsInternal%>" />

          <p>
          <strong>Choose Question Type</strong><br />
          Select the type of question from the list below.
        		<select onchange="document.frmSelection.submit();" name="FieldType">
         			<option <%if iFieldType="0"  then response.write " selected=""selected""" end if%> value="0">-- Select a Question Type --</option>
         			<option <%if iFieldType="2"  then response.write " selected=""selected""" end if%> value="2">Choose answer from list (Radio Box) or Text Block (No Answers)</option>
         			<option <%if iFieldType="4"  then response.write " selected=""selected""" end if%> value="4">Choose answer from drop down list One Answer (Select Box)</option>
         			<option <%if iFieldType="6"  then response.write " selected=""selected""" end if%> value="6">Choose multiple answers from list (Check Box)</option>
         			<option <%if iFieldType="8"  then response.write " selected=""selected""" end if%> value="8">Open Answer - One Line Response (Text Box)</option>
         			<option <%if iFieldType="10" then response.write " selected=""selected""" end if%> value="10">Open Answer - Essay (Text Area)</option>
         			<!--option <%if iFieldType="12" then response.write " selected=""selected""" end if%> value="12">Page Break</option-->
         	</select>
          </p>
          </form>
          <!--END: FORM FIELD SELECTION INFORMATION -->
          <%
           'BEGIN: Add question builder information ---------------------------
            if iFieldType <> "0" then
               response.write "<form name=""frmAddField"" action=""add_field.asp"" method=""POST"">" & vbcrlf
               response.write "  <input type=""hidden"" name=""iformid"" value="""    & iformid     & """ />" & vbcrlf
               response.write "  <input type=""hidden"" name=""iorgid"" value="""     & iorgid      & """ />" & vbcrlf
               response.write "  <input type=""hidden"" id=""fieldtype"" name=""fieldtype"" value="""  & iFieldType  & """ />" & vbcrlf
               response.write "  <input type=""hidden"" name=""isinternal"" value=""" & iIsInternal & """ />" & vbcrlf
               response.write "  <input type=""hidden"" name=""blnaddanother"" value=""YES"" />" & vbcrlf
               response.write "  <input type=""hidden"" name=""task"" value=""ADD"" />" & vbcrlf

              	if iTask <> "ADD" then
                		call subCreateField(iFieldType)
               end if

               response.write "</form>" & vbcrlf

            end if
           'END: Add question builder information -----------------------------
          %>
      </td>
  </tr>
</table>
	</div>
</div>

<!--#Include file="../admin_footer.asp"-->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
<script>
$(document).ready(function() {
    var onclick_max = 1024;
    var answerlist_max = 2048;
    $('#onclick_feedback').html(onclick_max + ' characters remaining');
    $('#answerlist_feedback').html(answerlist_max + ' characters remaining');

    $('#onclick').keyup(function() {
        var text_length = $('#onclick').val().length;
        var text_remaining = onclick_max - text_length;

	if (text_remaining < 0)
	{
		$('#onclick').val($('#onclick').val().substring(0,onclick_max));
		text_remaining = 0;
	}

        $('#onclick_feedback').html(text_remaining + ' characters remaining');
    });
    $('#answerlist').keyup(function() {
        var text_length = $('#answerlist').val().length;
        var text_remaining = answerlist_max - text_length;

	if (text_remaining < 0)
	{
		$('#answerlist').val($('#answerlist').val().substring(0,answerlist_max));
		text_remaining = 0;
	}

        $('#answerlist_feedback').html(text_remaining + ' characters remaining');
    });


});
    function validate(addanother)
    {
	    msg = "";
	    //Don't allow user to add without answers.
	    if ($('#fieldtype').val() == "6" && $('#answerlist').val() == "")
	    {
		    msg += "\nYou must enter at least one answer choice."
	    }


	    if (msg == "")
	    {
	
	    	if (!addanother)
	    	{
		    	document.frmAddField.blnaddanother.value='NO';
	    	}
	    	document.frmAddField.submit();
	    }
	    else
	    {
		    alert("The form couldn't be saved for the following reasons:\n" + msg);
	    }
    }
</script>
</body>
</html>
<%
'------------------------------------------------------------------------------
sub subCreateField(iFieldType)

 'FieldType values:
 '-----------------
 '  2  = Radio
 '  4  = Select
 '  6  = Checkbox
 '-----------------
 '  8  = Text
 '  10 = Text Area
 '-----------------
  if iFieldType = "2" OR iFieldType = "4" OR iFieldType = "6" then
     lcl_displayAnswerList = "Y"
  elseif iFieldType = "8" OR iFieldType = "10" then
     lcl_displayAnswerList = "N"
  end if

	 response.write "<table style=""background:#e0e0e0;"">" & vbcrlf
	 response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below. (limit of 1024 characters)</td></tr>" & vbcrlf
	 response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick"" id=""onclick""></textarea><div id=""onclick_feedback""></div></td></tr>" & vbcrlf
	 
	 response.write "  <tr><td><input type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong></td></tr>" & vbcrlf

  if lcl_displayAnswerList = "Y" then
   	 response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below. "
	  if iFieldType = "2" then
		  response.write "<b>Leave this blank for Text Block</b>"
	  end if
	 response.write " (limit of 2048 characters)</td></tr>" & vbcrlf
	    response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist"" id=""answerlist""></textarea><div id=""answerlist_feedback""></div></td></tr>" & vbcrlf
  else
     subDrawPDFFormFieldName()
  end if

	 response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <input type=""button"" value=""Save and Add Another Question"" class=""button"" onclick=""validate(true);"" />" & vbcrlf
  response.write "          <input type=""button"" value=""Save and Return to Form"" class=""button"" onclick=""validate(false);"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
	 response.write "</table>" & vbcrlf

'  select Case iFieldType
'  		 case "2"
'       	call subCreateChooseSingleAnswerRadio()
'   		case "4"
'       	call subCreateChooseSingleAnswerSelect()
'	 	  case "6"
'       	call subCreateChooseMultipleAnswersCheckbox()
'   		case "8"
'       	call subCreateSingleLinewPrompt()
'  	 	case "10"
'       	call subCreateEssay()
'  		 case else
'	 end if

end sub

'------------------------------------------------------------------------------
sub subAddQuestion(p_isRequired, p_formid)

  if request("isinternal") > 0 then
     lcl_query_isInternal = "= " & request("isinternal")
     lcl_isInternal       = request("isinternal")
  else
     lcl_query_isInternal = "<> 1 OR isinternalonly IS NULL "
     lcl_isInternal       = 0
  end if
	
  if p_isRequired <> "" then
   		blnisrequired = 1
  else
     blnisrequired = NULL
  end if

 'Get the max sequence id

  sSQL = "SELECT isnull(MAX(sequence),0) AS max_sequence "
  sSQL = sSQL & " FROM egov_action_form_questions "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND formid = " & p_formid
  sSQL = sSQL & " AND (isinternalonly " & lcl_query_isInternal & ") "

 	set oGetMaxSeq = Server.CreateObject("ADODB.Recordset")
	 oGetMaxSeq.Open sSQL, Application("DSN"), 3, 1

  if not oGetMaxSeq.eof then
     lcl_max_sequence = oGetMaxSeq("max_sequence") + 1
  else
     lcl_max_sequence = "1"
  end if

  oGetMaxSeq.close
  set oGetMaxSeq = nothing

	'Insert new question/field
 	sSQL = "INSERT INTO egov_action_form_questions ("
  sSQL = sSQL & "formid, "
  sSQL = sSQL & "orgid, "
  sSQL = sSQL & "prompt, "
  sSQL = sSQL & "fieldtype, "
  sSQL = sSQL & "sequence, "
  sSQL = sSQL & "isrequired, "
  sSQL = sSQL & "answerlist, "
  sSQL = sSQL & "isinternalonly, "
  sSQL = sSQL & "pdfformname "
  sSQL = sSQL & ") VALUES ("
  sSQL = sSQL & "'" & dbsafe(request("iformid"))     & "', "
  sSQL = sSQL & "'" & dbsafe(session("orgid"))       & "', "
  sSQL = sSQL & "'" & dbsafe(request("onclick")) & "', "
  sSQL = sSQL & "'" & dbsafe(request("fieldtype"))   & "', "
  sSQL = sSQL & lcl_max_sequence                     & ", "
  sSQL = sSQL & "'" & blnisrequired                  & "', "
  sSQL = sSQL & "'" & dbsafe(request("answerlist"))  & "', "
  sSQL = sSQL & "'" & lcl_isInternal                 & "', "
  sSQL = sSQL & "'" & dbsafe(request("PDFName"))     & "' "
  sSQL = sSQL & ")"

 	set oNewQues = Server.CreateObject("ADODB.Recordset")
	 oNewQues.Open sSQL, Application("DSN"), 3, 1
	 set oNewQues = nothing

end sub

'------------------------------------------------------------------------------
sub subDrawPDFFormFieldName()

	'User Security Check
 	if userhaspermission(session("userid"),"requestmergeforms") OR userhaspermission(session("userid"),"form letters") then
   		response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
   		response.write "  <tr style=""background:#c0c0c0;""><td><strong>(Optional)</strong></td></tr>" & vbcrlf
   		response.write "  <tr><td><strong>Merge Field Name</strong><br />Enter the name of the merge field for this prompt used when merging request data in order to create form letters, permits, work orders, etc.  Leave blank if not used.</td></tr>" & vbcrlf
   		response.write "  <tr><td><input type=""text"" class=""prompt"" name=""PDFName"" style=""width:300px;"" maxlength=""255"" /></td></tr>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )

  if not VarType( strDB ) = vbString then DBsafe = strDB : exit function
  DBsafe = Replace( strDB, "'", "''" )

end function

'------------------------------------------------------------------------------
'sub subCreate_Text_or_TextArea()

' 	response.write "<table style=""background:#e0e0e0;"">" & vbcrlf
' 	response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
' 	response.write "  <tr><td><textarea class=""prompt"" name=""onclick""></textarea></td></tr>" & vbcrlf
' 	response.write "  <tr><td><input type=""checkbox"" name=""isrequired""><strong>Is Required?</strong></td></tr>" & vbcrlf

'  subDrawPDFFormFieldName()

' 	response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" class=""button"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" class=""button"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
' 	response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreate_CheckboxRadioSelect()

'	 response.write "<table style=""background:#e0e0e0;"">" & vbcrlf
'	 response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr><td><input type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong></td></tr>" & vbcrlf
'	 response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" class=""button"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" class=""button"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'	 response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateEssay()

'	 response.write "<table style=""background:#e0e0e0;"">" & vbcrlf
'	 response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr><td><input type=""checkbox"" name=""validation"" /><strong>Is Required?</strong></td></tr>" & vbcrlf

'  subDrawPDFFormFieldName()

'	 response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'	 response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateSingleLinewPrompt()

' 	response.write "<table style=""background:#e0e0e0;"">" & vbcrlf
' 	response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
' 	response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick""></textarea></td></tr>" & vbcrlf
' 	response.write "  <tr><td><input type=""checkbox"" name=""isrequired""><strong>Is Required?</strong></td></tr>" & vbcrlf

'  subDrawPDFFormFieldName()

' 	response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
' 	response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateChooseSingleAnswerRadio()
	
'	 response.write "<table>" & vbcrlf
'	 response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr><td><input type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong></td></tr>" & vbcrlf
'	 response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'	 response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateChooseSingleAnswerSelect()

'	 response.write "<table>" & vbcrlf
'	 response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr><td><input type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong></td></tr>" & vbcrlf
'	 response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'	 response.write "</table>" & vbcrlf

'end sub

'------------------------------------------------------------------------------
'sub subCreateChooseMultipleAnswersCheckbox()

'	 response.write "<table>" & vbcrlf
'	 response.write "  <tr><td><strong>Question Prompt</strong><br />Enter the text for your question below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""onclick""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr><td><input type=""checkbox"" name=""isrequired"" /><strong>Is Required?</strong></td></tr>" & vbcrlf
'	 response.write "  <tr><td><strong>Answer Choices</strong><br />Enter each choice on separate lines below.</td></tr>" & vbcrlf
'	 response.write "  <tr><td><textarea class=""PROMPT"" name=""answerlist""></textarea></td></tr>" & vbcrlf
'	 response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <input type=""submit"" value=""Save and Add Another Question"" />" & vbcrlf
'  response.write "          <input type=""button"" value=""Save and Return to Form"" onclick=""document.frmAddField.blnaddanother.value='NO';document.frmAddField.submit();"" />" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'	 response.write "</table>" & vbcrlf

'end sub
%>
