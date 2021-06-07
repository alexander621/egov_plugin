<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<% 
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
' 1.0	 02/05/07	 John Stullenberger - Initial Version
' 2.0  11/16/07  David Boyer - Added additional hidden field to radio and checkbox options to handle the
'                              bug that no data is sent when a radio/checkbox is uncheck and posted.
' 2.1  08/20/08  David Boyer - Added javascript field length check to textarea fields.
' 2.2  07/28/10  David Boyer - Fixed issue with double-quotes used in questions/answers and inserted into fields as values.
' 2.3  07/30/10  David Boyer - Combined "internal only fields" screen into this screen (correction_admin_fields.asp)
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim sError, lcl_scripts
 sLevel = "../../"  'Override of value from common.asp

'Set timezone information
 session("iUserOffset") = request.cookies("tz")

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
 lcl_hidden = "hidden"

'Set the field lengths for the custom/internal fields
 lcl_text_field_length     = 1024
 lcl_textarea_field_length = 4000

'Determine which field type to display
'INT - Internal/Admin Only
'PUB - Public
 if request("ftype") <> "" then
    lcl_fieldtype = UCASE(request("ftype"))
 else
    lcl_fieldtype = "PUB"
 end if
%>
<html>
<head>
  <title><%=langBSHome%></title>

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

function submitForm() {
  document.getElementById("frmblob").submit();
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
		.savemsg                     {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
</style>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();">

<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"--> 
<%
 'BEGIN: Page Content ---------------------------------------------------------
  response.write "<div id=""content"">" & vbcrlf
  response.write " 	<div id=""centercontent"">" & vbcrlf
  response.write "<font size=""+1""><strong>Edit Request Form</strong></font><br />" & vbcrlf
  response.write "<div><input type=""button"" name=""backButton"" id=""backButton"" value=""Return to Request"" class=""button"" onclick=""location.href='../action_respond.asp?control=" & request("irequestid") & "';"" /></div>" & vbcrlf
  'response.write "<img align=""absmiddle"" src=""../../../admin/images/arrow_2back.gif"" /> <a href=""../action_respond.asp?control=" & request("irequestid") & """>Return to Request</a>" & vbcrlf

 'Display to user that values were saved
		'if request("r") = "save" then
		'  	response.write "<p><span class=""savemsg"">Saved " & Now() & ".</span></p>" & vbcrlf
  'end if

  displayButtons "TOP", request("irequestid")

 'Determine if the "internal only" fields (INT) are to be displayed or the "public" fields (PUB).
  if lcl_fieldtype = "INT" then
   		if userhaspermission(session("userid"), "internalfields" ) then
 		   		if fnAdminSubmitted(request("irequestid")) then
         		subDisplayQuestions request("irequestid"), request("status"), 1, lcl_fieldtype, "EDIT"
        else
         		subDisplayQuestions request("irequestid"), request("status"), 1, lcl_fieldtype, "ADD"
        end if
     end if
  else
   		subDisplayQuestions request("irequestid"), request("status"), 0, lcl_fieldtype, "EDIT"
  end if

  displayButtons "BOTTOM", request("irequestid")

  response.write " 	</div>" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: Page Content -----------------------------------------------------------
%>
<!--#Include file="../../admin_footer.asp"-->  
<%
 'Determine if there are any inline scripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if

  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub subDisplayQuestions(iRequestID, iStatus, blnIsAdmin, iFieldType, iScreenMode)

 'Internal Only fields OR Regular fields
  if iScreenMode = "ADD" then
     sSQL = "SELECT q.questionid as field_request_question_id, "
     sSQL = sSQL & " q.formid, "
     sSQL = sSQL & " r.orgid, "
     sSQL = sSQL & " q.prompt as field_prompt, "
     sSQL = sSQL & " q.fieldtype as field_type, "
     sSQL = sSQL & " q.isenabled, "
     sSQL = sSQL & " q.sequence as field_sequence, "
     sSQL = sSQL & " q.answerlist as field_answerlist, "
     sSQL = sSQL & " q.isrequired as field_isrequired, "
     sSQL = sSQL & " q.validationlist, "
     sSQL = sSQL & " q.isinternalonly, "
     sSQL = sSQL & " q.pdfformname as field_pdfname, "
     sSQL = sSQL & " q.pushfieldid as pushfieldid "
     sSQL = sSQL & " FROM egov_actionline_requests r "
     sSQL = sSQL &      " INNER JOIN egov_action_form_questions q ON r.category_id = q.formid "
     sSQL = sSQL & " WHERE r.action_autoid = " & iRequestID
     sSQL = sSQL & " AND q.isinternalonly = 1 "
     sSQL = sSQL & " ORDER BY q.sequence "

     lcl_formtype = "adminfields"

  else

   		sSQL = "SELECT submitted_request_field_id as field_request_question_id, "
     sSQL = sSQL & " submitted_request_field_prompt as field_prompt, "
     sSQL = sSQL & " submitted_request_id, "
     sSQL = sSQL & " submitted_request_field_sequence as field_sequence, "
     sSQL = sSQL & " submitted_request_field_answerlist as field_answerlist, "
     sSQL = sSQL & " submitted_request_field_isrequired as field_isrequired, "
     sSQL = sSQL & " submitted_request_field_pdf_name as field_pdfname, "
     sSQL = sSQL & " submitted_request_field_type_id as field_type, "
     sSQL = sSQL & " submitted_request_field_isinternal, "
     sSQL = sSQL & " submitted_request_field_pushfieldid as pushfieldid "
     sSQL = sSQL & " FROM egov_submitted_request_fields "
     sSQL = sSQL & " WHERE submitted_request_id='" & iRequestID & "' "

    	if blnIsAdmin then
        sSQL = sSQL & " AND (submitted_request_field_isinternal = 1) "
    	else
        sSQL = sSQL & " AND (submitted_request_field_isinternal = 0 OR submitted_request_field_isinternal IS NULL) "
     end if

     sSQL = sSQL & " ORDER BY submitted_request_field_sequence"

     lcl_formtype = "nonblob"

  end if
'response.write sSQL & "<br />"
  response.write "<form name=""frmblob"" id=""frmblob"" action=""correction_request_form_cgi.asp"" method=""POST"">" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""irequestid"" value=""" & iRequestID & """ />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""status"" value=""" & iStatus & """ />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""ftype"" value=""" & iFieldType & """ />" & vbcrlf
		response.write "  <input type=""" & lcl_hidden & """ name=""formtype"" value="""   & lcl_formtype                   & """ />" & vbcrlf
		response.write "  <input type=""" & lcl_hidden & """ name=""substatus"" value="""  & request("substatus")           & """ />" & vbcrlf

		response.write "<div class=""shadow"">" & vbcrlf
		response.write "<table class=""tablelist"" cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">" & vbcrlf

	 set oQuestions = Server.CreateObject("ADODB.Recordset")
	 oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	 if not oQuestions.eof then
   		'response.write "<form name=""frmblob"" action=""correction_request_form_cgi.asp"" method=""POST"">" & vbcrlf
 				'response.write "  <input type=""" & lcl_hidden & """ name=""irequestid"" value=""" & request("irequestid") & """ />" & vbcrlf
 				'response.write "  <input type=""" & lcl_hidden & """ name=""status"" value=""" & request("status") & """ />" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <th class=""corrections"" align=""left"" colspan=""2"">&nbsp;Request Form</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <td colspan=""2"">" & vbcrlf
     response.write "          <p class=""instructions"">Please update the request form information and press <b>Save</b> when finished making changes.</p>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
    	'response.write "  <tr>" & vbcrlf
     'response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
     '                          displayButtons request("irequestid")
     'response.write "      </td>" & vbcrlf
     'response.write "  </tr>" & vbcrlf
   		response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
	
     iQuestionCount = 0

    	do while not oQuestions.eof
     			'arrAnswers = split(oQuestions("field_answerlist"),chr(10))
     			iQuestionCount        = iQuestionCount + 1
        lcl_answerslist       = ""
      		sIsRequired           = oQuestions("field_isrequired")
        sDisplayRequiredLabel = ""
        lcl_prompt_req_value  = oQuestions("field_prompt")
        lcl_prompt_display    = oQuestions("field_prompt")
        lcl_prompt_value      = oQuestions("field_prompt")

     			if oQuestions("field_answerlist") <> "" then
			        lcl_answerslist = oQuestions("field_answerlist")
			        lcl_answerslist = Replace(lcl_answerslist,Chr(34),"&quot;")
     			end if

      		if sIsRequired then
   		     	sDisplayRequiredLabel = " <font color=""#ff0000"">*</font> "
     			end if

       'Format the Prompt and Value
        if lcl_prompt_value <> "" then
           lcl_prompt_value = replace(lcl_prompt_value,"""","&quot;")
        end if

        if lcl_prompt_req_value <> "" then
           lcl_prompt_req_value = left(lcl_prompt_req_value,75) & "..."
           lcl_prompt_req_value = replace(lcl_prompt_req_value,"""","&quot;")
        end if

     		 select case oQuestions("field_type")

       			case "2"
      			   'Build Radio Question
             lcl_checked_total = 0
          			arrAnswers        = split(oQuestions("field_answerlist"),chr(10))

          			response.write "<tr>" & vbcrlf
             response.write "    <td class=""question"">" & sDisplayRequiredLabel & lcl_prompt_display & "</td>" & vbcrlf
             response.write "</tr>" & vbcrlf
       			   response.write "<tr>" & vbcrlf
             response.write "    <td>" & vbcrlf

          			for alist = 0 to ubound(arrAnswers)
                lcl_checked = ""
                lcl_answer  = ""

                if iScreenMode = "EDIT" then
                   lcl_checked_total = lcl_checked_total

                   if IsQuestionValueMatch(oQuestions("field_request_question_id"),arrAnswers(alist),"CHECKED") = "CHECKED" then
                      lcl_checked       = " checked=""checked"""
                      lcl_checked_total = lcl_checked_total + 1
                   end if
                end if

                if arrAnswers(alist) <> "" then
                   'lcl_answerslist = arrAnswers(alist)
                   lcl_answer = arrAnswers(alist)
                'else
                   'lcl_answerslist = ""
                '   lcl_answer = ""
                end if

               'Format the value
                'if lcl_answerslist <> "" then
                '   lcl_answerslist = replace(lcl_answerslist,"""","&quot;")
                'end if
                if lcl_answer <> "" then
                   lcl_answer = replace(lcl_answer,"""","&quot;")
                end if

          			   'response.write "<tr>" & vbcrlf
                'response.write "    <td>" & vbcrlf
                if alist > 0 then
                   response.write "<br />" & vbcrlf
                end if

                response.write "<input type=""radio"" name=""frmanswer" & oQuestions("field_request_question_id") & """ value=""" & lcl_answer & """ class=""formradio""" & lcl_checked & " />" & arrAnswers(alist) & vbcrlf
'response.write("<br />[" & lcl_answer & "]")
             			'if sIsRequired then
                '			response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
             			'end if

             			'response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
                'response.write "    </td>" & vbcrlf
                'response.write "</tr>" & vbcrlf
          			next

          			if sIsRequired then
             			response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
          			end if

          			response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
             response.write "    </td>" & vbcrlf
             response.write "</tr>" & vbcrlf

             if lcl_checked_total > 0 then
                lcl_checked = ""
             else
                lcl_checked = " checked=""checked"""
             end if

             response.write "<tr style=""display: none"">"
             response.write "    <td><input type=""radio"" name=""frmanswer" & oQuestions("field_request_question_id") & """ value=""default_novalue"" " & lcl_checked & " /></td>" & vbcrlf
             response.write "</tr>"

       			case "4"
      			   'Build Select Question
          			response.write "<tr>" & vbcrlf
             response.write "    <td class=""question"">" & sDisplayRequiredLabel & lcl_prompt_display & "</td>" & vbcrlf
             response.write "</tr>" & vbcrlf

          			arrAnswers = split(oQuestions("field_answerlist"),chr(10))
			
       						response.write "<tr>" & vbcrlf
             response.write "    <td>" & vbcrlf
             response.write "        <select class=""formselect"" name=""frmanswer" & oQuestions("field_request_question_id") & """>" & vbcrlf

          			for alist = 0 to ubound(arrAnswers)
                lcl_selected = ""

                if iScreenMode = "EDIT" then
                   if IsQuestionValueMatch(oQuestions("field_request_question_id"),arrAnswers(alist),"SELECTED") = "SELECTED" then
                      lcl_selected = " selected=""selected"""
                   end if
                end if

          			   response.write "<option value=""" & formatSelectOptionValue(arrAnswers(alist)) & """" & lcl_selected & ">" & arrAnswers(alist) & "</option>" & vbcrlf
          			next

          			response.write "        </select>" & vbcrlf

           		if sIsRequired then
            				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
             end if

             response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
             response.write "    </td>" & vbcrlf
             response.write "</tr>" & vbcrlf

       			case "6"
      			   'Build Checkbox Question
          			response.write "<tr>" & vbcrlf
             response.write "    <td class=""question"">" & sDisplayRequiredLabel & lcl_prompt_display & "</td>" & vbcrlf
             response.write "</tr>" & vbcrlf

          			arrAnswers = split(oQuestions("field_answerlist"),chr(10))

          			i = 0

			    						response.write "<tr>" & vbcrlf
          			response.write "    <td>" & vbcrlf

   	  							for alist = 0 to ubound(arrAnswers)
   			    						i = i + 1
  		  		   					lcl_checked = ""
                lcl_answer  = arrAnswers(alist)

       									if lcl_answer <> "" then
		    				   			  	lcl_answer = replace(lcl_answer,"""","&quot;")
   	    								end if

     									  if iScreenMode = "EDIT" then
			        								lcl_checked_total = lcl_checked_total

		  				     			  	if IsQuestionValueMatch(oQuestions("field_request_question_id"),arrAnswers(alist),"CHECKED") = "CHECKED" then
   	  				    				  		lcl_checked = " checked=""checked"""
  				  		   			    		lcl_checked_total = lcl_checked_total + 1
       	  									'else
		      		  		   		'			lcl_checked = ""
   			        					'			lcl_checked_total = lcl_checked_total
   					  		    			end if
     							  		'else
     									  '  	lcl_checked = ""
          						end if

       								'Format the value
       									'lcl_answerslist = arrAnswers(alist)


       									'if lcl_answerslist <> "" then
		    				   			'  	lcl_answerslist = replace(lcl_answerslist,"""","&quot;")
   	    								'end if

   			    						'response.write "<tr>" & vbcrlf
             			'response.write "    <td>" & vbcrlf
                if alist > 0 then
                   response.write "<br />" & vbcrlf
                end if

   			          response.write "        <input type=""checkbox"" name=""frmanswer" & oQuestions("field_request_question_id") & """ id=""fmanswer_" & iQuestionCount & "_" & i & """ value=""" & lcl_answer & """ class=""formcheckbox"" onclick=""validateCheckbox('fmanswer_" & iQuestionCount & "')""" & lcl_checked & " />" & arrAnswers(alist) & vbcrlf

     			        'if sIsRequired then
           			  '   response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
     	  						  'end if

       			      'response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
             			'response.write "    </td>" & vbcrlf
   	    		      'response.write "</tr>"
     			  			next

  			        if sIsRequired then
        			     response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
  	  						  end if

    			      response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
          			response.write "    </td>" & vbcrlf
	    		      response.write "</tr>"

          			i = i + 1

       			 		if lcl_checked_total > 0 then
          						lcl_checked = ""
       	  			else
					          	lcl_checked = " checked=""checked"""
      		  			end if

       						response.write "<tr style=""display: none"">" & vbcrlf
		    							response.write "    <td>" & vbcrlf
				    					response.write "        <input type=""checkbox"" name=""frmanswer" & oQuestions("field_request_question_id") & """ id=""fmanswer_" & iQuestionCount & "_" & i & """ value=""default_novalue"" onclick=""validateCheckbox('fmanswer_" & iQuestionCount & "')""" & lcl_checked & " />" & vbcrlf
						    			response.write "        <span id=""total_options_fmanswer_" & iQuestionCount & """>" & i & "</span>" & vbcrlf
								    	response.write "    </td>" & vbcrlf
  									  response.write "</tr>" & vbcrlf

       			case "8"
      			   'Build Text Question
             lcl_value = ""

             if iScreenMode = "EDIT" then
                lcl_value = getQuestionValue(oQuestions("field_request_question_id"))
             end if

        					response.write "<tr>" & vbcrlf
             response.write "    <td class=""question"">" & sDisplayRequiredLabel & lcl_prompt_display & "</td>" & vbcrlf
             response.write "</tr>"& vbcrlf
          			response.write "<tr>" & vbcrlf
             response.write "    <td>" & vbcrlf
             response.write "        <input name=""frmanswer" & oQuestions("field_request_question_id") & """ value=""" & lcl_value & """ type=""text"" style=""width:300px;"" maxlength=""" & lcl_text_field_length & """>" & vbcrlf

          			if sIsRequired then
            				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
             end if

           		response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
             response.write "    </td>" & vbcrlf
             response.write "</tr>" & vbcrlf

       			case "10"
      			   'Build TextArea Question
             lcl_value = ""

             if iScreenMode = "EDIT" then
                lcl_value = getQuestionValue(oQuestions("field_request_question_id"))
             end if

       						response.write "<tr>" & vbcrlf
             response.write "    <td class=""question"">" & sDisplayRequiredLabel & lcl_prompt_display & "</td>" & vbcrlf
             response.write "</tr>" & vbcrlf
       						response.write "<tr>" & vbcrlf
       			   response.write "    <td>" & vbcrlf
       			   response.write "        <textarea name=""frmanswer" & oQuestions("field_request_question_id") & """ id=""frmanswer" & oQuestions("field_request_question_id") & """ class=""formtextarea"" maxlength=""" & lcl_textarea_field_length & """ onchange=""checkMaxLength();"">" & lcl_value & "</textarea>" & vbcrlf

        					if sIsRequired then
            				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
             end if

          			response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
       			   response.write "    </td>" & vbcrlf
       			   response.write "</tr>" & vbcrlf

       			case else

     		 end select

        lcl_value = ""

        if iScreenMode = "EDIT" then
           lcl_value = GetFormFieldNameValue(oQuestions("field_request_question_id"))
        end if

      		response.write "<tr>" & vbcrlf
        response.write "    <td>&nbsp;" & vbcrlf
     			response.write "<input type=""" & lcl_hidden & """ name=""fieldtype"   & iQuestionCount & """ value=""" & oQuestions("field_type")       & """ />" & vbcrlf
     			response.write "<input type=""" & lcl_hidden & """ name=""answerslist" & iQuestionCount & """ value=""" & lcl_answerslist                & """ />" & vbcrlf
     			response.write "<input type=""" & lcl_hidden & """ name=""isrequired"  & iQuestionCount & """ value=""" & oQuestions("field_isrequired") & """ />" & vbcrlf
     			response.write "<input type=""" & lcl_hidden & """ name=""pdfname"     & iQuestionCount & """ value=""" & oQuestions("field_pdfname")    & """ />" & vbcrlf
     			response.write "<input type=""" & lcl_hidden & """ name=""sequence"    & iQuestionCount & """ value=""" & oQuestions("field_sequence")   & """ />" & vbcrlf
     			response.write "<input type=""" & lcl_hidden & """ name=""pushfieldid" & iQuestionCount & """ value=""" & oQuestions("pushfieldid")      & """ />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ name=""submitted_request_form_field_name" & oQuestions("field_request_question_id") & """ value=""" & lcl_value & """ size=""30"" maxlength=""255"" />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ name=""submitted_request_pushfieldid"     & oQuestions("field_request_question_id") & """ value=""" & oQuestions("pushfieldid") & """ size=""30"" maxlength=""255"" />" & vbcrlf
      		response.write "    </td>" & vbcrlf
        response.write "</tr>" & vbcrlf

        oQuestions.movenext

    	loop

  		'Display Buttons
   		'response.write "<tr>" & vbcrlf
     'response.write "    <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
     '                        displayButtons request("irequestid")
 		  'response.write "    </td>" & vbcrlf
     'response.write "</tr>" & vbcrlf
 				'response.write "</table>" & vbcrlf
 				'response.write "</div>" & vbcrlf
 				'response.write "</form>" & vbcrlf
	
  else
     if iFieldType = "INT" then
        response.write "  <tr>" & vbcrlf
        response.write "      <td align=""center"" style=""color:#ff0000; padding-top:10px"">" & vbcrlf
        response.write "         <p>There are no administrative fields assigned to this form.</p>" & vbcrlf
        response.write "         <p>Administrative fields can be added by users with the <strong>Form Creator</strong> security feature enabled.</p>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        lcl_scripts = lcl_scripts & "document.getElementById(""saveButtonTOP"").disabled = true;" & vbcrlf
        lcl_scripts = lcl_scripts & "document.getElementById(""saveButtonBOTTOM"").disabled = true;" & vbcrlf

     else
       'Form Information non-blob not found
      		response.write "<p><div class=""correctionsboxnotfound"">Note: The original request question\answer formatting is not available for this request.  The form was submitted on a previous release of the E-Gov Link software.</div></p>" & vbcrlf
      		'response.write "<form name=""frmblob"" action=""correction_request_form_cgi.asp"" method=""POST"">" & vbcrlf
      		'response.write "  <input type=""" & lcl_hidden & """ name=""irequestid"" value=""" & request("irequestid") & """>" & vbcrlf
      		'response.write "  <input type=""" & lcl_hidden & """ value=""" & request("status") & """ name=""status"">" & vbcrlf
       	response.write "  <input type=""" & lcl_hidden & """ name=""formtype"" value=""blob"">" & vbcrlf
      		response.write "<div class=""shadow"">" & vbcrlf
      		response.write "<table class=""tablelist"" cellpadding=""0"" cellspacing=""0"" style=""padding-left:10px;"">" & vbcrlf
      		response.write "  <tr><th class=""corrections"" align=""left"" colspan=""2"">&nbsp;Request Form</th></tr>" & vbcrlf

     		'Display Instructions
      		response.write "  <tr><td colspan=""2""><p class=""instructions"">Please update the request form information and press <b>Save</b> when finished making changes.</p></td></tr>" & vbcrlf
       	response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

      		SubGenericEdit irequestid
     end if

   		response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
  end if

  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf

 	set oQuestions = nothing 

end sub
	
'------------------------------------------------------------------------------
function GetQuestionValue(ifieldid)

 	sSQL = "SELECT DISTINCT CAST(submitted_request_field_response as VARCHAR(MAX))  as submitted_request_field_response "
  sSQL = sSQL & " FROM egov_submitted_request_field_responses "
  sSQL = sSQL & " WHERE submitted_request_field_id='" & ifieldid & "'"

 	set oQuestionsValue = Server.CreateObject("ADODB.Recordset")
	 oQuestionsValue.Open sSQL, Application("DSN"), 3, 1

 	if not oQuestionsValue.eof then
     if oQuestionsValue("submitted_request_field_response") <> "" then
        lcl_field_response = oQuestionsValue("submitted_request_field_response")
        lcl_field_response = replace(lcl_field_response,"""","&quot;")
     else
        lcl_field_response = ""
     end if


	 	  GetQuestionValue = lcl_field_response
	 end if

	 set oQuestionsValue = nothing

end function

'------------------------------------------------------------------------------
function GetFormFieldNameValue(ifieldid)

	sSQL = "SELECT distinct submitted_request_form_field_name "
 sSQL = sSQL & " FROM egov_submitted_request_field_responses "
 sSQL = sSQL & " WHERE submitted_request_field_id='" & ifieldid & "'"

	set oFormFieldValue = Server.CreateObject("ADODB.Recordset")
	oFormFieldValue.Open sSQL, Application("DSN"), 3, 1

	if not oFormFieldValue.eof then
		  GetFormFieldNameValue = oFormFieldValue("submitted_request_form_field_name")
	end if

	set oFormFieldValue = nothing

end function

'------------------------------------------------------------------------------
function IsQuestionValueMatch(ifieldid,sValue,sTrueValue)

	sReturnValue = ""

 lcl_value = sValue

 if lcl_value <> "" then
    lcl_value = trim(lcl_value)
    lcl_value = replace(lcl_value,chr(10),"")
    lcl_value = replace(lcl_value,chr(13),"")
    lcl_value = replace(lcl_value,"&quot;","""")
    lcl_value = dbsafe(lcl_value)
 end if

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_submitted_request_field_responses "
 sSQL = sSQL & " WHERE submitted_request_field_id='"       & ifieldid   & "' "
' sSQL = sSQL & " AND submitted_request_field_response = '" & lcl_value & "'"

	set oQuesMatch = Server.CreateObject("ADODB.Recordset")
	oQuesMatch.Open sSQL, Application("DSN"), 3, 1

	if not oQuesMatch.eof then
    do while not oQuesMatch.eof
       if trim(oQuesMatch("submitted_request_field_response")) <> "" then
          lcl_submitted_request_field_response = oQuesMatch("submitted_request_field_response")
          lcl_submitted_request_field_response = trim(lcl_submitted_request_field_response)
          lcl_submitted_request_field_response = replace(lcl_submitted_request_field_response,chr(10),"")
          lcl_submitted_request_field_response = replace(lcl_submitted_request_field_response,chr(13),"")
          lcl_submitted_request_field_response = replace(lcl_submitted_request_field_response,"&quot;","""")
          lcl_submitted_request_field_response = dbsafe(lcl_submitted_request_field_response)
       else
          lcl_submitted_request_field_response = ""
       end if

   		  if lcl_submitted_request_field_response = lcl_value then
			       sReturnValue = sTrueValue 
     		end if

       oQuesMatch.movenext
    loop
	end if

	set oQuesMatch = nothing

	IsQuestionValueMatch = sReturnValue 

end function

'------------------------------------------------------------------------------
Sub SubGenericEdit(irequestid)

	sSQL = "SELECT comment "
 sSQL = sSQL & " FROM egov_actionline_requests "
 sSQL = sSQL & " WHERE action_autoid='" & irequestid & "'"

	Set oGenericEdit = Server.CreateObject("ADODB.Recordset")
	oGenericEdit.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oGenericEdit.EOF Then

  		sBlob = oGenericEdit("comment")

  		If Trim(sBlob) <> "" Then
   			'CHANGE TO QUESTION\ANSWER FORM

    		'SPLIT QUESTIONS INTO ARRAY
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
       				response.write "question: <input name=""question" & iQues & """ type=""" & lcl_hidden & """ value=""" & sQues & """ /><br />"
				
      				'DISPLAY
       				response.write "<tr><td class=""correctionslabel"">" & UCASE(sQues) & "</td></tr>"
       				response.write "<tr><td><textarea name=""answer" & iQues & """ class=""correctionstextarea"">" & sAnswer & "</textarea></td></tr>"
			    Next
  		End If
  		oGenericEdit.Close
	End If

	Set oGenericEdit = Nothing

end sub

'------------------------------------------------------------------------------
function fnAdminSubmitted(iRequestID)
 	blnReturnValue = False

 	sSQL = "SELECT * FROM egov_submitted_request_fields "
 	sSQL = sSQL & " WHERE submitted_request_field_isinternal = 1 "
	 sSQL = sSQL & " AND submitted_request_id='" & iRequestID & "' "
 	sSQL = sSQL & " ORDER BY submitted_request_field_sequence"

 	set aAdminSubmitted = Server.CreateObject("ADODB.Recordset")
	 aAdminSubmitted.Open sSQL, Application("DSN"), 3, 1

 	if not aAdminSubmitted.eof then
	   	blnReturnValue = True
 	end if

	 set aAdminSubmitted = nothing

 	fnAdminSubmitted = blnReturnValue

end function

'------------------------------------------------------------------------------
function formatSelectOptionValue(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,chr(10),"")
     lcl_return = replace(lcl_return,chr(13),"")
     lcl_return = replace(lcl_return,"<br>","")
     lcl_return = replace(lcl_return,"<br />","")
     lcl_return = replace(lcl_return,"<BR>","")
     lcl_return = replace(lcl_return,"<BR />","")
     lcl_return = replace(lcl_return,"""","&quot;")
  end if

  formatSelectOptionValue = lcl_return

end function

'------------------------------------------------------------------------------
sub displayButtons(iLocation, iRequestID)

  response.write "<p>" & vbcrlf
  response.write "  <input type=""button"" name=""cancelButton" & iLocation & """ id=""cancelButton" & iLocation & """ class=""button"" value=""Cancel"" onclick=""location.href='../action_respond.asp?control=" & iRequestID & "';"" />" & vbcrlf
  response.write "  <input type=""button"" name=""saveButton" & iLocation & """ id=""saveButton" & iLocation & """ class=""button"" value=""Save Changes"" onclick=""submitForm()"" />" & vbcrlf
  response.write "</p>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     lcl_value = p_value
     lcl_value = replace(lcl_value,"'","''")
  else
     lcl_value = "NULL"
  end if

  sSQL1 = "INSERT INTO my_table_dtb(notes) VALUES ('" & lcl_value & "') "

	 set oDTB2 = Server.CreateObject("ADODB.Recordset")
	 oDTB2.Open sSQL1, Application("DSN"), 3, 1

  set oDTB2 = nothing

end sub
%>
