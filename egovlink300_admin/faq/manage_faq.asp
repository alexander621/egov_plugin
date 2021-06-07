<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: manage_faq.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module edits FAQ categories
'
' MODIFICATION HISTORY
' 1.0 09/11/06	Steve Loar - Changed to add categories
' 1.2	10/09/06	Steve Loar - Security, Header and Nav changed
' 1.3	11/09/07	Steve Loar - Added pub start and end dates
' 1.3 03/20/09 David Boyer - Added "faqtype" for the new "Rumor Mill" data
' 1.4 06/10/09	David Boyer - Added checkbox for "send to" function.  (Send to features like RSS and eventually Twitter, etc.)
' 1.5 07/15/09 David Boyer - Added "push" content
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Check for the faqtype
 if request("faqtype") <> "" then
    lcl_faqtype = UCASE(request("faqtype"))
 else
    lcl_faqtype = "FAQ"
 end if

'Based on the faqtype check for the proper permission
 if lcl_faqtype = "RUMORMILL" then
    lcl_pagetitle       = "Rumor Mill"
    lcl_userpermission  = "rumormill_manage"
    lcl_feature_rssfeed = "rssfeeds_rumormill"
    lcl_pushcontent     = "pushcontent_rumormill"
 else
    lcl_pagetitle       = "FAQ"
    lcl_userpermission  = "manage faq"
    lcl_feature_rssfeed = "rssfeeds_faqs"
    lcl_pushcontent     = "pushcontent_faqs"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 iFaqID        = request("ifaqid")
 iOrgID        = session("orgid")
 lcl_success   = request("success")
 lcl_requestid = ""

 if iFaqID <> "" then
    lcl_screen_mode = "Edit"
    lcl_sendToLabel = "Update"
 else
    lcl_screen_mode = "Add"
    lcl_sendToLabel = "Create"
 end if

 'if request.servervariables("REQUEST_METHOD") = "POST" then
 '  	select case request("TASK")

 ' 		case "new_question"
  	   'Add New Question to Form
 '   			Call subAddQuestion(iorgid,iFormID)
 ' 		case else
   			'Default Action
 '  	end select
 'end if

'Check for org features
 lcl_orghasfeature_rssfeeds        = orghasfeature(lcl_feature_rssfeed)
 lcl_orghasfeature_pushcontent     = orghasfeature(lcl_pushcontent)

'Check for user permissions
 lcl_userhaspermission_rssfeeds    = userhaspermission(session("userid"),lcl_feature_rssfeed)
 lcl_userhaspermission_pushcontent = userhaspermission(session("userid"),lcl_pushcontent)

'Check for a screen message
 lcl_onload = "setMaxLength();"

 if lcl_success <> "" then
    if lcl_success = "SU" then
       lcl_msg = "Successfully Updated..."
    elseif lcl_success = "SA" then
       lcl_msg = "Successfully Created..."
    end if

    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"

 end if

'Determine if there is any additional processing needed from the past update
 if lcl_orghasfeature_rssfeeds AND lcl_userhaspermission_rssfeeds AND (lcl_success = "SU" OR lcl_success = "SA") then
    if request("sendTo_RSS") <> "" then
       lcl_onload = lcl_onload & "sendToRSS('" & request("sendTo_RSS") & "');"
    end if
 end if

'Check to see if this record is being "pushed" from a request
 if lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent then
    lcl_requestid = request("requestid")
 end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Maintain <%=lcl_pagetitle%>s}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />

 	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
	 <script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script type="text/javascript">
<!--
function doCalendar(sField) {
  var w = (screen.width - 350)/2;
		var h = (screen.height - 350)/2;
		eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=manage_faq", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function storeCaret (textEl) {
  if (textEl.createTextRange) 
		 	  textEl.caretPos = document.selection.createRange().duplicate();
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
		    var caretPos = textEl.caretPos;
  			 caretPos.text =
		    caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
			   text + ' ' : text;
  }
	  else
			textEl.value = textEl.value + text;
}

function doPicker(sFormField) {
  w = (screen.width - 350)/2;
		h = (screen.height - 350)/2;
		eval('window.open("../picker/default.asp?name=' + sFormField + '&starting_folder=CITY_ROOT", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function fnCheckSubject() {
  if (document.manage_faq.Subject.value != '') {
      return true;
		}else{
      return false;
		}
}

function validate() {
  var rege;
		var OK;
  var lcl_return_false = 0;

		// check the publication end date
		if (document.manage_faq.publicationend.value != "") {
  				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		  		Ok   = rege.test(document.manage_faq.publicationend.value);
  				if (! Ok) {
          document.getElementById("publicationend").focus();
          inlineMsg(document.getElementById("publicationend_cal").id,'<strong>Invalid Value: </strong> The "Publication End" must be in date format or blank.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'publicationend_cal');
          lcl_return_false = lcl_return_false + 1;
      }else{
          clearMsg("publicationend");
  				}
		}

		//check the publication start date
		if (document.manage_faq.publicationstart.value != "") {
  				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		  		Ok   = rege.test(document.manage_faq.publicationstart.value);
				  if (! Ok) {
          document.getElementById("publicationstart").focus();
          inlineMsg(document.getElementById("publicationstart_cal").id,'<strong>Invalid Value: </strong> The "Publication Start" must be in date format or blank.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'publicationstart_cal');
          lcl_return_false = lcl_return_false + 1;
      }else{
          clearMsg("publicationstart");
  				}
		}

  if (document.getElementById("FaqQ").value == "") {
      document.getElementById("FaqQ").focus();
      inlineMsg(document.getElementById("FaqQ").id,'<strong>Required Field Missing: </strong>Question',10,'FaqQ');
      lcl_return_false = lcl_return_false + 1;
  }else{
      clearMsg("FaqQ");
  }

  if(lcl_return_false > 0) {
     return false;
  }else{
     document.getElementById("manage_faq").action = 'save_faq.asp';
     document.getElementById("manage_faq").submit();
  }
}

<% if lcl_orghasfeature_rssfeeds AND lcl_userhaspermission_rssfeeds then %>
function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&faqtype=<%=lcl_faqtype%>';
  sParameter    += '&isAjax=Y';

  doAjax('faq_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
}
<% end if %>

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
  w = 600;
  h = 400;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";
  lcl_folderStart     = 0;

  //Determine which options will be displayed
  if((p_displayDocuments=="")||(p_displayDocuments==undefined)) {
      lcl_displayDocuments = "";
  }else{
      lcl_displayDocuments = "&displayDocuments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayActionLine=="")||(p_displayActionLine==undefined)) {
      lcl_displayActionLine = "";
  }else{
      lcl_displayActionLine = "&displayActionLine=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayPayments=="")||(p_displayPayments==undefined)) {
      lcl_displayPayments = "";
  }else{
      lcl_displayPayments = "&displayPayments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayURL=="")||(p_displayURL==undefined)) {
      lcl_displayURL = "";
  }else{
      lcl_displayURL = "&displayURL=Y";
  }

  //if(lcl_folderStart > 0) {
     //lcl_showFolderStart = "&folderStart=unpublished_documents";
  //   lcl_showFolderStart = "&folderStart=CITY_ROOT";
  //}

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
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
<body onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <div style=""margin-top:20px; margin-left:20px;"">" & vbcrlf
  response.write "            <div class=""group"">" & vbcrlf
  response.write "              <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""600"">" & vbcrlf
  response.write "                <tr>" & vbcrlf
  response.write "                    <td class=""label"">" & vbcrlf
  response.write "                        <font size=""+1""><strong>" & lcl_pagetitle & "s - " & lcl_screen_mode & "</strong></font><br /><br />" & vbcrlf
  response.write "                        <input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to " & lcl_pagetitle & " List"" class=""button"" onclick=""location.href='list_faq.asp?faqtype=" & lcl_faqtype & "';"" />" & vbcrlf

  if lcl_requestid <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent then
     response.write "<input type=""button"" name=""returnToRequestButton"" id=""returnToRequestButton"" class=""button"" value=""Return to Request"" onclick=""location.href='../action_line/action_respond.asp?control=" & lcl_requestid & "';"" />" & vbcrlf
  end if

  response.write "                    </td>" & vbcrlf
  response.write "                    <td align=""right"" valign=""bottom""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "                </tr>" & vbcrlf
  response.write "              </table>" & vbcrlf
  response.write "              <br />" & vbcrlf
  response.write "              <div class=""orgadminboxf"">" & vbcrlf
  response.write "                <form name=""manage_faq"" id=""manage_faq"" action="""" method=""post"">" & vbcrlf
  response.write "                  <input type=""hidden"" name=""iFaqID"" id=""iFaqID"" value=""" & iFaqID & """ />" & vbcrlf
  response.write "                  <input type=""hidden"" name=""faqtype"" id=""faqtype"" value=""" & lcl_faqtype & """ />" & vbcrlf
  response.write "                  <input type=""hidden"" name=""requestid"" id=""requestid"" value=""" & lcl_requestid & """ />" & vbcrlf

                                  displayButtons lcl_screen_mode

                                  GetFaqs session("orgid"), _
                                          iFaqID, _
                                          lcl_faqtype, _
                                          lcl_requestid

                                  displayButtons lcl_screen_mode

  response.write "                </form>" & vbcrlf
  response.write "              </div>" & vbcrlf
  response.write "            </div>" & vbcrlf
%>
            <!--include file="bottom_include.asp"-->
<%
  response.write "        		</div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub GetFaqs(iOrgID, iFaqID, iFAQType, iRequestID)
	Dim sSql, oForm

 iFaqQ                = ""
 iFaqA                = ""
 iPublicationStart    = ""
 iPublicationEnd      = ""
 iCreatedByID         = ""
 iCreatedByName       = ""
 iCreatedDate         = ""
 iLastUpdatedByID     = ""
 iLastUpdatedByName   = ""
 iLastUpdatedDate     = ""
 iPushedFromRequestID = iRequestID

 if iFAQType = "" then
    iFAQType = "FAQ"
 end if

 lcl_length_question = "700"
 lcl_length_answer   = "4000"

 if iFaqID <> "" then
    if isnumeric(iFaqID) then
      	sSQL = "SELECT LEFT(f.FaqQ," & lcl_length_question & ") as FaqQ, "
       sSQL = sSQL & " LEFT(CAST(f.FaqA AS VARCHAR(4000))," & lcl_length_answer & ") as FaqA, "
       sSQL = sSQL & " isnull(f.FAQCategoryId,0) as FAQCategoryId, "
       sSQL = sSQL & " f.publicationstart, "
       sSQL = sSQL & " f.publicationend, "
       sSQL = sSQL & " f.createdbyid, "
       sSQL = sSQL & " f.createddate, "
       sSQL = sSQL & " f.lastupdatedbyid, "
       sSQL = sSQL & " f.lastupdateddate, "
       sSQL = sSQL & " f.pushedfrom_requestid, "
       sSQL = sSQL & " (select u.firstname + ' ' + u.lastname from users u where u.userid = f.createdbyid) as createdbyname, "
       sSQL = sSQL & " (select u.firstname + ' ' + u.lastname from users u where u.userid = f.lastupdatedbyid) as lastupdatedbyname, "
       sSQL = sSQL & " (select a.[Tracking Number] from egov_rpt_actionline a where a.action_autoid = f.pushedfrom_requestid) as trackingnumber "
       sSQL = sSQL & " FROM FAQ f "
       sSQL = sSQL & " WHERE f.FaqID = " & iFaqID
       sSQL = sSQL & " AND f.orgid = "   & iOrgID
       sSQL = sSQL & " AND UPPER(f.faqtype) = '" & UCASE(iFAQType) & "'"

      	set oForm = Server.CreateObject("ADODB.Recordset")
      	oForm.Open sSQL, Application("DSN"), 3, 1

       if not oForm.eof then
          iFaqQ                  = oForm("FaqQ")
          iFaqA                  = oForm("FaqA")
          iFAQCategoryID         = oForm("FAQCategoryID")
          iPublicationStart      = oForm("publicationstart")
          iPublicationEnd        = oForm("publicationend")
          iCreatedByID           = oForm("createdbyid")
          iCreatedByName         = trim(oForm("createdbyname"))
          iCreatedDate           = oForm("createddate")
          iLastUpdatedByID       = oForm("lastupdatedbyid")
          iLastUpdatedByName     = trim(oForm("lastupdatedbyname"))
          iLastUpdatedDate       = oForm("lastupdateddate")
          iPushedFromRequestID   = oForm("pushedfrom_requestid")
          iPushedFromTrackingNum = oForm("trackingnumber")
       end if

      	oForm.close 
      	set oForm = nothing 

    else
       iFaqID         = 0
       iFAQCategoryID = 0
    end if
 else
    iFaqID         = 0
    iFAQCategoryID = 0
 end if

'Check to see if the user is "pushing" content from an Action Line Request
'If "yes" then
'1. Get the tracking number to display.
'2. Take all of the "answers" from the "user questions/answers" and default them into the "Question" field.
'3. Take all of the "external comments" (aka: Note(s) to Citizen) and default them into the "Answer" field.
 if iRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent then
    iFaqQ = ""
    iFaqA = ""
    iPushedFromTrackingNum = ""

    sSQL = "SELECT a.[Tracking Number] as trackingnumber "
    sSQL = sSQL & " FROM egov_rpt_actionline a "
    sSQL = sSQL & " WHERE a.action_autoid = " & iRequestID

    set oTrackNum = Server.CreateObject("ADODB.Recordset")
    oTrackNum.Open sSQL, Application("DSN"), 3, 1

    if not oTrackNum.eof then
       iPushedFromTrackingNum = oTrackNum("trackingnumber")
    end if

    sSQL = "SELECT r.submitted_request_field_response "
    sSQL = sSQL & " FROM egov_submitted_request_field_responses r "
    sSQL = sSQL & " WHERE r.submitted_request_field_id IN (SELECT f.submitted_request_field_id "
    sSQL = sSQL &                                        " FROM egov_submitted_request_fields f "
    sSQL = sSQL &                                        " WHERE f.submitted_request_id = " & iRequestID & ") "
    sSQL = sSQL & " AND r.submitted_request_field_response NOT LIKE 'default_novalue' "
    sSQL = sSQL & " AND r.submitted_request_field_response NOT LIKE '' "
    sSQL = sSQL & " AND r.submitted_request_field_response IS NOT NULL "
    'sSQL = sSQL & " ORDER BY submitted_request_field_sequence

    set oQuestionData = Server.CreateObject("ADODB.Recordset")
    oQuestionData.Open sSQL, Application("DSN"), 3, 1

    if not oQuestionData.eof then
       do while not oQuestionData.eof

          iFaqQ = iFaqQ & oQuestionData("submitted_request_field_response")

          oQuestionData.movenext
       loop
    end if

   'Get any/all "Note(s) to Citizen" from the "Action Log History" section
    sSQL = "SELECT action_externalcomment "
    sSQL = sSQL & " FROM egov_action_responses "
    sSQL = sSQL & " WHERE action_autoid = " & iRequestID
    sSQL = sSQL & " AND CAST(action_externalcomment AS VARCHAR) <> '' "
    sSQL = sSQL & " ORDER BY action_editdate DESC "

    set oAnswerData = Server.CreateObject("ADODB.Recordset")
    oAnswerData.Open sSQL, Application("DSN"), 3, 1

    if not oAnswerData.eof then
       do while not oAnswerData.eof

          iFaqA = iFaqA & oAnswerData("action_externalcomment")

          oAnswerData.movenext
       loop
    end if

    oTrackNum.close
    oQuestionData.close
    oAnswerData.close

    set oTrackNum     = nothing
    set oQuestionData = nothing
    set oAnswerData   = nothing
 end if

 response.write "<p>" & vbcrlf
 response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"">" & vbcrlf

'Display FAQ category picks
 ShowCategories iOrgID, _
                iFAQCategoryID, _
                iFAQType

 response.write "  <tr><td colspan=""2""><strong>Question:</strong></td></tr>" & vbcrlf
 response.write "  <tr><td colspan=""2""><textarea class=""formtextarea"" name=""FaqQ"" id=""FaqQ"" maxlength=""" & lcl_length_question & """ wrap=""soft"">" & iFaqQ & "</textarea></td></tr>" & vbcrlf
	response.write "  <tr valign=""bottom"">" & vbcrlf
 response.write "      <td><strong>Answer:</strong></td>" & vbcrlf
 'response.write "      <td align=""right""><input type=""button"" value=""Add Link"" class=""button"" onClick=""doPicker('manage_faq.FaqA');"" /></td>" & vbcrlf
 response.write "      <td align=""right""><input type=""button"" value=""Add Link"" class=""button"" onClick=""doPicker('manage_faq.FaqA','Y','Y','Y','Y');"" /></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"">" & vbcrlf
 response.write "          <textarea class=""formtextareaBig"" name=""FaqA"" id=""FaqA"" maxlength=""" & lcl_length_answer & """ wrap=""soft"" onselect=""storeCaret(this);"" onclick=""storeCaret(this);"" onkeyup=""storeCaret(this);"" ondblclick=""storeCaret(this);"">" & iFaqA & "</textarea>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"">" & vbcrlf
 response.write "          <strong>Publication Start:</strong>" & vbcrlf
 response.write "          &nbsp;<input type=""text"" name=""publicationstart"" id=""publicationstart"" value=""" & iPublicationStart & """ size=""10"" maxlength=""10"" onchange=""clearMsg('publicationstart_cal');"" />" & vbcrlf
 response.write "          &nbsp;<span class=""calendarimg"" style=""cursor:pointer;""><img src=""../images/calendar.gif"" id=""publicationstart_cal"" height=""16"" width=""16"" border=""0"" onclick=""clearMsg('publicationstart_cal');doCalendar('publicationstart');"" /></span>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"">" & vbcrlf
 response.write "          <strong>Publication End:</strong>" & vbcrlf
 response.write "          &nbsp;<input type=""text"" name=""publicationend"" id=""publicationend"" value=""" & iPublicationEnd & """ size=""10"" maxlength=""10"" onchange=""clearMsg('publicationend_cal');"" />" & vbcrlf
 response.write "          &nbsp;<span class=""calendarimg"" style=""cursor:pointer;""><img src=""../images/calendar.gif"" id=""publicationend_cal"" height=""16"" width=""16"" border=""0"" onclick=""clearMsg('publicationend_cal');doCalendar('publicationend');"" /></span>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf

 if lcl_orghasfeature_rssfeeds AND lcl_userhaspermission_rssfeeds then
    response.write "  <tr>" & vbcrlf
    response.write "      <td colspan=""2"">" & vbcrlf
    response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
    response.write "            <tr valign=""top"">" & vbcrlf
    response.write "                <td nowrap=""nowrap""><strong>On " & lcl_sendToLabel & " Send To:</strong></td>" & vbcrlf
    response.write "                <td>" & vbcrlf
                                        displaySendToOption "RSS", lcl_screen_mode, "Y", lcl_orghasfeature_rssfeeds, lcl_userhaspermission_rssfeeds
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf
    response.write "          </table>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Display History Info
 if iCreatedByID <> "" OR (iPushedFromRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent) then
    response.write "  <tr>" & vbcrlf
    response.write "      <td colspan=""2"">" & vbcrlf
    response.write "          <fieldset>" & vbcrlf
    response.write "            <legend>History Log&nbsp;</legend>" & vbcrlf
    response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:5px;"">" & vbcrlf

   'Created By
    if iCreatedByID <> "" then
       lcl_createdby = iCreatedByName & " on " & iCreatedDate
       response.write "              <tr>" & vbcrlf
       response.write "                  <td><strong>Created By:</strong></td>" & vbcrlf
       response.write "                  <td style=""color:#800000"">" & lcl_createdby & "</td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
    end if

   'Last Updated By
    if iLastUpdatedByID <> "" then
       lcl_lastupdatedby = iLastUpdatedByName & " on " & iLastUpdatedDate
       response.write "              <tr>" & vbcrlf
       response.write "                  <td><strong>Last Updated By:</strong></td>" & vbcrlf
       response.write "                  <td style=""color:#800000"">" & lcl_lastupdatedby & "</td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
    end if

   'Originated From
    if iPushedFromRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent then
       response.write "              <tr>" & vbcrlf
       response.write "                  <td><strong>Originated From:</strong></td>" & vbcrlf
       response.write "                  <td><a href=""../action_line/action_respond.asp?control=" & iPushedFromRequestID & """>" & iPushedFromTrackingNum & "</a></td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
    end if

    response.write "            </table>" & vbcrlf
    response.write "          </fieldset>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

 response.write "</table>" & vbcrlf
 response.write "</p>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub ShowCategories(iOrgID, iFAQCategoryId, iFAQType)
	Dim sSql, oFAQCats

	sSQL = "SELECT FAQCategoryName, "
 sSQL = sSQL & " FAQCategoryId, "
 sSQL = sSQL & " isnull(internalonly,0) AS internalonly "
 sSQL = sSQL & " FROM faq_categories "
 sSQL = sSQL & " WHERE orgid = " & iOrgID
 sSQL = sSQL & " AND UPPER(faqtype) = '" & UCASE(iFAQType) & "'"
 sSQL = sSQL & " ORDER BY displayorder"

	set oFAQCats = Server.CreateObject("ADODB.Recordset")
	oFAQCats.Open sSql, Application("DSN"), 3, 1

 if iFAQCategoryId = 0 then
			 lcl_selected_category = " selected=""selected"""
 else
    lcl_selected_category = ""
 end if

	response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"">" & vbcrlf
 response.write "          <strong>Category:</strong>&nbsp;" & vbcrlf
 response.write "          <select name=""FAQCategoryId"" id=""FAQCategoryId"">" & vbcrlf
	response.write "            <option value=""0""" & lcl_selected_category & ">None</option>" & vbcrlf

	do while not oFAQCats.eof
    lcl_selected_category = ""
    lcl_internal_label    = ""

    if CLng(iFAQCategoryId) = CLng(oFAQCats("FAQCategoryId")) then
       lcl_selected_category = " selected=""selected"""
    end if

  		if oFAQCats("internalonly") then
       lcl_internal_label = " (Internal)"
    end if

  		response.write "            <option value=""" & oFAQCats("FAQCategoryId") & """" & lcl_selected_category & ">" & oFAQCats("FAQCategoryName") & lcl_internal_label & "</option>" & vbcrlf

    oFAQCats.movenext
 loop

	response.write "          </select>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf

	oFAQCats.close
	set oFAQCats = nothing 

end sub

'------------------------------------------------------------------------------
sub displayButtons(iScreenMode)

  if UCASE(iScreenMode) = "EDIT" then
     lcl_button_label = "UPDATE"
  else
     lcl_button_label = "ADD"
  end if

  response.write "<p>" & vbcrlf
  response.write "  <input type=""button"" class=""button"" value=""" & lcl_button_label & " " & lcl_pagetitle & """ onclick=""validate();"" />" & vbcrlf
  response.write "</p>" & vbcrlf

end sub

'------------------------------------------------------------------------------
Function DBsafe( strDB )
 	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	 DBsafe = Replace( strDB, "'", "''" )
End Function
%>
