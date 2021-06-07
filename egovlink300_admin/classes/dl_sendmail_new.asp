<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: dl_sendmail_new.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 05/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  05/10/06  JOHN STULLENBERGER - INITIAL VERSION
' 1.2	 10/05/06	 Steve Loar - Security, Header and nav changed
' 1.3  01/28/08  David Boyer - Incorporated isFeatureOffline check
' 1.4  01/28/08  David Boyer - Incorporated Job/Bid Postings
' 1.5  11/20/08  David Boyer - Added check for "subscriptions_footer" org Edit-Display
' 1.6  06/26/09  David Boyer - Now saving subscription info when sent
' 1.7  06/29/09  David Boyer - Added default of subscription info from subscription log
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptions,job_postings,bid_postings") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 Dim bSent, iEmailFormat, sFromEmail, sFromName

 sLevel          = "../"     'Override of value from common.asp
 lcl_hidden      = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide
 lcl_onload      = ""
 lcl_list_type   = request("listtype")
 lcl_screen_mode = request("screen_mode")

 if request("iEmailFormat") = "" then
   	iEmailFormat = clng(2)
 else
   	iEmailFormat = clng(request("iEmailFormat"))
 end If 

'Check to see if the user has the proper permission
 if lcl_list_type = "JOB" then
    if not UserHasPermission( session("userid"), "job_postings" ) then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 elseif lcl_list_type = "BID" then
    if not UserHasPermission( session("userid"), "bid_postings" ) then
       	response.redirect sLevel & "permissiondenied.asp"
    end if
 else
    if not UserHasPermission( session("userid"), "send distribution emails" ) then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

'Check for POST.  If so process email
 if request.servervariables("REQUEST_METHOD") = "POST" then

    if request("MailList") <> "" then
       lcl_total_emails = getTotalEmails(request("MailList"))

       if lcl_total_emails <> "" then

          if lcl_total_emails > 100 then
             saveSubscriptionInfo lcl_list_type, _
                                  lcl_total_emails, _
                                  "SQL", _
                                  lcl_dl_logid


             'sendNotificationToEmailCreator

             sendEmail "dboyer28@yahoo.com", _
                       "dboyer@eclink.com", _
                       "", _
                       "SQL job sent", _
                       "blahblahblah", _
                       "", _
                       "Y"
          else
             saveSubscriptionInfo lcl_list_type, _
                                  lcl_total_emails, _
                                  "ASP", _
                                  lcl_dl_logid
            	subProcessEmail

             updateSubscriptionInfoStatus lcl_dl_logid, _
                                          "COMPLETED"
          end if
       end if
    end if

    if request("screen_mode") = "AUTOSEND" then
       'response.redirect "../job_bid_postings/" & session("return_url")
       'if request("success") = "SU" then
       'lcl_message = "<b style=""color:#FF0000"">*** Email Successfully Sent... ***</b>"
       lcl_onload = lcl_onload & "displayScreenMsg('Email Successfully Sent');"
       lcl_show_return_postings_link = "Y"
    end if

   	bSent = True
 else
   	bSent = False 
 end if

'-- Auto-populate fields if coming from a subscription log record -------------
 if request("dl_logid") <> "" then
   'Retrieve all of the data for the subscription log
    sSQL = "SELECT dl_logid, "
    sSQL = sSQL & " email_fromname, "
    sSQL = sSQL & " email_fromemail, "
    sSQL = sSQL & " email_subject, "
    sSQL = sSQL & " email_body, "
    sSQL = sSQL & " email_format, "
    sSQL = sSQL & " containsHTML, "
    sSQL = sSQL & " dl_listids "
    sSQL = sSQL & " FROM egov_class_distributionlist_log "
    sSQL = sSQL & " WHERE orgid = "  & session("orgid")
    sSQL = sSQL & " AND dl_logid = " & request("dl_logid")

   	set oGetLogInfo = Server.CreateObject("ADODB.Recordset")
    oGetLogInfo.Open sSQL, Application("DSN"), 3, 1
	
   	if not oGetLogInfo.eof then
       sFromName                = oGetLogInfo("email_fromname")
       sFromEmail               = oGetLogInfo("email_fromemail")
       lcl_subject_line         = oGetLogInfo("email_subject")
       lcl_html_body            = oGetLogInfo("email_body")
       iEmailFormat             = clng(oGetLogInfo("email_format"))
       lcl_dl_listids           = oGetLogInfo("dl_listids")
       lcl_checked_containsHTML = setupContainsHTMLCheckbox(oGetLogInfo("containsHTML"))
    end if

    oGetLogInfo.close
    set oGetLogInfo = nothing
 else
   '-- Set up the email form fields -------------------------------------------
   'From Email
    If request("sFromEmail") <> "" Then
      	sFromEmail = request("sFromEmail")
    Else
      	sFromEmail = GetInternalDefaultEmail( session("orgid") )  ' In common.asp
    End If 

   'From Name
    If request("sFromName") <> "" Then
      	sFromName = trim(request("sFromName"))
    Else
      	sFromName = GetInternalDefaultContact( session("orgid") )  ' In common.asp
    End If

   'Subject Line
    if request("sSubjectLine") <> "" then
       lcl_subject_line = request("sSubjectLine")
    else
       if lcl_screen_mode = "AUTOSEND" then
          lcl_subject_line = left(session("jobbid_id") & " [" & session("jobbid_title") & "]",150)
       else
          lcl_subject_line = ""
       end if
    end if

   'HTML Body
    if request("sHTMLBody") <> "" then
       lcl_html_body = request("sHTMLBody")
    else
       if lcl_screen_mode = "AUTOSEND" then
          if lcl_list_type = "JOB" then
             lcl_posting_display = "jobpostings_autosend_email_body"
          else
             lcl_posting_display = "bidpostings_autosend_email_body"
          end if

         'Check to see if the org has the "Edit Display" setup for the Posting Type (JOB or BID).
          if OrgHasDisplay( session("orgid"), lcl_posting_display ) then
             lcl_html_body = GetOrgDisplay( session("orgid"), lcl_posting_display )
          else
             lcl_html_body = "This complimentary message is being sent to opt-in subscribers that might be interested "
             lcl_html_body = lcl_html_body & "in its content.  If you do not wish to continue receiving these messages, "
             lcl_html_body = lcl_html_body & "please accept our apologies, and unsubscribe by following the instructions "
             lcl_html_body = lcl_html_body & "at the bottom of this message."
          end if
       else
          lcl_html_body = ""
       end if
    end if

   'Contains HTML checkbox
    lcl_checked_containsHTML = setupContainsHTMLCheckbox(request("containsHTML"))
 end if

'Determine which list type to display.
 if lcl_list_type = "JOB" then
    lcl_page_title          = "Job Posting"
    lcl_select_avail_title  = "Job Posting(s)"
 elseif lcl_list_type = "BID" then
    lcl_page_title          = "Bid Posting"
    lcl_select_avail_title  = "Bid Posting(s)"
 else
    lcl_page_title          = ""
    lcl_select_avail_title  = "Distribution List(s)"
 end if

'Check for org features
 lcl_orghasfeatures_default_fromemail_to_user = orghasfeature("default_fromemail_to_user")
%>
<html>
<head>
 	<title>E-Gov Administration Console {Send <%=lcl_select_avail_title%>}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
 	<link rel="stylesheet" type="text/css" href="classes.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

 	<script language="javascript" src="selectbox_script.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

 	<script language="javascript">
<!--
//function updatemaillistadd() {

		// ADD TO MAIL LIST DISPLAY
//  var sAddText  = document.frmlocation.SendList[document.frmlocation.SendList.selectedIndex].text;
//		var sAddValue = document.frmlocation.SendList[document.frmlocation.SendList.selectedIndex].value;	
//		addToList(frmlocation.AvailableList,sAddText,sAddValue);
		
		// REMOVE FROM AVAILABLE LIST
//		removeFromList(frmlocation.SendList);

		//UPDATE HIDDEN LIST
//		sMailList = document.frmlocation.maillist.value;
//		sMailList = sMailList.replace(sAddValue+'X','');
//		document.frmlocation.maillist.value = sMailList;
//}

//function updatemaillistremove() {
		// ADD TO AVAILABLE LIST
//		var sAddText  = document.frmlocation.AvailableList[document.frmlocation.AvailableList.selectedIndex].text;
//		var sAddValue = document.frmlocation.AvailableList[document.frmlocation.AvailableList.selectedIndex].value;	
//		addToList(frmlocation.SendList,sAddText,sAddValue);
		
		//  REMOVE FROM MAILING LIST
//		removeFromList(frmlocation.AvailableList);

		//UPDATE HIDDEN LIST
//		document.frmlocation.maillist.value = document.frmlocation.maillist.value + sAddValue + 'X';
//}

function updatemaillistadd() {

		// ADD TO MAIL LIST DISPLAY
  var sAddText  = document.getElementById("SendList")[document.getElementById("SendList").selectedIndex].text;
		var sAddValue = document.getElementById("SendList")[document.getElementById("SendList").selectedIndex].value;	
		addToList(document.getElementById("AvailableList"),sAddText,sAddValue);
		
		// REMOVE FROM AVAILABLE LIST
		removeFromList(document.getElementById("SendList"));

		//UPDATE HIDDEN LIST
		sMailList = document.getElementById("maillist").value;
		sMailList = sMailList.replace(sAddValue+'X','');
		document.getElementById("maillist").value = sMailList;
}

function updatemaillistremove() {
		// ADD TO AVAILABLE LIST
		var sAddText  = document.getElementById("AvailableList")[document.getElementById("AvailableList").selectedIndex].text;
		var sAddValue = document.getElementById("AvailableList")[document.getElementById("AvailableList").selectedIndex].value;	
		addToList(document.getElementById("SendList"),sAddText,sAddValue);
		
		//  REMOVE FROM MAILING LIST
		removeFromList(document.getElementById("AvailableList"));

		//UPDATE HIDDEN LIST
		document.getElementById("maillist").value = document.getElementById("maillist").value + sAddValue + 'X';
}

function doSitePicker(sFormField) {
		w = (screen.width - 350)/2;
		h = (screen.height - 350)/2;
		eval('window.open("../sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

//function doPicker(sFormField) {
//	 w = (screen.width - 350)/2;
//	 h = (screen.height - 350)/2;
//	 eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
//}

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
  w = 600;
  h = 400;
  l = (screen.availWidth/2)-(w/2);
  t = (screen.availHeight/2)-(h/2);

  lcl_showFolderStart = "";
  lcl_folderStart     = 0;
//  lcl_displayLinkText = "&displayLinkText=Y"

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

  if(lcl_folderStart > 0) {
     //lcl_showFolderStart = "&folderStart=published_documents";
     lcl_showFolderStart = "&folderStart=CITY_ROOT";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += "&returnAsHTMLLink=<%=lcl_returnAsHTMLLink%>";
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;
//  pickerURL += lcl_displayLinkText;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
		    var caretPos = textEl.caretPos;
   		 caretPos.text =
		    caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
  			 text + ' ' : text;
	 }
   else
   		 // Append the link to the textarea
    	 textEl.value = textEl.value + text;
	 }

function checkform()	{
  var lcl_return_false = 0;

   //Selected Distribution Lists
   if(document.getElementById("maillist").value == '') {
      document.getElementById("AvailableList").focus();
      inlineMsg(document.getElementById("AvailableList").id,'<strong>Required Field Missing: </strong> At least one distribution list must be selected',10,'AvailableList');
      lcl_return_false = lcl_return_false + 1;
   }else{
      clearMsg("AvailableList");
   }

   //Body
   if(document.getElementById("sHTMLBody").value == '') {
      document.getElementById("sHTMLBody").focus();
      inlineMsg(document.getElementById("sHTMLBody").id,'<strong>Required Field Missing: </strong> Email Body',10,'sHTMLBody');
      lcl_return_false = lcl_return_false + 1;
   }else{
      clearMsg("sHTMLBody");
   }

   //Subject
   if(document.getElementById("sSubjectLine").value == '') {
      document.getElementById("sSubjectLine").focus();
      inlineMsg(document.getElementById("sSubjectLine").id,'<strong>Required Field Missing: </strong> Subject',10,'sSubjectLine');
      lcl_return_false = lcl_return_false + 1;
   }else{
      clearMsg("sSubjectLine");
   }

   //From Name
   if(document.getElementById("sFromName").value == '') {
      document.getElementById("sFromName").focus();
      inlineMsg(document.getElementById("sFromName").id,'<strong>Required Field Missing: </strong> From Name',10,'sFromName');
      lcl_return_false = lcl_return_false + 1;
   }else{
      if(document.getElementById("sFromName").value.indexOf(",") > 0) {
         document.getElementById("sFromName").focus();
         inlineMsg(document.getElementById("sFromName").id,'<strong>Invalid Value: </strong> Cannot contain a comma.',10,'sFromName');
         lcl_return_false = lcl_return_false + 1;
      }else{
         clearMsg("sFromName");
      }
   }

<% if not lcl_orghasfeatures_default_fromemail_to_user then %>
   //From Email Address
   if(document.getElementById("sFromEmail").value == '') {
      document.getElementById("sFromEmail").focus();
      inlineMsg(document.getElementById("sFromEmail").id,'<strong>Required Field Missing: </strong> From Email Address',10,'sFromEmail');
      lcl_return_false = lcl_return_false + 1;
   }else{
      //Check for comma(s)
      if(document.getElementById("sFromEmail").value.indexOf(",") > 0) {
         document.getElementById("sFromEmail").focus();
         inlineMsg(document.getElementById("sFromEmail").id,'<strong>Invalid Value: </strong> Cannot contain a comma.',10,'sFromEmail');
         lcl_return_false = lcl_return_false + 1;
      }else{
         //Check for apostrophe(s)
//         if(document.getElementById("sFromEmail").value.indexOf("'") > 0) {
//            document.getElementById("sFromEmail").focus();
//            inlineMsg(document.getElementById("sFromEmail").id,'<strong>Invalid Value: </strong> Cannot contain apostrophe(s).',10,'sFromEmail');
//            lcl_return_false = lcl_return_false + 1;
//         }else{
            clearMsg("sFromEmail");
//         }
      }
			}
<% end if %>

   if(lcl_return_false > 0) {
      return false;
   }else{
    		//document.frmlocation.submit();
    		document.getElementById("frmlocation").submit();
   }

	}

	function clearform()
	{
//		document.frmlocation.sHTMLBody.value    = '';
//		document.frmlocation.sSubjectLine.value = '';
//		document.frmlocation.sFromName.value    = '';
//		document.frmlocation.sFromEmail.value   = '';
		document.getElementById("sHTMLBody").value    = '';
		document.getElementById("sSubjectLine").value = '';
		document.getElementById("sFromName").value    = '';
		document.getElementById("sFromEmail").value   = '';
	}

<%
  'If this is an AUTOSEND then we need to default the distribution list(s) in the "Selected" list
  'and remove then from the "Available" list
   if lcl_screen_mode = "AUTOSEND" OR lcl_dl_listids <> "" then

      if lcl_screen_mode = "AUTOSEND" then
         lcl_dl_listids = session("email_dlids")
      end if

      if lcl_dl_listids = "" then
         lcl_dl_listids = 0
      end if

      response.write "function setupAutoSendLists() {" & vbcrlf

      sSQL = "SELECT distributionlistid, distributionlistname "
      sSQL = sSQL & " FROM egov_class_distributionlist "
      sSQL = sSQL & " WHERE distributionlistid IN (" & lcl_dl_listids & ") "
      sSQL = sSQL & " AND orgid = " & session("orgid")

      set rsdl = Server.CreateObject("ADODB.Recordset")
      rsdl.Open sSQL, Application("DSN"), 0, 1

      if not rsdl.eof then
         do while not rsdl.eof
               'response.write "   addToList(frmlocation.SendList,""" & rsdl("distributionlistname") & """,""" & rsdl("distributionlistid") & """);" & vbcrlf
               'response.write "   sMailList = document.frmlocation.maillist.value;" & vbcrlf
               'response.write "   sMailList = sMailList + """ & rsdl("distributionlistid") & "X"";" & vbcrlf
               'response.write "   document.frmlocation.maillist.value = sMailList;" & vbcrlf

               response.write "   addToList(document.getElementById(""SendList""),""" & rsdl("distributionlistname") & """,""" & rsdl("distributionlistid") & """);" & vbcrlf
               response.write "   sMailList = document.getElementById(""maillist"").value;" & vbcrlf
               response.write "   sMailList = sMailList + """ & rsdl("distributionlistid") & "X"";" & vbcrlf
               response.write "   document.getElementById(""maillist"").value = sMailList;" & vbcrlf

            rsdl.movenext
         loop
      end if

      response.write " }" & vbcrlf
   end if
%>

function checkMsgLength() {

  lcl_length_subject = document.getElementById("sSubjectLine").value.length;
  lcl_length_body    = document.getElementById("sHTMLBody").value.length;

  document.getElementById("sCharacterCount").innerHTML = lcl_length_subject + lcl_length_body;
}

function addHTMLTag(p_tag) {
  var lcl_body = document.getElementById("sHTMLBody").value;

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

  document.getElementById("sHTMLBody").value = lcl_body;
}

function viewEmail() {
  var w = 850;
  var h = 400;
  var l = (screen.availWidth/2)-(w/2);
  var t = (screen.availHeight/2)-(h/2);

  var lcl_previewEmail_URL = '';
  var lcl_emailbody        = '';
  var lcl_containsHTML     = 'N';

  if(document.getElementById("sHTMLBody").value != '') {
     lcl_emailbody = document.getElementById("sHTMLBody").value;
  }

  if(document.getElementById("containsHTML").checked) {
     lcl_containsHTML = 'Y';
  }

  lcl_previewEmail_URL  = 'preview_email.asp';
  lcl_previewEmail_URL += '?listtype=<%=lcl_list_type%>';
  lcl_previewEmail_URL += '&emailbody='    + encodeURIComponent(lcl_emailbody);
  lcl_previewEmail_URL += '&containsHTML=' + encodeURIComponent(lcl_containsHTML);

  //  window.open(lcl_previewEmail_URL,"_blank");
  eval('window.open("' + lcl_previewEmail_URL + '", "_preview", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
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
<%
 'Check to see if this message has been automatically sent via Job/Bid Postings.
  if bSent then
     lcl_onload = lcl_onload & "displayScreenMsg('Message successfully sent to the selected subscribers.');"
  else
     if lcl_screen_mode = "AUTOSEND" OR lcl_dl_listids <> "" then
        lcl_onload = lcl_onload & "setupAutoSendLists();"
     else
        lcl_onload = ""
     end if
  end if

  lcl_onload = lcl_onload & "checkMsgLength();"

  response.write "<body onload=""" & lcl_onload & """>" & vbcrlf
  response.write "<form name=""frmlocation"" id=""frmlocation"" action=""dl_sendmail_new.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""maillist"" id=""maillist"" value="""" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""listtype"" id=""listtype"" value=""" & lcl_list_type & """ size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""screen_mode"" id=""screen_mode"" value=""" & lcl_screen_mode & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""iEmailFormat"" id=""iEmailFormat"" value=""2"" />" & vbcrlf

  ShowHeader sLevel
%>
<!-- #include file="../menu/menu.asp"--> 
<%
 'Determine if the org would like to default the email to the current user or allow the user to edit the value
  if lcl_orghasfeatures_default_fromemail_to_user then
     sFromEmail = getUserEmail(session("userid"))
     sFromName  = GetAdminName(session("userid"))

    'If the email is NULL then send the email to the org default email
     if sFromEmail = "" then
        sFromEmail = GetInternalDefaultEmail( session("orgid") )
     end if

     'response.write sFromEmail & vbcrlf
     'response.write "<input type=""hidden"" name=""sFromEmail"" id=""sFromEmail"" value=""" & sFromEmail & """ />" & vbcrlf

     lcl_fromemail_type    = "hidden"
     lcl_fromemail_value   = sFromEmail
     lcl_fromemail_display = sFromEmail & vbcrlf
  else
     lcl_fromemail_type    = "text"
     lcl_fromemail_value   = sFromEmail
     lcl_fromemail_display = ""

     'response.write "<input type=""text"" name=""sFromEmail"" id=""sFromEmail"" maxlength=""150"" size=""50"" value=""" & sFromEmail & """ onchange=""clearMsg('sFromEmail');"" />" & vbcrlf
  end if

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""margin-bottom:5px"">" & vbcrlf
  response.write "  <tr valign=""bottom"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "           <h5>Send Emails to " & lcl_page_title & " Subscribers</h5>" & vbcrlf

  if lcl_show_return_postings_link = "Y" then
     response.write "<br />" & vbcrlf
     response.write "<input type=""button"" name=""returnToPostingsButton"" id=""returnToPostingsButton"" class=""button"" value=""Return to " & lcl_page_title & """ onclick=""../job_bid_postings/job_bid_list.asp?" & session("return_url") & """ />" & vbcrlf
  end if

  response.write "       </td>" & vbcrlf
  response.write "       <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "   </tr>" & vbcrlf
  response.write " </table>" & vbcrlf

  response.write "<div id=""subscriptionshadow"">" & vbcrlf
  response.write "<table cellpadding=""5"" cellspacing=""0"" border=""0"" id=""subscription"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "		   	<th>Email Information</th>" & vbcrlf
  response.write "		</tr>" & vbcrlf
  response.write "		<tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "				      <table cellpadding=""3"" cellspacing=""0"" border=""0"">" & vbcrlf
  response.write "       					<tr>" & vbcrlf
  response.write "          						<td nowrap=""nowrap"">From Email Address:</td>" & vbcrlf
  response.write "           					<td>" & vbcrlf
  response.write                      lcl_fromemail_display
  response.write "                    <input type=""" & lcl_fromemail_type & """ name=""sFromEmail"" id=""sFromEmail"" maxlength=""150"" size=""50"" value=""" & lcl_fromemail_value & """ onchange=""clearMsg('sFromEmail');"" />" & vbcrlf
  response.write "           					</td>" & vbcrlf
  response.write "           	</tr>" & vbcrlf
  response.write "           	<tr>" & vbcrlf
  response.write "           					<td>From Name:</td>" & vbcrlf
  response.write "           	    <td>" & vbcrlf
  response.write "           	        <input type=""text"" name=""sFromName"" id=""sFromName"" maxlength=""150"" size=""50"" value=""" & sFromName & """ onchange=""clearMsg('sFromName');"" />" & vbcrlf
  response.write "           	    </td>" & vbcrlf
  response.write "           	</tr>" & vbcrlf
  response.write "           	<tr>" & vbcrlf
  response.write "           	    <td>Subject:</td>" & vbcrlf
  response.write "           	    <td>" & vbcrlf
  response.write "           	        <input type=""text"" name=""sSubjectLine"" id=""sSubjectLine"" maxlength=""150"" size=""80"" value=""" & lcl_subject_line & """ onchange=""clearMsg('sSubjectLine');"" onkeyup=""checkMsgLength();"" />" & vbcrlf
  response.write "           	    </td>" & vbcrlf
  response.write "           	</tr>" & vbcrlf
  response.write "       					<tr>" & vbcrlf
  response.write "          						<td colspan=""2"" valign=""top"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""5"">" & vbcrlf
  response.write "                      <tr valign=""bottom"">" & vbcrlf
  response.write "                          <td>" & vbcrlf
  response.write "                              Email Body:&nbsp;<input class=""button"" type=""button"" onclick=""clearform();"" value=""Clear Form"" />" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                          <td align=""right"">" & vbcrlf
  response.write "                              <input type=""checkbox"" name=""containsHTML"" id=""containsHTML"" value=""Y""" & lcl_checked_containsHTML & " />&nbsp;Customized HTML Email Body (contains HTML tags: HTML, HEAD, and/or BODY)&nbsp;" & vbcrlf
  response.write "                              <input type=""button"" value=""Add a Link"" class=""button"" onClick=""doPicker('frmlocation.sHTMLBody','Y','Y','Y','Y');"" />" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                      <tr valign=""bottom"">" & vbcrlf
  response.write "                          <td align=""right"" colspan=""2"">" & vbcrlf
  response.write "                              <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                                <tr>" & vbcrlf
  response.write "                                    <td width=""40%"" align=""right"" nowrap=""nowrap""><strong>Common HTML formatting tags:</strong></td>" & vbcrlf
                                                      displayHTMLButton "Y", "Bold"
                                                      displayHTMLButton "Y", "Italics"
                                                      displayHTMLButton "Y", "FONT"
                                                      displayHTMLButton "Y", "H1"
                                                      displayHTMLButton "Y", "H2"
                                                      displayHTMLButton "Y", "H3"
                                                      displayHTMLButton "Y", "LINK"
                                                      displayHTMLButton "Y", "IMG"
                                                      displayHTMLButton "Y", "BR"
                                                      displayHTMLButton "Y", " P "
  response.write "                                    <td align=""center"">" & vbcrlf
  response.write "                                        Alignment:<br />" & vbcrlf
  response.write "                                        <select name=""p_format_alignment"" onchange=""addHTMLTag(this.value);"">" & vbcrlf
  response.write "                                          <option value=""""></option>" & vbcrlf
  response.write "                                          <option value=""P_LEFT"">LEFT</option>" & vbcrlf
  response.write "                                          <option value=""P_CENTER"">CENTER</option>" & vbcrlf
  response.write "                                          <option value=""P_RIGHT"">RIGHT</option>" & vbcrlf
  response.write "                                        </select>" & vbcrlf
  response.write "                                    </td>" & vbcrlf
  response.write "                                    <td nowrap=""nowrap"">[<a href=""http://www.w3schools.com/tags/default.asp"" target=""_blank"">Additional TAGs</a>]</td>" & vbcrlf
  response.write "                                </tr>" & vbcrlf
  response.write "                              </table>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td colspan=""2"">" & vbcrlf
  response.write "                              <textarea name=""sHTMLBody"" id=""sHTMLBody"" style=""width:800px; height:200px;"" onchange=""clearMsg('sHTMLBody');"" onkeyup=""checkMsgLength();"">" & lcl_html_body & "</textarea>" & vbcrlf
  response.write "                              <div style=""color:#ff0000;"">" & vbcrlf
  response.write "                                <input type=""button"" name=""previewEmailButton"" id=""previewEmailButton"" value=""Preview Email"" class=""button"" onclick=""viewEmail()"" />" & vbcrlf
  response.write "                                Total Character Count <em>(includes Subject and Email Body)</em>: [<span id=""sCharacterCount""></span>]" & vbcrlf
  response.write "                              </div>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "          						</td>" & vbcrlf
  response.write "       					</tr>" & vbcrlf
  response.write "        				<tr>" & vbcrlf
  response.write "           					<td colspan=""2"">Select Groups to receive this email:<br /><br />" & vbcrlf
  response.write "              						<table cellpadding=""10"" cellspacing=""0"" border=""0"">" & vbcrlf
  response.write "               							<tr>" & vbcrlf
  response.write "                  								<td>" & vbcrlf
  response.write "                     									<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
  response.write "                      										<tr>" & vbcrlf
  response.write "                                    <td height=""20""><strong>Selected " & lcl_select_avail_title & "</strong></td>" & vbcrlf
  response.write "                                </tr>" & vbcrlf
  response.write "                      										<tr>" & vbcrlf
  response.write "                         											<td>" & vbcrlf
  response.write "                            												<select size=""15"" style=""width:370px"" name=""SendList"" id=""SendList"">" & vbcrlf
  response.write "                            												</select>" & vbcrlf
  response.write "                         											</td>" & vbcrlf
  response.write "                      										</tr>" & vbcrlf
  response.write "                     									</table>" & vbcrlf
  response.write "                  								</td>" & vbcrlf
  response.write "                   							<td align=""center"">" & vbcrlf
  response.write "                              <a href=""#""><img src=""../images/ieforward.gif"" name=""removeDLList"" id=""removeDLList"" align=""absmiddle"" border=""0"" onClick=""updatemaillistadd();"" class=""hotspot"" onmouseover=""tooltip.show('Click to REMOVE Distribution List(s)');"" onmouseout=""tooltip.hide();"" /></a>" & vbcrlf
  response.write "                              <p>" & vbcrlf
  response.write "                              <a href=""#""><img src=""../images/ieback.gif"" name=""addDLList"" id=""addDLList"" align=""absmiddle"" border=""0"" onClick=""clearMsg('AvailableList');updatemaillistremove();"" class=""hotspot"" onmouseover=""tooltip.show('Click to ADD Distribution List(s)');"" onmouseout=""tooltip.hide();"" /></a>" & vbcrlf
  response.write "                  								</td>" & vbcrlf
  response.write "                  								<td>" & vbcrlf
  response.write "                     									<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
  response.write "                       									<tr>" & vbcrlf
  response.write "                                    <td height=""20""><strong>Available " & lcl_select_avail_title & "</strong></td>" & vbcrlf
  response.write "                                </tr>" & vbcrlf
  response.write "                      										<tr>" & vbcrlf
  response.write "                         											<td>" & vbcrlf
  response.write "                            												<select size=""15"" style=""width:370px"" name=""AvailableList"" id=""AvailableList"">" & vbcrlf
                                                            ListDistributionLists request("listtype"),lcl_screen_mode
  response.write "                            												</select>" & vbcrlf
  response.write "                         											</td>" & vbcrlf
  response.write "                      										</tr>" & vbcrlf
  response.write "                     									</table>" & vbcrlf
  response.write "                  								</td>" & vbcrlf
  response.write "                 					</tr>" & vbcrlf
  response.write "                				</table>" & vbcrlf
  response.write "          						</td>" & vbcrlf
  response.write "       					</tr>" & vbcrlf
  response.write "       					<tr>" & vbcrlf
  response.write "           					<td colspan=""2"">" & vbcrlf
  response.write "               					<input type=""button"" name=""sendMessageButton"" id=""sendMessageButton"" class=""button"" onclick=""checkform();"" value=""Send Your Message"" />" & vbcrlf
  response.write "           					</td>" & vbcrlf
  response.write "        				</tr>" & vbcrlf
  response.write "      				</table>" & vbcrlf
  response.write "   			</td>" & vbcrlf
  response.write "		</tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</form>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Sub ListDistributionLists( p_list_type, p_screen_mode )
	Dim sSql, oList, rs2

	'GET ALL DISTRIBUTION LISTS FOR ORG
	sSql = "SELECT * "
	sSql = sSql & " FROM egov_class_distributionlist "
	sSql = sSql & " WHERE orgid = " & session("orgid")

	If p_list_type <> "" Then 
		sSql = sSql & " AND distributionlisttype = '" & p_list_type & "' "
	Else 
		sSql = sSql & " AND (distributionlisttype is null OR distributionlisttype = '') "
	End If 
	sSql = sSql & " AND (parentid = '' OR parentid IS NULL) "
	sSql = sSql & " ORDER BY distributionlistname"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSql, Application("DSN"), 0, 1

	If NOT oList.EOF Then

	  'LOOP THRU AND DISPLAY ROWS
 	  Do While Not oList.EOF

		'If the screen_mode = "AUTOSEND" then do not display any of the distribution lists
		If p_screen_mode = "AUTOSEND" Then 
			If checkForDLID(oList("distributionlistid")) = "N" Then 
				response.write "<option value=""" & oList("distributionlistid") & """>" &  oList("distributionlistname") & "</option>" & vbcrlf
			End If 
		Else 
			response.write "<option value=""" & oList("distributionlistid") & """>" &  oList("distributionlistname") & "</option>" & vbcrlf
		End if 

		'Check for any sub-categories
		sSql = "SELECT * "
		sSql = sSql & " FROM egov_class_distributionlist "
		sSql = sSql & " WHERE orgid = " & session("orgid")
		sSql = sSql & " AND parentid = " & oList("distributionlistid")
		sSql = sSql & " AND distributionlisttype = '" & p_list_type & "'"
		sSql = sSql & " ORDER BY distributionlistname"

		Set rs2  = Server.CreateObject("ADODB.Recordset")
		rs2.Open sSql, Application("DSN"), 0, 1

		If Not rs2.EOF Then  
			Do While Not rs2.EOF
				'If the screen_mode = "AUTOSEND" then do not display any of the distribution lists
				If p_screen_mode = "AUTOSEND" Then 
					If checkForDLID(rs2("distributionlistid")) = "N" Then 
						response.write "<option value=""" & rs2("distributionlistid") & """>&nbsp;&nbsp;&nbsp;" &  rs2("distributionlistname") & "</option>" & vbcrlf
					End If 
				Else 
					response.write "<option value=""" & rs2("distributionlistid") & """>&nbsp;&nbsp;&nbsp;" &  rs2("distributionlistname") & "</option>" & vbcrlf
				End If 
				rs2.MoveNext
			Loop 
		End If 
		oList.MoveNext
		Loop 

	Else
		' NO DISTRIBUTION LISTS WERE FOUND
	End If

	oList.Close
	Set oList = Nothing 

End Sub

'------------------------------------------------------------------------------
sub saveSubscriptionInfo(ByVal iListType, ByVal iTotalEmails, ByVal iEmailSendType, ByRef lcl_dl_logid)

 'Validate the fields
  sListType           = "NULL"
  sTotalEmails        = 0
  sEmailSendType      = ""
  lcl_sentdate        = "'" & now & "'"
  lcl_completedate    = "NULL"
  lcl_sendstatus      = "'INPROGRESS'"
  lcl_email_fromname  = "NULL"
  lcl_email_fromemail = "NULL"
  lcl_email_subject   = "NULL"
  lcl_email_body      = "NULL"
  lcl_email_format    = "NULL"
  sdllist             = "NULL"
  lcl_containsHTML    = 0

  if lcl_list_type <> "" then
     sListType = "'" & UCASE(dbsafe(iListType)) & "'"
  end if

  if iTotalEmails <> "" then
     if not containsApostrophe(iTotalEmails) then
        sTotalEmails = clng(iTotalEmails)
     end if
  end if

  if iEmailSendType <> "" then
     if not containsApostrophe(iEmailSendType) then
        sEmailSendType = ucase(iEmailSendType)
        sEmailSendType = dbsafe(sEmailSendType)
        sEmailSendType = "'" & sEmailSendType & "'"
     end if
  end if

  if request("sFromName") <> "" then
     lcl_email_fromname = "'" & dbsafe(request("sFromName")) & "'"
  end if

  if request("sFromEmail") <> "" then
     lcl_email_fromemail = "'" & dbsafe(request("sFromEmail")) & "'"
  end if

  if request("sSubjectLine") <> "" then
     lcl_email_subject = "'" & dbsafe(request("sSubjectLine")) & "'"
  end if

  if request("sHTMLBody") <> "" then
     lcl_email_body = "'" & dbsafe(request("sHTMLBody")) & "'"
  end if

  if request("iEmailFormat") <> "" then
     lcl_email_format = "'" & dbsafe(request("iEmailFormat")) & "'"
  end if

 'Get the distribution list(s)
  if request("MailList") <> "" then
    	sdllist = replace(request("Maillist"),"X",",")  'Build comma separate list
    	iLength = LEN(sdllist)
    	sdllist = LEFT(sdllist,(iLength - 1))  'Trim trailing comma
     sdllist = "'" & dbsafe(sdllist) & "'"
  end if

  if request("containsHTML") = "Y" then
     lcl_containsHTML = 1
  end if

  sSQL = "INSERT INTO egov_class_distributionlist_log ("
  sSQL = sSQL & "orgid, "
  sSQL = sSQL & "distributionlisttype, "
  sSQL = sSQL & "sentbyuserid, "
  sSQL = sSQL & "sentdate, "
  sSQL = sSQL & "completedate, "
  sSQL = sSQL & "sendstatus, "
  sSQL = sSQL & "email_fromname, "
  sSQL = sSQL & "email_fromemail, "
  sSQL = sSQL & "email_subject, "
  sSQL = sSQL & "email_body, "
  sSQL = sSQL & "email_format, "
  sSQL = sSQL & "containsHTML, "
  sSQL = sSQL & "dl_listids, "
  sSQL = sSQL & "total_emails_to_send, "
  sSQL = sSQL & "emailSendType "
  sSQL = sSQL & ") VALUES ("
  sSQL = sSQL & session("orgid")    & ", "
  sSQL = sSQL & sListType           & ", "
  sSQL = sSQL & session("userid")   & ", "
  sSQL = sSQL & lcl_sentdate        & ", "
  sSQL = sSQL & lcl_completedate    & ", "
  sSQL = sSQL & lcl_sendstatus      & ", "
  sSQL = sSQL & lcl_email_fromname  & ", "
  sSQL = sSQL & lcl_email_fromemail & ", "
  sSQL = sSQL & lcl_email_subject   & ", "
  sSQL = sSQL & lcl_email_body      & ", "
  sSQL = sSQL & lcl_email_format    & ", "
  sSQL = sSQL & lcl_containsHTML    & ", "
  sSQL = sSQL & sdllist             & ", "
  sSQL = sSQL & sTotalEmails        & ", "
  sSQL = sSQL & sEmailSendType
  sSQL = sSQL & ")"

  lcl_dl_logid = RunInsertStatement(sSQL)


end sub

'------------------------------------------------------------------------------
sub updateSubscriptionInfoStatus(iDLLogID, iStatus)

  sSQL = "UPDATE egov_class_distributionlist_log SET "
  sSQL = sSQL & " sendstatus = '"   & UCASE(dbsafe(iStatus)) & "', "
  sSQL = sSQL & " completedate = '" & now                    & "' "
  sSQL = sSQL & " WHERE dl_logid = " & iDLLogID

	 set oDLLogUpdate = Server.CreateObject("ADODB.Recordset")
		oDLLogUpdate.Open sSQL, Application("DSN"), 0, 1

  set oDLLogUpdate = nothing

end sub

'------------------------------------------------------------------------------
sub subProcessEmail()
	dim iRowCount

	iRowCount = CLng(0) 

'Get the distribution list(s)
	sdllist = replace(request("Maillist"),"X",",") ' BUILD COMMA SEPARATE LIST
	iLength = LEN(sdllist)
	sdllist = LEFT(sdllist,(iLength - 1)) ' TRIM TRAILING COMMA

	'GET LIST OF UNIQUE EMAIL ADDRESSES
	sSQL = "SELECT DISTINCT userid, useremail "
	sSQL = sSQL & " FROM egov_dl_user_list "
	sSQL = sSQL & " WHERE (distributionlistid IN (" & sdllist & ")) "
 sSQL = sSQL & " ORDER BY userid "

 set oEmail = Server.CreateObject("ADODB.Recordset")
 oEmail.Open sSQL, Application("DSN"), 0, 1

'Loop thru email addresses sending email
	do while not oEmail.eof
  		iRowCount = iRowCount + CLng(1)

 		'This "If" is to handle skokie crashes. Change the user id found in the log to pick up from the crash point
    'If ( CLng(session("orgid")) = CLng(131) And CLng(oEmail("userid")) > CLng(182789)) Or (CLng(session("orgid")) <> CLng(131)) Then 

			'SEND EMAIL		-- isValidEmail() is in common.asp
 			if not IsNull( oEmail("useremail") ) then
  				'AddToLog "-----------------------------------------------------------------------------------------"
   				'AddToLog "Orgid: " & session("orgid")
		   		AddToLog "Sending # " & iRowCount
   				AddToLog "From: "     & request("sFromName") & "[" & request("sFromEmail") & "]"
			   	AddToLog "Subject:"   & request("sSubjectLine")
   				AddToLog "To: "       & LCase(oEmail("useremail")) & " Userid: " & oEmail("userid")

   				if isValidEmail( LCase(oEmail("useremail")) ) then
  	   				subSendEmail request("sFromName"), request("sFromEmail"), request("sSubjectLine"), request("sHTMLBody"), LCase(oEmail("useremail")), oEmail("userid"), clng(request("iEmailFormat")), request("containsHTML"), sdllist
      				AddToLog "Successful Send"
       else
  	   				AddToLog "***** Email not sent due to invalid email format *****"
       end if
  				'AddtoLog "-----------------------------------------------------------------------------------------"
 			end if
    'End If 
   	oEmail.movenext
	loop

	oEmail.Close
	set oEmail = nothing 

end sub


'------------------------------------------------------------------------------
sub subSendEmail( sFromName, sFromEmail, sSubjectLine, sHTMLBody, sToEmail, iUserId, iEmailFormat, iContainsHTML, iDLListID )
 	dim sEmailFooter

	 sEmailFooter = ""

  lcl_email_from     = sFromName & " <" & sFromEmail & ">"
  lcl_email_to       = sToEmail
  lcl_email_subject  = sSubjectLine

 	if not userisrootadmin(session("userid")) then
     lcl_footer = ""

 		  sEmailFooter = vbcrlf & vbcrlf & vbcrlf
     sEmailFooter = sEmailFooter & "<font style=""font-size:10pt;"">" & vbcrlf
     sEmailFooter = sEmailFooter & "<p>" & vbcrlf

    'Check to see if the org has overridden the default subscription footer
     if OrgHasDisplay(session("orgid"),"subscriptions_footer") then
        lcl_footer = GetOrgDisplay(session("orgid"),"subscriptions_footer")
        lcl_footer = checkForCustomFields(lcl_footer,iUserID,sToEmail)
     end if

     if trim(lcl_footer) <> "" then
        sEmailFooter = sEmailFooter & lcl_footer
     else
        'Get the name of the list(s) to be unsubscribed from
         lcl_display_dlids   = ""
         lcl_display_dlnames = ""

         'sSQL = "SELECT distributionlistname "
         'sSQL = sSQL & " FROM egov_class_distributionlist "
         'sSQL = sSQL & " WHERE distributionlistid IN (" & iDLListID& ") "
         'sSQL = sSQL & " ORDER BY distributionlistname "

         sSQL = "SELECT dl.distributionlistid, "
         sSQL = sSQL & " dl.distributionlistname "
         sSQL = sSQL & " FROM egov_class_distributionlist dl, "
         sSQL = sSQL &      " egov_class_distributionlist_to_user dltu "
         sSQL = sSQL & " WHERE dl.distributionlistid = dltu.distributionlistid "
         sSQL = sSQL & " AND dl.distributionlistid IN (" & iDLListID& ") "
         sSQL = sSQL & " AND dltu.userid = " & iUserID
         sSQL = sSQL & " ORDER BY dl.distributionlistname "

         set oGetDLNames = Server.CreateObject("ADODB.Recordset")
         oGetDLNames.Open sSQL, Application("DSN"), 0, 1

         if not oGetDLNames.eof then
            do while not oGetDLNames.eof

              'Build the "display" distribution list IDs (for the unsubscribe URL)
               if lcl_display_dlids = "" then
                  lcl_display_dlids = oGetDLNames("distributionlistid")
               else
                  lcl_display_dlids = lcl_display_dlids & "," & oGetDLNames("distributionlistid")
               end if

              'Build the "display" distribution list names
               if lcl_display_dlnames = "" then
                  lcl_display_dlnames = oGetDLNames("distributionlistname")
               else
                  lcl_display_dlnames = lcl_display_dlnames & ", " & oGetDLNames("distributionlistname")
               end if

               oGetDLNames.movenext
            loop
         end if

         oGetDLNames.close
         set oGetDLNames = nothing

        'sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " sent this e-mail to you because your Notification Preferences "
        sEmailFooter = sEmailFooter & session("sOrgName") & " sent this e-mail to you because your Notification Preferences "
        sEmailFooter = sEmailFooter & "indicate that you want to receive information from us. We will not request personal data (password, "
        sEmailFooter = sEmailFooter & "credit card/bank numbers) in an e-mail. You are subscribed as " & sToEmail & ", "
        'sEmailFooter = sEmailFooter & "registered on " & session("egovclientwebsiteurl") & "."
        'sEmailFooter = sEmailFooter & "registered on " & session("sOrgName") & " (" & session("egovclientwebsiteurl") & ")."
        sEmailFooter = sEmailFooter & "registered on " & session("sOrgName") & " (<a href=""" & session("egovclientwebsiteurl") & """>" & session("egovclientwebsiteurl") & "</a>)."
        sEmailFooter = sEmailFooter & "</p>"
        sEmailFooter = sEmailFooter & "<p>"
        sEmailFooter = sEmailFooter & "<strong>Click Here to Unsubscribe From this List(s):</strong>" & vbcrlf
        'sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId & "&dl=" & iDLListID
        sEmailFooter = sEmailFooter & "You will be removed from the following lists: " & lcl_display_dlnames & "<br />"
        sEmailFooter = sEmailFooter & "<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId & "&dl=" & iDLListID & """>" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId & "&dl=" & lcl_display_dlids & "</a>.<br />" & vbcrlf
        sEmailFooter = sEmailFooter & "</p>"
        sEmailFooter = sEmailFooter & "<p>"
        sEmailFooter = sEmailFooter & "<strong>Manage Subscriptions:</strong>" & vbcrlf
        sEmailFooter = sEmailFooter & "If you do not wish to receive further communications, or you wish to view "
        sEmailFooter = sEmailFooter & "and/or modify which lists you are subscribed to, simply click the link below.<br />"
        'sEmailFooter = sEmailFooter & "If you do not wish to receive further communications, or you wish "
        'sEmailFooter = sEmailFooter & "to view and or modify which lists you are subscribed to, sign "
        'sEmailFooter = sEmailFooter & "into " & session("sOrgName") & " by clicking on the ""Login"" link "
        'sEmailFooter = sEmailFooter & "found at the bottom of the home page and change your Notification "
        'sEmailFooter = sEmailFooter & "Preferences or simply click the link below.<br />"
        sEmailFooter = sEmailFooter & "<a href=""" & session("egovclientwebsiteurl") & "/manage_mail_lists.asp"">" & session("egovclientwebsiteurl") & "/manage_mail_lists.asp</a>"
     end if

     sEmailFooter = sEmailFooter & "</p>"
     sEmailFooter = sEmailFooter & "</font>"
 	end if

 'Build the email body
 	sHTMLBody = sHTMLBody & sEmailFooter

  if iContainsHTML = "" then
     iContainsHTML = "N"
  end if

'		if iEmailFormat < 3 then

 			'include a plain text body
'  			if not UserIsRootAdmin(session("userid")) then
'	     		sHTMLBody = sHTMLBody & vbcrlf & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId & "&dl=" & iDLListID
' 		 	end if

	  	 '.TextBody = clearHTMLTags(sHTMLBody)  'This is in common.asp
' 	end if

  lcl_email_htmlbody = BuildHTMLMessage(sHTMLBody, iContainsHTML)

 'Remove the name from the email address
  lcl_validate_email = formatSendToEmail(sFromEmail)

 'The function isValidEmail (found in common.asp) allows an email to simply have an "@" sign at the end of the email.
 'However, this will crash the application.  Check to see if the last character in the email entered is an "@".
  if lcl_validate_email <> "" AND RIGHT(lcl_validate_email,1) <> "@" then
     if isValidEmail(lcl_validate_email) then

       'Send the email if it is valid.
        if iEmailFormat = 1 then
           sendEmail lcl_email_from, lcl_email_to, "", lcl_email_subject, "", lcl_email_htmlbody, "Y"
        else
           sendEmail lcl_email_from, lcl_email_to, "", lcl_email_subject, lcl_email_htmlbody, "", "Y"
        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
'Sub subSendEmail( sFromName, sFromEmail, sSubjectLine, sHTMLBody, sToEmail, iUserId, iEmailFormat )
'	Dim sEmailFooter, oCdoMail, oCdoConf

'	sEmailFooter = ""

	' CREATE MAIL OBJECT
'	Set oCdoMail = Server.CreateObject("CDO.Message")
'	Set oCdoConf = Server.CreateObject("CDO.Configuration")

'	With oCdoConf
'		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")  = 2
'		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("SMTP_Server")
'		.Fields.Update
'	End With

'	With oCdoMail
' 		Set .Configuration = oCdoConf
' 		.From = sFromName & " <" & sFromEmail & ">"
'	 	.To   = sToEmail
'	End With 

'	if not userisrootadmin(session("userid")) then
'    lcl_footer = ""

'		  sEmailFooter = vbcrlf & vbcrlf & vbcrlf
'    sEmailFooter = sEmailFooter & "<font style=""font-size:10pt;"">" & vbcrlf
'    sEmailFooter = sEmailFooter & "<p>" & vbcrlf

   'Check to see if the org has overridden the default subscription footer
'    if OrgHasDisplay(session("orgid"),"subscriptions_footer") then
'       lcl_footer = GetOrgDisplay(session("orgid"),"subscriptions_footer")
'       lcl_footer = checkForCustomFields(lcl_footer,iUserID,sToEmail)
'    end if

'    if trim(lcl_footer) <> "" then
'       sEmailFooter = sEmailFooter & lcl_footer
'    else
'       sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " sent this e-mail to you because your Notification Preferences " & vbcrlf
'       sEmailFooter = sEmailFooter & "indicate that you want to receive information from us. We will not request personal data (password, " & vbcrlf
'       sEmailFooter = sEmailFooter & "credit card/bank numbers) in an e-mail. You are subscribed as " & sToEmail & ", " & vbcrlf
'       sEmailFooter = sEmailFooter & "registered on " & session("egovclientwebsiteurl") & "." & vbcrlf
'       sEmailFooter = sEmailFooter & "</p>" & vbcrlf
'       sEmailFooter = sEmailFooter & "<p>" & vbcrlf
'       sEmailFooter = sEmailFooter & "If you do not wish to receive further communications, sign into " & vbcrlf
'       sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " by clicking on the ""Login"" link found at the bottom of the " & vbcrlf
'       sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " home page and change your Notification Preferences or click " & vbcrlf
'       sEmailFooter = sEmailFooter & "the link below to unsubscribe from this mailing list." & vbcrlf
'       sEmailFooter = sEmailFooter & "</p>" & vbcrlf
'       sEmailFooter = sEmailFooter & "<p>" & vbcrlf
'       sEmailFooter = sEmailFooter & "<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId & """>Click Here to Unsubscribe From our Mailing Lists</a>." & vbcrlf
'    end if

'    sEmailFooter = sEmailFooter & "</p>" & vbcrlf
'    sEmailFooter = sEmailFooter & "</font>" & vbcrlf
'	end if

'	sHTMLBody = sHTMLBody & sEmailFooter

'	With oCdoMail
' 		.Subject = sSubjectLine
'	 	If iEmailFormat > 1 Then 
		  	'include an HTML body
'   			.HTMLBody = sHTMLBody
' 		End If 

' 		if iEmailFormat < 3 then
  			'include a plain text body
'   			if not UserIsRootAdmin(session("userid")) then
'		     		sHTMLBody = sHTMLBody & vbcrlf & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId 
'  		 	end if
'		  	 .TextBody = clearHTMLTags(sHTMLBody)  'This is in common.asp
' 		end if
'	end with

'	oCdoMail.Configuration = oCdoConf
'	oCdoMail.Send

'	set oCdoConf = nothing
'	set oCdoMail = nothing

'End Sub


'------------------------------------------------------------------------------
function getPostingsName(p_value)
  sSqlp = "SELECT title "
  sSqlp = sSqlp & " FROM egov_jobs_bids "
  sSqlp = sSqlp & " WHERE orgid = "    & session("orgid")
  sSqlp = sSqlp & " AND posting_id = " & p_value

  set rsp = Server.CreateObject("ADODB.Recordset")
  rsp.Open sSqlp, Application("DSN"), 0, 1

  if not rsp.eof then
     lcl_return = rsp("title")
  else
     lcl_return = ""
  end if

  getPostingsName = lcl_return

  set rsp = nothing

end function


'------------------------------------------------------------------------------
function getCategoryName(p_value)
  sSqlc = "SELECT distributionlistname "
  sSqlc = sSqlc & " FROM egov_class_distributionlist "
  sSqlc = sSqlc & " WHERE distributionlistid = " & p_value
  sSqlc = sSqlc & " AND orgid = " & session("orgid")

  set rsc = Server.CreateObject("ADODB.Recordset")
  rsc.Open sSqlc, Application("DSN"), 0, 1

  if not rsc.eof then
     lcl_return = rsc("distributionlistname")
  else
     lcl_return = ""
  end if

  getCategoryName = lcl_return

  set rsc = nothing

end function


'------------------------------------------------------------------------------
function checkForDLID(p_value)
  if session("email_dlids") <> "" then
     lcl_email_dlids = session("email_dlids")
  else
     lcl_email_dlids = 0
  end if

  sSql1 = "SELECT distributionlistid "
  sSql1 = sSql1 & " FROM egov_class_distributionlist "
  sSql1 = sSql1 & " WHERE distributionlistid IN (" & lcl_email_dlids & ") "
  sSql1 = sSql1 & " AND orgid = " & session("orgid")
  sSql1 = sSql1 & " AND distributionlistid = " & p_value

  set rs1 = Server.CreateObject("ADODB.Recordset")
  rs1.Open sSql1, Application("DSN"), 0, 1

  if not rs1.eof then
     lcl_return = "Y"
  else
     lcl_return = "N"
  end if

  checkForDLID = lcl_return

end function


'------------------------------------------------------------------------------
function checkForCustomFields(p_value,p_userid,p_useremail)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,"<<USER_EMAIL>>",p_useremail)
     lcl_return = replace(lcl_return,"<<ORGWEBSITE>>",session("egovclientwebsiteurl"))
     lcl_return = replace(lcl_return,"<<UNSUBSCRIBE>>","<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & p_userid & """>Click Here to Unsubscribe From our Mailing Lists</a>")

    'Now check to see if they have a custom "Unsubscribe" text
    'The variable to be used in the "Edit Display" is: <<UNSUBSCRIBE_text goes here_UNSUBSCRIBE_END>>
     lcl_unsubscribe_start = "N"
     lcl_unsubscribe_end   = "N"

    'First check for the start of the "unsubscribe"
     if instr(lcl_return,"<<UNSUBSCRIBE_") > 0 then
        lcl_unsubscribe_start = "Y"

       'If the "start" exists then check for the end of the "unsubscribe"
        if instr(lcl_return,"_UNSUBSCRIBE_END>>") > 0 then
           lcl_unsubscribe_end = "Y"
        end if
     end if

    'If both the start and end of the "unsubscribe" exist then we can format them out
    'and build the unsubscribe link around the custom text.
     if lcl_unsubscribe_start = "Y" AND lcl_unsubscribe_end = "Y" then
        lcl_return = replace(lcl_return,"<<UNSUBSCRIBE_","<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & p_userid & """>")
        lcl_return = replace(lcl_return,"_UNSUBSCRIBE_END>>","</a>")
     end if
  end if

  checkForCustomFields = lcl_return

end function

'----------------------------------------------------------------------------------------
Sub AddtoLog( sText )
    ' WRITES SUPPLIED TEXT TO FILE WITH DATETIME
'	Set oFSO = Server.Createobject("Scripting.FileSystemObject")
'    Set oFile = oFSO.GetFile(Application("SubscriptionLog"))
'    Set oText = oFile.OpenAsTextStream(8)
'    oText.WriteLine (Now() & Chr(9) & sText)
'    oText.Close
    
'    Set oText = Nothing
'    Set oFile = Nothing
 '   Set oFSO = Nothing

	Dim sSql

	sSql = "INSERT INTO subscriptionlog ( orgid, logentry ) VALUES ( " & session("orgid") & ", '" & replace(sText,"'","''") & "' )"
	RunSQLStatement( sSql )

End Sub 

'------------------------------------------------------------------------------
function setupContainsHTMLCheckbox(iContainsHTML)

  lcl_return = ""

  if iContainsHTML <> "" then
     if iContainsHTML = "Y" then
        lcl_return = " checked=""checked"""
     elseif iContainsHTML then
        lcl_return = " checked=""checked"""
    end if
  end if

  setupContainsHTMLCheckbox = lcl_return

end function

'------------------------------------------------------------------------------
sub displayHTMLButton(iIncludeTDTags, p_name)

  lcl_includeTDTags = false

  if iIncludeTDTags <> "" then
     if iIncludeTDTags = "Y" then
        lcl_includeTDTags = true
     end if
  end if

  if p_name <> "" then
     if lcl_includeTDTags then
        response.write "<td>" & vbcrlf
     end if

     response.write "<input type=""button"" name=""" & p_name & "Button"" id=""" & p_name & "Button"" value=""" & p_name & """ class=""button"" onclick=""addHTMLTag('" & ucase(trim(p_name)) & "');"" />" & vbcrlf

     if lcl_includeTDTags then
        response.write "</td>" & vbcrlf
     end if
  end if

end sub

'------------------------------------------------------------------------------
function getTotalEmails(iMailList)

  dim lcl_return, sMailList, sdllist, iLength, sSQL, oGetTotalEmails

  lcl_return = 0
  sMailList  = ""

  if iMailList <> "" then
     if not containsApostrophe(iMailList) then
        sMailList = ucase(iMailList)
     end if

    'Get the distribution list(s)
    	sdllist = replace(sMailList,"X",",")  'Build comma separate list
    	iLength = len(sdllist)
    	sdllist = left(sdllist,(iLength - 1))  'Trim trailing comma
     sdllist = dbsafe(sdllist)

    	sSQL = "SELECT count(distinct(useremail)) as total_emails "
    	sSQL = sSQL & " FROM egov_dl_user_list "
    	sSQL = sSQL & " WHERE distributionlistid IN (" & sdllist & ") "
     sSQL = sSQL & " AND useremail is not null "
     sSQL = sSQL & " AND useremail <> '' "

     set oGetTotalEmails = Server.CreateObject("ADODB.Recordset")
     oGetTotalEmails.Open sSQL, Application("DSN"), 0, 1

     if not oGetTotalEmails.eof then
        lcl_return = oGetTotalEmails("total_emails")
     end if

     oGetTotalEmails.close
     set oGetTotalEmails = nothing

  end if

  getTotalEmails = lcl_return

end function
%>
