<!DOCTYPE HTML>
<%
Response.AddHeader "Access-Control-Allow-Origin", "*"
showGraph = true
Server.ScriptTimeout = 3600

blnInvalid = false



if request("wpsend") = "wpmethod" then 
	'Log the user in and redirect them to subscriptions
	sSQL = "SELECT u.UserID,u.orgid,o.orgname,o.OrgEgovWebsiteURL  " _
		& " FROM Users u  " _
		& " INNER JOIN organizations o ON o.orgid = u.orgid " _
		& " WHERE  u.IsRootAdmin = 1 and u.Username = 'eclink' and o.OrgVirtualSiteName = '" & GetVirtualDirectyName() & "'"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then
		response.cookies("User")("UserID") = oRs("UserID")
		response.cookies("User")("OrgID") = oRs("orgid")
		Session("orgid") = oRs("orgid")
		Session("userid") = oRs("userid")

		intOrgID = oRs("orgid")
		intUserID = oRs("userid")
		strOrgName = oRs("OrgName")
		strEGovURL = oRs("OrgEgovWebsiteURL")
		
	end if
end if
%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: DL_SENDMAIL.ASP
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

if request("wpsend") <> "wpmethod" then 
 intOrgID = session("orgid")
 intUserID = session("userid")
 strOrgName = session("sOrgName")
 strEGovURL = session("egovclientwebsiteurl")
 end if

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
    if isValidEmail( LCase(request("sFromEmail")) ) then

        'Don't do this if coming from scheduled post
        if request("logid") = "" then
    	    saveSubscriptionInfo lcl_list_type, lcl_dl_logid
        else
	    lcl_dl_logid = request("logid")
        end if
    
        'Only Process these if sending is "NOW"
        if request("scheduled") <> "LATER" then
    	    subProcessEmail lcl_dl_logid
    	    if request.form("sendlist") <> "" then SendPushNotification
    	    updateSubscriptionInfoStatus lcl_dl_logid, "COMPLETED"
        else
    	    updateJustSubscriptionInfoStatus lcl_dl_logid, "SCHEDULED"
        end if
    
        'if coming from scheduled POST, then we need to clear the schedule so it doesn't run again
        if request("logid") <> "" then
    	        clearSubscriptionInfoSchedule lcl_dl_logid
	        response.end
        end if
    
        if request("wpsend") = "wpmethod" then 
	        response.write "Send"
	        response.end
        end if
    

        if request("screen_mode") = "AUTOSEND" then
    
	    'Show the right message
    	    if request("scheduled") <> "LATER" then
       		    lcl_onload = lcl_onload & "displayScreenMsg('Email Successfully Sent');"
	    else
       		    lcl_onload = lcl_onload & "displayScreenMsg('Email Successfully Scheduled');"
	    end if
           lcl_show_return_postings_link = "Y"
        end if
    
        bSent = True
    else
	strMsg = "MESSAGE NOT SENT. ***<br /> ***From Email Address is INVALID"
	if not isNotSuppressed(request("sFromEmail")) then
		strMsg = strMsg & " because it is suppressed.***<br />***Please contact support."
	end if

        lcl_onload = lcl_onload & "displayScreenMsg('" & strMsg & "');"
	blnInvalid = true
    end if
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

<!--  <meta http-equiv="X-UA-Compatible" content="IE=9"> -->

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
 	<link rel="stylesheet" type="text/css" href="classes.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

 	<script type="text/javascript" src="selectbox_script.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<!-- 	<script type="text/javascript" src="../ckeditor/ckeditor.js"></script> -->
 	<script type="text/javascript" src="../ckeditor/ckeditor.js"></script>
  <script type="text/javascript" src="../ckeditor/adapters/jquery.js"></script>


 	<script type="text/javascript">
<!--
//function init() {
  //$('#sHTMLBodyEditor').ckeditor();
//}

$(document).ready(function() {
   $('#sHTMLBodyEditor').ckeditor();

   //$('#sHTMLBodyEditor').change(function() {

   CKEDITOR.instances.sHTMLBodyEditor.on('blur', function(e) {
      checkMsgLength();
   });
   //}

//updateElement
   //});

});
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

  updateEditor_or_Field('FIELD','');

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

function updateEditor_or_Field(iUpdateMode, iLink) {
 	// Get the editor instance that we want to interact with.
 	var oEditor          = CKEDITOR.instances.sHTMLBodyEditor;
  var lcl_editor_value = $('#sHTMLBodyEditor').val();
  var lcl_field_value  = $('#sHTMLBodyEditor').val();
  var lcl_update_mode  = 'FIELD';  //FIELD: The EDITOR has been updated and we need to update the FIELD
                                   //EDITOR: The FIELD has been updated and we need to update the EDITOR

 	// Check the active editing mode.
	 if(oEditor.mode == 'wysiwyg' ) {
     // Insert HTML code.
		   // http://docs.cksource.com/ckeditor_api/symbols/CKEDITOR.editor.html#insertHtml
     if(iLink != '') {
        lcl_field_value  = lcl_field_value + '&nbsp;' + iLink
        lcl_editor_value = '&nbsp;' + iLink;
     }

     if(iUpdateMode != '') {
        lcl_update_mode = iUpdateMode;
     }

     if(lcl_update_mode == 'EDITOR') {
    		  oEditor.insertHtml(lcl_editor_value);
     }

     $('#sHTMLBody').val(lcl_field_value);
 	} else {
    	alert( 'You must be in WYSIWYG mode!' );
  }

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
   if($('#sHTMLBodyEditor').val() == '') {
      $('#sHTMLBody').val('');
      $('#sHTMLBodyEditor').focus();
      inlineMsg(document.getElementById("sHTMLBodyDIV").id,'<strong>Required Field Missing: </strong> Email Body',10,'sHTMLBodyDIV');
      lcl_return_false = lcl_return_false + 1;
   }else{
      updateEditor_or_Field('FIELD','');
      //$('#sHTMLBody').val($('#sHTMLBodyEditor').val());
      clearMsg("sHTMLBodyDIV");
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
   if(document.frmlocation.scheduled.value == "LATER")
   {
	   var curDate = new Date();
	   var selDate = new Date(document.frmlocation.fromDate.value + " " + document.frmlocation.time.value);
	   if (curDate >= selDate)
	   {
		if (!confirm("You've selected a date/time that is in the past.  Are you sure you want to send this message now?"))
		{
   			lcl_return_false = lcl_return_false + 1;
		}
	   }

   }


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
  $('#sHTMLBodyEditor').val('');
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
      sSQL = sSQL & " WHERE distributionlistid IN (" & lcl_dl_listids & ") and NOT distributionlistid IN (1953,1972)  "
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

function checkMsgLength()
{
  lcl_length_subject = document.getElementById("sSubjectLine").value.length;
  //lcl_length_body    = document.getElementById("sHTMLBody").value.length;

  lcl_emailbody = $('#sHTMLBodyEditor').val();
  $('#sHTMLBody').val(lcl_emailbody);

  lcl_length_body = lcl_emailbody.trim().length;

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

  updateEditor_or_Field('FIELD','');

  if($('#sHTMLBodyEditor').val() != '') {
     lcl_emailbody = $('#sHTMLBodyEditor').val();
  }

  //if(document.getElementById("containsHTML").checked) {
     lcl_containsHTML = 'Y';
  //}

  lcl_previewEmail_URL  = 'preview_email.asp';
  lcl_previewEmail_URL += '?listtype=<%=lcl_list_type%>';
  lcl_previewEmail_URL += '&emailbody='    + encodeURIComponent(lcl_emailbody);
  lcl_previewEmail_URL += '&containsHTML=' + encodeURIComponent(lcl_containsHTML);

  //  window.open(lcl_previewEmail_URL,"_blank");
  //eval('window.open("' + lcl_previewEmail_URL + '", "_preview", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
  eval('window.open("about:blank", "email_preview", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
  document.email_preview.listtype.value = '<%=lcl_list_type%>';
  document.email_preview.emailbody.value = lcl_emailbody;
  document.email_preview.containsHTML.value = lcl_containsHTML;
  document.email_preview.submit();
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     <% if not blnInvalid then %>
     window.setTimeout("clearScreenMsg()", (10 * 1000));
     <% end if %>
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
	if request("scheduled") <> "LATER" then
     		lcl_onload = lcl_onload & "displayScreenMsg('Message successfully sent to the selected subscribers.');"
	else
     		lcl_onload = lcl_onload & "displayScreenMsg('Message successfully scheduled.');"
	end if
  elseif blnInvalid then
  else
     if lcl_screen_mode = "AUTOSEND" OR lcl_dl_listids <> "" then
        lcl_onload = lcl_onload & "setupAutoSendLists();"
     else
        lcl_onload = ""
     end if
  end if

  lcl_onload = lcl_onload & "checkMsgLength();"

  'response.write "<body onload=""init();" & lcl_onload & """>" & vbcrlf
  response.write "<body onload=""" & lcl_onload & """>" & vbcrlf
  response.write "<form name=""email_preview"" action=""preview_email.asp"" method=""POST"" target=""email_preview""><input type=""hidden"" name=""listtype"" value="""" /><input type=""hidden"" name=""emailbody"" value="""" /><input type=""hidden"" name=""containsHTML"" value="""" /></form>"
  response.write "<form name=""frmlocation"" id=""frmlocation"" action=""dl_sendmail.asp"" method=""post"">" & vbcrlf
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

  ShowEmailWarning

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
  'response.write "                              <input type=""checkbox"" name=""containsHTML"" id=""containsHTML"" value=""Y""" & lcl_checked_containsHTML & " />&nbsp;Customized HTML Email Body (contains HTML tags: HTML, HEAD, and/or BODY)&nbsp;" & vbcrlf
  response.write "                              <input type=""hidden"" name=""containsHTML"" id=""containsHTML"" value=""Y"" size=""1"" maxlength=""3"" />" & vbcrlf
  response.write "                              <input type=""button"" value=""Add a Link"" class=""button"" onClick=""doPicker('frmlocation.sHTMLBody','Y','Y','Y','Y');"" />" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td colspan=""2"">" & vbcrlf
  response.write "                              <div id=""sHTMLBodyDIV"">" & vbcrlf  'This DIV is ONLY for the error message to find with the editor!
  'response.write "                                <textarea name=""sHTMLBodyEditor"" id=""sHTMLBodyEditor"" style=""width:800px; height:200px;"" onchange=""alert('1');clearMsg('sHTMLBodyDIV');"" onkeyup=""checkMsgLength();"">" & lcl_html_body & "</textarea>" & vbcrlf
  response.write "                                <textarea name=""sHTMLBodyEditor"" id=""sHTMLBodyEditor"" style=""width:800px; height:200px;"">" & lcl_html_body & "</textarea>" & vbcrlf
  response.write "                                <textarea name=""sHTMLBody"" id=""sHTMLBody"" style=""display:none"">" & lcl_html_body & "</textarea>" & vbcrlf
  response.write "                              </div>" & vbcrlf
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
  response.write "						<input type=""radio"" name=""scheduled"" value=""NOW"" checked onClick=""disableSchedule()"" /> Send Now<br />"
  response.write "						<input type=""radio"" name=""scheduled"" value=""LATER"" onClick=""enableSchedule()"" /> Send Later<br /><div style=""display:none"" id=""schedulediv"">"
  response.write "						<input type=""text"" name=""fromDate"" id=""fromDate"" value=""" & date() & """ size=""10"" style=""margin-left:25px;"" maxlength=""10"" onchange=""clearMsg('fromDateCalPop');"" />"
  response.write "      			                <a href=""javascript:void doCalendar('From');""><img src=""../images/calendar.gif"" id=""fromDateCalPop"" border=""0"" onclick=""clearMsg('fromDateCalPop');"" /></a>"
  response.write "						<select name=""time"">"
  								stTime = date() & " 12:00 am"
								selected = false
								for x = 0 to 47
									selText = ""
									if selected = false and dateadd("n",(x-1) * 30,stTime) > now() then 
										selText = " selected"
										selected = true
									end if
									looptime = dateadd("n",x * 30,stTime)

									response.write "<option" & selText & ">" & formatdatetime(looptime,3) & "</option>" & vbcrlf
								next
  response.write "						</select></div>"
  response.write "<script>"
  response.write "function enableSchedule() {"
  response.write "	document.getElementById('schedulediv').style.display='';"
  response.write "}"
  response.write "function disableSchedule() {"
  response.write "	document.getElementById('schedulediv').style.display='none';"
  response.write "}"
  response.write "function doCalendar(ToFrom) {"
    response.write "w = (screen.width - 350)/2;"
    response.write "h = (screen.height - 350)/2;"
    response.write "eval('window.open(""../action_line/calendarpicker.asp?p=1&ToFrom=' + ToFrom + '"", ""_calendar"", ""width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '"")');"
  response.write "}"
  response.write "</script>"

  response.write "<br />"
  response.write "<br />"
  response.write "           					</td>" & vbcrlf
  response.write "        				</tr>" & vbcrlf

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
	sSql = sSql & " AND (parentid = '' OR parentid IS NULL) and NOT distributionlistid IN (1953,1972) "
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
sub saveSubscriptionInfo(ByVal iListType, ByRef lcl_dl_logid)

 'Validate the fields
  sListType           = "NULL"
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

  lcl_scheduledDateTime = "NULL"
  if request("scheduled") = "LATER" then
     lcl_scheduledDateTime = "'" & request("fromDate") & " " & request("time") & "'"
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
  sSQL = sSQL & " scheduledDateTime "
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
  sSQL = sSQL & sdllist		    & ", "
  sSQL = sSQL & lcl_scheduledDateTime
  sSQL = sSQL & ")"

  lcl_dl_logid = RunInsertStatement(sSQL)


end sub

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
%>
<!--#Include file="inc_sendmail.asp"-->  
