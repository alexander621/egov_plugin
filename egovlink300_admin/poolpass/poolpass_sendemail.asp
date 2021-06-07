<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_sendemail.asp
' AUTHOR: Steve Loar
' CREATED: 06/23/2010
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page is for sending mass emails to membership (head of households).
'
' MODIFICATION HISTORY
' 1.0 06/23/2010	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iMembershipId, oMembership

 sLevel = "../"  'Override of value from common.asp

 'if not userhaspermission(session("userid"), "send_emails_to_members") then
 if not orghasfeature("send_emails_to_members") then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine the membership type
 lcl_membership_type = "pool"

 set oMembership = New classMembership
 'set the membershipid to the one for pools
 oMembership.SetMembershipId(lcl_membership_type)

'Get default From Email and name
'---------------------------------------------------------
' sFromEmail = GetClassPOCEmail( iClassId, sFromName )
'---------------------------------------------------------

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg = setupScreenMsg(lcl_success)

    if lcl_success = "SS" AND iSentCount <> "" then
       lcl_msg = iSentCount & "&nbsp;" & lcl_msg
    end if

    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

 sSQL                = session("sendEmailsToMembers_query")
 lcl_userid          = ""
 iTotalDistinctCount = 0
 lcl_scripts         = ""

 if sSQL = "" then
    lcl_onload = lcl_onload & "setupSendButton('DISABLED');"
 end if
%>
<html>
<head>
	<title>E-Gov Administration Console {Send Email to Members}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

 <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--

function Validate() {
  var lcl_false_count = 0;

  if(document.getElementById("messagebody").value=="") {
     lcl_focus = document.getElementById("messagebody");
     inlineMsg(document.getElementById("messagebody").id,'<strong>Required Field Missing: </strong> Message Body',10,'messagebody');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("messagebody");
  }

  if(document.getElementById("subject").value=="") {
     lcl_focus = document.getElementById("subject");
     inlineMsg(document.getElementById("subject").id,'<strong>Required Field Missing: </strong> Subject',10,'subject');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("subject");
  }

  if(document.getElementById("fromemail").value=="") {
     lcl_focus = document.getElementById("fromemail");
     inlineMsg(document.getElementById("fromemail").id,'<strong>Required Field Missing: </strong> From Email',10,'fromemail');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("fromemail");
  }

  if(document.getElementById("fromname").value=="") {
     lcl_focus = document.getElementById("fromname");
     inlineMsg(document.getElementById("fromname").id,'<strong>Required Field Missing: </strong> From Name',10,'fromname');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("fromname");
  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("MailForm").submit();
     return true;
  }
	}

function setupSendButton(p_mode) {
  if(p_mode == "" || p_mode == "undefined") {
     lcl_mode = "ENABLED";
  }else{
     lcl_mode = p_mode;
  }

  if(lcl_mode == "DISABLED") {
     document.getElementById("sendButton").disabled = true;
  }else{
     document.getElementById("sendButton").disabled = false;
  }
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
 'BEGIN: Page Content ----------------------------------------------------------
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""600"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Send Email to Members:</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  	ShowEmailWarning

  response.write "<input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to Member List"" class=""button"" onclick=""location.href='member_list.asp';"" />" & vbcrlf
  response.write "</p>" & vbcrlf

 'Send Email
  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "  <form name=""MailForm"" id=""MailForm"" method=""post"" action=""poolpass_sendemail_action.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""classid"" value=""" & iClassId & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />" & vbcrlf

 'From Name
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>From Name:</strong></td>" & vbcrlf
  response.write "      <td><input type=""text"" name=""fromname"" id=""fromname"" value=""" & sFromName & """ size=""60"" maxsize=""75"" onchange=""clearMsg('fromname')"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'From Email
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>From Email:</strong></td>" & vbcrlf
  response.write "      <td><input type=""text"" name=""fromemail"" id=""fromemail"" value=""" & sFromEmail & """ size=""60"" maxsize=""75"" onchange=""clearMsg('fromemail')"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Send To
  'response.write "  <tr>" & vbcrlf
  'response.write "      <td><strong>Send To:</strong></td>" & vbcrlf
  'response.write "      <td><input type=""button"" name=""viewMembers"" id=""viewMembers"" value=""View Members to be Emailed"" class=""button"" onclick=""alert('view members');"" /></td>" & vbcrlf
  'response.write "  </tr>" & vbcrlf

 'Subject
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>Subject:</strong></td>" & vbcrlf
  response.write "      <td><input type=""text"" name=""subject"" id=""subject"" value="""" size=""60"" maxsize=""100"" onchange=""clearMsg('subject')"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Message Body
  response.write "  <tr><td colspan=""2""><strong>Message Body:</strong></td></tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <textarea name=""messagebody"" id=""messagebody"" rows=""20"" cols=""80"" onchange=""clearMsg('messagebody')"" /></textarea><br />" & vbcrlf
  response.write "          * This message will be sent as HTML" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<p><input type=""button"" class=""button"" name=""send"" id=""sendButton"" value=""Send Email"" onclick=""Validate();"" /></p>" & vbcrlf

 'Check for members to be emailed
  if sSQL <> "" then
     set oEmailMembers = Server.CreateObject("ADODB.Recordset")
    	oEmailMembers.Open sSQL, Application("DSN"), 0, 1

     if not oEmailMembers.eof then
        do while not oEmailMembers.eof
           if lcl_userid = "" then
              lcl_userid = oEmailMembers("userid")
           else
              lcl_userid = lcl_userid & ", " & oEmailMembers("userid")
           end if

           oEmailMembers.movenext
        loop
     end if

     oEmailMembers.close
     set oEmailMembers = nothing

     if lcl_userid <> "" then
        sSQL = "SELECT DISTINCT useremail "
        sSQL = sSQL & " FROM egov_users "
        sSQL = sSQL & " WHERE userid IN (" & lcl_userid & ") "
        sSQL = sSQL & " AND useremail <> '' "
        sSQL = sSQL & " AND useremail is not null "
        sSQL = sSQL & " ORDER BY useremail "

        set oEmailMembersDistinct = Server.CreateObject("ADODB.Recordset")
       	oEmailMembersDistinct.Open sSQL, Application("DSN"), 0, 1

        if not oEmailMembersDistinct.eof then
           response.write "<p>" & vbcrlf
           response.write "<fieldset>" & vbcrlf
           response.write "  <legend>Members to be emailed:&nbsp;</legend>" & vbcrlf
           response.write "<br />" & vbcrlf

           do while not oEmailMembersDistinct.eof

              iTotalDistinctCount = iTotalDistinctCount + 1

              response.write oEmailMembersDistinct("useremail") & "<br />" & vbcrlf

              oEmailMembersDistinct.movenext
           loop

           response.write "  <p><strong>Total: </strong>[" & iTotalDistinctCount & "]</p>" & vbcrlf
           response.write "</fieldset>" & vbcrlf
           response.write "</p>" & vbcrlf
        end if

        oEmailMembersDistinct.close
        set oEmailMembersDistinct = nothing

     end if
  end if

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: Page Content -----------------------------------------------------------
%>

<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "SS" then
        lcl_return = "Message(s) Successfully Sent..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>
