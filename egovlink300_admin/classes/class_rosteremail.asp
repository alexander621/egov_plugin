<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_rosteremail.asp
' AUTHOR: Steve Loar
' CREATED: 05/08/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page is for sending mass emais to class/event rosters.
'
' MODIFICATION HISTORY
' 1.0 04/26/06	Steve Loar - INITIAL VERSION
' 1.1	10/17/06	Steve Loar - Security, Header and nav changed
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "registration") then
    response.redirect sLevel & "permissiondenied.asp"
 end if

 Dim iClassId, iTimeId, sFromEmail, sFromName, sOrgName

 iClassId   = request("classid")
 iTimeId    = request("timeid")
 iSentCount = request("sentcount")
 sFromName  = ""

'Get default From Email and name
 sFromEmail = GetClassPOCEmail( iClassId, sFromName )

'Set the default From Name
 'sOrgName = GetOrgName( Session("orgid") )
 'sFromName = sOrgName & " E-GOV WEBSITE"

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
%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

 <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--

	function Validate() 
	{
		// Check that a from name was entered
//		if (document.MailForm.fromname.value == "")
//		{
//			alert('Please provide a From Name.');
//			document.MailForm.fromname.focus();
//			return;
//		}
		// Check that a from email was entered
//		if (document.MailForm.fromemail.value == "")
//		{
//			alert('Please provide a From Email Address.');
//			document.MailForm.fromemail.focus();
//			return;
//		}
		// Check that a subject was entered
//		if (document.MailForm.subject.value == "")
//		{
//			alert('Please provide a Subject.');
//			document.MailForm.subject.focus();
//			return;
//		}
		// Check that a message body was entered
//		if (document.MailForm.messagebody.value == "")
//		{
//			alert('Please provide a Message Body.');
//			document.MailForm.messagebody.focus();
//			return;
//		}

//		document.MailForm.submit();
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
  response.write "      <td><font size=""+1""><strong>Recreation: Class Roster - Send Email to Attendees</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to Roster"" class=""button"" onclick=""location.href='view_roster.asp?classid=" & iClassId & "&timeid=" & iTimeId & "';"" />" & vbcrlf
  response.write "</p>" & vbcrlf

  ShowEmailWarning

 'Send Email To Attendees
  response.write "<p><strong>Send Email to Attendees in:</strong><span style=""color:#800000"">&nbsp;" & GetClassName( iClassId ) & " &nbsp; ( " & GetActivityNo( iTimeId ) & " )</span></p>" & vbcrlf


  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "  <form name=""MailForm"" id=""MailForm"" method=""post"" action=""class_rosteremailsend.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""classid"" value=""" & iClassId & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />" & vbcrlf

 'From Name
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>From Name:</strong></td>" & vbcrlf
  response.write "      <td><input type=""text"" name=""fromname"" id=""fromname"" value=""" & sFromName & """ size=""50"" maxsize=""75"" onchange=""clearMsg('fromname')"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'From Email
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>From Email:</strong></td>" & vbcrlf
  response.write "      <td><input type=""text"" name=""fromemail"" id=""fromemail"" value=""" & sFromEmail & """ size=""50"" maxsize=""75"" onchange=""clearMsg('fromemail')"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Send To
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>Send To:</strong></td>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <select name=""sendto"" id=""sendto"" onchange=""clearMsg('sendto')"">" & vbcrlf
  response.write "            <option value=""1"">Active Only</option>" & vbcrlf
  response.write "          		<option value=""2"">Waitlist Only</option>" & vbcrlf
  response.write "          		<option value=""3"">Active &amp; Waitlist</option>" & vbcrlf
  response.write "          </select>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Subject
  response.write "  <tr>" & vbcrlf
  response.write "      <td><strong>Subject:</strong></td>" & vbcrlf
  response.write "      <td><input type=""text"" name=""subject"" id=""subject"" value="""" size=""100"" maxsize=""100"" onchange=""clearMsg('subject')"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Message Body
  response.write "  <tr><td colspan=""2""><strong>Message Body:</strong></td></tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <textarea name=""messagebody"" id=""messagebody"" value="""" onchange=""clearMsg('messagebody')"" /></textarea><br />" & vbcrlf
  response.write "          * This message will be sent as HTML" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<p><input type=""button"" class=""button"" name=""send"" value=""Send Email"" onclick=""Validate();"" /></p>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: Page Content -----------------------------------------------------------
%>

<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
