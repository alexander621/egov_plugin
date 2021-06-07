<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: MultiClassEmail.asp
' AUTHOR: Steve Loar
' CREATED: 02/19/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page is for sending mass emais to several class/event rosters.
'				Taken from the single roster email send page
'
' MODIFICATION HISTORY
' 1.0	02/19/2013	Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
sLevel = "../"  'Override of value from common.asp

If Not userhaspermission(session("userid"), "registration") Then 
	response.redirect sLevel & "permissiondenied.asp"
End If 

Dim iSelectedClassIds, sFromEmail, sFromName, sOrgName, item, iClassCount, sClassEnd, sSubject, sMessageBody, iSendTo

'iClassId   = request("classid")
iSelectedClassIds = ""
iSentCount = request("sentcount")
sFromName  = ""
sFromEmail = ""
sSubject = ""
sMessageBody = ""
iClassCount = 0
sClassEnd = ""
iSendTo = clng(1)

If request("classid") <> "" Then 

	For Each item In request("classid")
		If iSelectedClassIds <> "" Then 
			iSelectedClassIds = iSelectedClassIds & ","
		End If 
		iSelectedClassIds = iSelectedClassIds & item 
		iClassCount = iClassCount + 1
		'Get default From Email and name
		sFromEmail = GetClassPOCEmail( item, sFromName )
	Next 

Else
	' call back from the sending page
	If request("selectedclassids") <> "" Then
		iSelectedClassIds = request("selectedclassids")
		iClassCount = request("classcount")
	End If 
End If 

If iClassCount > 1 Then 
	sClassEnd = "es"
End If 

If request("sendto") <> "" Then
	iSendTo = clng(request("sendto"))
End If 

If request("fromemail") <> "" Then
	sFromEmail = request("fromemail")
End If 

If request("fromname") <> "" Then
	sFromName = request("fromname")
End If 

If request("subject") <> "" Then
	sSubject = request("subject")
End If 

If request("messagebody") <> "" Then
	sMessageBody = request("messagebody")
End If 



'Check for a screen message
lcl_onload  = ""
lcl_success = request("success")

If lcl_success <> "" Then 
	lcl_msg = setupScreenMsg( lcl_success )

	If lcl_success = "SS" And iSentCount <> "" Then 
		lcl_msg = iSentCount & "&nbsp;" & lcl_msg
	End If 

	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
End If 

%>

<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="javascript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	<!--
		function Validate() 
		{
			var lcl_false_count = 0;

			if($("#messagebody").val() === "") 
			{
				lcl_focus = $("#messagebody");
				inlineMsg($("#messagebody").attr('id'),'<strong>Required Field Missing: </strong> Message Body',10,'messagebody');
				lcl_false_count = lcl_false_count + 1;
			}
			else
			{
				clearMsg("messagebody");
			}

			if($("#subject").val() === "") 
			{
				lcl_focus = $("#subject");
				inlineMsg($("#subject").attr('id'),'<strong>Required Field Missing: </strong> Subject',10,'subject');
				lcl_false_count = lcl_false_count + 1;
			}
			else
			{
				clearMsg("subject");
			}

			if($("#fromemail").val() === "") 
			{
				lcl_focus = $("#fromemail");
				inlineMsg(document.getElementById("fromemail").id,'<strong>Required Field Missing: </strong> From Email',10,'fromemail');
				lcl_false_count = lcl_false_count + 1;
			}
			else
			{
				clearMsg("fromemail");
			}

			if($("#fromname").val() === "") 
			{
				lcl_focus = $("#fromname");
				inlineMsg(document.getElementById("fromname").id,'<strong>Required Field Missing: </strong> From Name',10,'fromname');
				lcl_false_count = lcl_false_count + 1;
			}
			else
			{
				clearMsg("fromname");
			}

			if(lcl_false_count > 0) 
			{
				lcl_focus.focus();
				return false;
			}
			else
			{
				$("#MailForm").submit();
				return true;
			}
		}

		function displayScreenMsg( iMsg ) 
		{
			if( iMsg != "") 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
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

	response.write "<table id=""multititledisplay"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""600"">" & vbcrlf
	response.write "  <tr>"
	response.write "      <td><font size=""+1""><strong>Recreation: Send Email to Multiple Class Attendees</strong></font><br /><br /></td>"
	response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>"
	response.write "  </tr>"
	response.write "</table>" & vbcrlf

  	ShowEmailWarning

	response.write "<br /><input type=""button"" name=""returnButton"" id=""returnButton"" value=""<< Return to Class List"" class=""button"" onclick=""location.href='roster_list.asp';"" />" & vbcrlf
	

	response.write "<div id=""classcountdisplay"">" & iClassCount & " class" & sClassEnd & " selected.</div>" & vbcrlf
	

	response.write "<form name=""MailForm"" id=""MailForm"" method=""post"" action=""MultiClassEmailSend.asp"">" & vbcrlf
	response.write "<input type=""hidden"" name=""selectedclassids"" value=""" & iSelectedClassIds & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""classcount"" value=""" & iClassCount & """ />" & vbcrlf

	response.write "<table id=""multiemailinput"" border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

	'From Name
	response.write "  <tr>"
	response.write "      <td class=""multimaillabel""><strong>From Name:</strong></td>"
	response.write "      <td><input type=""text"" name=""fromname"" id=""fromname"" value=""" & sFromName & """ size=""50"" maxsize=""75"" onchange=""clearMsg('fromname')"" /></td>"
	response.write "  </tr>" & vbcrlf

	'From Email
	response.write "  <tr>"
	response.write "      <td class=""multimaillabel""><strong>From Email:</strong></td>" 
	response.write "      <td><input type=""text"" name=""fromemail"" id=""fromemail"" value=""" & sFromEmail & """ size=""50"" maxsize=""75"" onchange=""clearMsg('fromemail')"" /></td>"
	response.write "  </tr>" & vbcrlf

	'Send To
	response.write "  <tr>"
	response.write "      <td class=""multimaillabel""><strong>Send To:</strong></td>"
	response.write "      <td>"
	response.write "          <select name=""sendto"" id=""sendto"" onchange=""clearMsg('sendto')"">" & vbcrlf
	response.write "            <option value=""1"""
								If iSendTo = 1 Then
									response.write " selected=""selected"" "
								End If 
	response.write "            >Active Only</option>" & vbcrlf
	response.write "          		<option value=""2"""
								If iSendTo = 2 Then
									response.write " selected=""selected"" "
								End If 
	response.write "            >Waitlist Only</option>" & vbcrlf
	response.write "          		<option value=""3"""
								If iSendTo = 3 Then
									response.write " selected=""selected"" "
								End If 
	response.write "            >Active &amp; Waitlist</option>" & vbcrlf
	response.write "          </select>" & vbcrlf
	response.write "      </td>"
	response.write "  </tr>" & vbcrlf

	'Subject
	response.write "  <tr>"
	response.write "      <td class=""multimaillabel""><strong>Subject:</strong></td>"
	response.write "      <td><input type=""text"" name=""subject"" id=""subject"" value=""" & sSubject & """ size=""100"" maxsize=""100"" onchange=""clearMsg('subject')"" /></td>"
	response.write "  </tr>" & vbcrlf

	'Message Body
	response.write "  <tr><td colspan=""2"" class=""multimaillabel""><strong>Message Body:</strong></td></tr>" & vbcrlf
	response.write "  <tr>"
	response.write "      <td colspan=""2"" class=""multimaillabel"">"
	response.write "          <textarea name=""messagebody"" id=""messagebody"" onchange=""clearMsg('messagebody')"" />" & sMessageBody & "</textarea><br />"
	response.write "          * This message will be sent as HTML"
	response.write "      </td>"
	response.write "  </tr>" & vbcrlf
	
	response.write "</table>" & vbcrlf

	response.write "</form>" & vbcrlf

	response.write "<input type=""button"" class=""button"" name=""send"" value=""Send Email"" onclick=""Validate();"" /></p>" & vbcrlf

	response.write "  </div>" & vbcrlf
	response.write "</div>" & vbcrlf
	'END: Page Content -----------------------------------------------------------
%>

<!--#Include file="../admin_footer.asp"-->  


	</body>
</html>



