<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../meetings/ShowNiceTime.asp" //-->
<!-- #include file="../meetings/ShowAgendas.asp" //-->
<%
Dim oCmd, oRst, sMtgTopic, sMtgTime, sMtgPlace, sMtgReqBy, sMtgSummary, sMtgUrl
Dim smid, sAction, sScriptName
Dim sOutput, sAgendaSub, sAgendaDesc, sAgendaSort

Call Main()

Sub Main
	smid = Request.QueryString("mid")

	If smid & "" = "" then smid = Request.Form("mid")

	sAction = Request.Form("action")
	sScriptName = Request.ServerVariables("Script_Name")
'	
	If sAction  = "add"  then 
		CleanRecords
		NewAgenda
		Response.Redirect "../meetings/meeting_view.asp?mid=" & smid
	Else
		'GetMeetingRecord
		ShowForm
	End if
'		
End Sub
  
Sub CleanRecords
	sAgendaSub		= Request.Form("Subject")
	sAgendaDesc		= Request.Form("Description")
'	sAgendaSort		= clng(Request.Form("Sort"))
	sAgendaSort		= 1

End Sub

Sub UpdateAgenda


End Sub

Sub NewAgenda
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
	.ActiveConnection = Application("DSN")
	.CommandText = "NewAgenda"
	.CommandType = adCmdStoredProc
	.Parameters.Append oCmd.CreateParameter("NewAgendaID", adInteger, adParamReturnValue,4)
	.Parameters.Append oCmd.CreateParameter("MeetingID", adInteger, adParamInput, 4, smid)
	.Parameters.Append oCmd.CreateParameter("AgendaSubject", adVarChar, adParamInput, 50, sAgendaSub)	
	.Parameters.Append oCmd.CreateParameter("AgendaDescription", adVarChar, adParamInput, 250, sAgendaDesc)	
	.Parameters.Append oCmd.CreateParameter("AgendaSortNumber", adInteger, adParamInput, 4, sAgendaSort)
	.Execute
	End With
'
End Sub
%>

<%
Sub ShowAddAgenda()

    sOutput = "<table border=0 cellpadding=5 cellspacing=0>"
    sOutput = sOutput & "<tr>"
    sOutput = sOutput & "<td style=""font-weight:bold; color:#336699;"">" & langAgendaSub & ":</td>"
    sOutput = sOutput & "<td><Input name=subject type=text value=""" & sAgendaSub & """ size=50 maxlength=100></td>"
    sOutput = sOutput & "</tr>"
    sOutput = sOutput & "<tr>"
    sOutput = sOutput & "<td style=""font-weight:bold; color:#336699;"">" & langAgendaDesc & ":</td>"
'	<textarea name="Sum" rows=5 cols=50
    sOutput = sOutput & "<td><textarea rows=5 cols=50 name=description type=text value=""" & sAgendaDesc & """ maxlength=250></textarea></td>"
    sOutput = sOutput & "</tr>"
'    sOutput = sOutput & "<tr>"
'    sOutput = sOutput & "<td style=""font-weight:bold; color:#336699;"">" & langAgendaSort & "</td>"
'    sOutput = sOutput & "<td><Input name=sort type=text value=""" & sAgendaSort & """ size =4 maxlength=4></td>"
'    sOutput = sOutput & "</tr>"
	'
	ShowAddForm
End Sub
%>

<% Sub ShowAddForm%>
        
		<table width="100%" cellpadding="5" border="0" cellspacing="0"  class="messagehead">
		<Form name="AddAgenda" method=post action="<%=sScriptName%>">
		<input type=hidden name="action" value="add">
		<input type=hidden name="mid" value=<%=smid%>>
          <tr style="height:22px;">
            <td width="100%" bgcolor="#93bee1" class="section_hdr" style="border-bottom:1px solid #336699;">&nbsp;&nbsp;<%=langAddAgenda%>&nbsp;</td>
          </tr>
          <tr>
			<td colspan="2">
              <table border="0" cellpadding="1" cellspacing="0">
                <%= sOutput %>
              </table>
            </td>
			
          </tr>
	   </Form>
	   </table>

<%End Sub%>

<% Sub ShowForm%>
<html>
<head>
  <title><%=langBSMeetings%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript" src="../scripts/modules.js"></script>
  <Script language="JavaScript">
	<!--
		function ltrim (s) { return s.replace(/^\s*/,"")}
		function rtrim (s) { return s.replace(/\s*$/,"")}
		function trim (s) { return rtrim(ltrim(s))}
		
		var errorwindow;
		var startHTML = "<HTML><HEAD><TITLE>Meeting: Add Agenda : Errors</TITLE>";
		startHTML += "<B>These are the errors that need to be corrected</B></HEAD><BODY><BR><BR>";
		var errors = startHTML
		function OpenErrorWindow () {
			if (errorwindow != null ) { errorwindow.close; }
			errorwindow = window.open("", "Error", "HEIGHT=300, WIDTH=500, alwaysRaised=true");
		}
		
		function WriteError() {
			errorwindow.document.write(errors);
			errorwindow.document.close();
			errorwindow.focus;
		}
			
		function Validate()	{
			if ( Valid() == true ) {
					document.all.AddAgenda.submit();				
			};
		}

		function Valid () {

			var subject		= trim(document.forms.AddAgenda.subject.value);
			document.forms.AddAgenda.subject.value = subject;

			var description	= trim(document.forms.AddAgenda.description.value);
			document.forms.AddAgenda.description.value = description;

			var valid = true;
			if ( description == "")  {
				errors += "<li>Description can not be spaces</li>";
				valid = false;
			}
			if ( subject == "")  {
				errors += "<li>Subject can not be spaces</li>";
				valid = false;
			}
	
			
			if (valid) { return true }
			else { OpenErrorWindow();
					errors += "</BODY></HTML>";
					WriteError();
					errors = startHTML
			}
			
		}
		
	//-->

	</script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
    <%DrawTabs tabMeetings,1%>
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_meeting.jpg"></td>
      <td><font size="+1"><b><%=langNewAgenda%></b></font><br>
      
      <img src="../images/spacer.gif"  height=16 width=16 align="absmiddle">&nbsp;</td>
<!--    
      <img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../meetings"><%=langBack2MeetingsList%></a></td>
      <td width="200">&nbsp;</td>
-->
    </tr>
    <tr>
      <td valign="top">
      <!-- #include file="quicklinks.asp" //-->
		<% Call DrawQuicklinks("",1) %> 
      </td>
      <td colspan="2" valign="top">
		<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:Validate();"><%=langCreate%></a></div>
		<%'GeneralInfoTable%>
<%'Response.Write " Here is the show agendas add_agenda .. smid = " & smid %>
		<%'ShowAgendas%>
		<%ShowAddAgenda%>
		<div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:Validate();"><%=langCreate%></a></div>
</body>
</html>
<%End Sub%>
