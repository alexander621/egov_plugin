
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../meetings/ShowNiceTime.asp" //-->
<!-- #include file="../meetings/ShowAgendas.asp" //-->
<%
Dim oCmd, oRstItems,intAgendaID, index, arrColors(2), imgURL, iType, sItems
Dim oRst, sMtgTopic, sMtgTime, sMtgPlace, sMtgReqBy, sMtgSummary, sMtgUrl
Dim smid, sAction, sScriptName
Dim sOutput, sAgendaSub, sAgendaDesc, sAgendaSort, iUserID

Call Main ()

Sub Main

	intAgendaID=Request.Form("AgendaID")
	if intAgendaID & "" = "" then intAgendaID=Request.QueryString("aid") End If

'Response.End
	
	smid = Request.QueryString("mid")
	If smid & "" = "" then smid = Request.Form("mid") End If
	
	sAction = Request.Form("action")
	sScriptName = Request.ServerVariables("Script_Name")	
	If sAction  = "update"  then 
		InitRecord
		UpdateAgenda
		Response.Redirect "../meetings/meeting_view.asp?mid=" & smid & "&aid=" & intAgendaID
	Else
		CreateItemList
		GetAgendaRecord
		'GetMeetingRecord
		'GetUserName
		ShowForm
	End if
End Sub

Sub MainItemRefs
		CreateItemList
		ShowForm
End sub

Sub CreateItemList

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
	.ActiveConnection = Application("DSN")
	.CommandText = "ListAgendaItems"
	.CommandType = adCmdStoredProc
	.Parameters.Append oCmd.CreateParameter("AgendaID", adInteger, adParamInput, 4, intAgendaID)
	End With

	Set oRstItems = Server.CreateObject("ADODB.Recordset")
	With oRstItems
	.CursorLocation = adUseClient
	.CursorType = adOpenStatic
	.LockType = adLockReadOnly
	.Open oCmd
	End With
	Set oCmd = Nothing

	arrColors(0)="ffffff"
	arrColors(1)="eeeeee"
	index=0

	If oRstItems.RecordCount = 0 then
		sItems = "<tr><td colspan=4>There are no references set for this Agenda</td></tr>"
	End if

	Do while not oRstItems.EOF
	iType=oRstItems("AgendaItemTypeID")
	 
	if iType = ITEM_TYPE_DOCUMENT then imgURL="../images/document_home.gif"
	if iType = ITEM_TYPE_VOTE then imgURL="../images/newpoll.gif"
	if iType = ITEM_TYPE_DISCUSSION then imgURL="../images/newdisc.gif"
	if iType = ITEM_TYPE_TEXT then imgURL="../images/document.gif"
	  
	sItems = sItems & "<tr bgcolor='" & arrColors(index) & "'>"
	If HasPermission("CanEditMeetings") Then
		sItems = sItems & "<td valign='top'><input type ='checkbox' class='listcheck' name='del_" & oRstItems("AgendaItemID") & "'></td>"
	Else
		sItems = sItems & "<td>&nbsp;</td>"
	End IF
	sItems = sItems & "<td style=""padding:0px;""><img src=""" & imgURL & """></td>"
	sItems = sItems & "<td><a href=""updateagendaitem.asp?id=" & oRstItems("AgendaItemID")
	sItems = sItems & "&mid=" & smid & """>" 
	sItems = sItems & oRstItems("AgendaItemTitle") & "</a></td>"
	sItems = sItems & "<td>"
	  
	if iType = ITEM_TYPE_DOCUMENT then sItems = sItems & "<a href=""../docs/viewdoc.asp?did=" & oRstItems("AgendaItemURL") & """> viewDoc.asp?did=" & oRstItems("AgendaItemURL") & "</a></td>"
	if iType = ITEM_TYPE_VOTE then sItems = sItems & "<a href=""../polls/takepoll.asp?id=" & oRstItems("AgendaItemURL") & """>Polls/takepoll.asp?id=" & oRstItems("AgendaItemURL") & "</a></td>"
	if iType = ITEM_TYPE_DISCUSSION then sItems = sItems & "<a href=""../discussions/thread.asp?mid=" & oRstItems("AgendaItemURL") & """>Discussions/thread.asp?mid=" & oRstItems("AgendaItemURL") & "</a></td>"
	if iType = ITEM_TYPE_TEXT then sItems = sItems & "&nbsp;</td>"
	oRstItems.MoveNext
	index=1-index
	Loop
'
End Sub

Sub InitRecord
	sAgendaSub		= Request.Form("Subject")
	sAgendaDesc		= Request.Form("Description")
'	sAgendaSort		= clng(Request.Form("Sort"))
	sAgendaSort		=	1
End Sub

Sub UpdateAgenda

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
	.ActiveConnection = Application("DSN")
	.CommandText = "UpdateAgenda"
	.CommandType = adCmdStoredProc
	.Parameters.Append oCmd.CreateParameter("AgendaID", adInteger, adParamInput, 4, intAgendaID)
	.Parameters.Append oCmd.CreateParameter("AgendaSubject", adVarChar, adParamInput, 50, sAgendaSub)	
	.Parameters.Append oCmd.CreateParameter("AgendaDescription", adVarChar, adParamInput, 250, sAgendaDesc)	
	.Parameters.Append oCmd.CreateParameter("AgendaSortNumber", adInteger, adParamInput, 4, sAgendaSort)
	.Execute
	End With
'
End Sub

Sub GetAgendaRecord  
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetAgenda"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("AgendaID", adInteger, adParamInput, 4, intAgendaID)
		.Execute
	End With
'
	Set oRst = Server.CreateObject("ADODB.Recordset")
	With oRst
		.CursorLocation = adUseClient
		.CursorType = adOpenStatic
		.LockType = adLockReadOnly
		.Open oCmd
	End With
	Set oCmd = Nothing
'
	sAgendaSub		= oRst("AgendaSubject")
	sAgendaDesc		= oRst("AgendaDescription")
'	sAgendaSort		= oRst("AgendaSortNumber")
	SAgendaSort		=	1	
	Set oRst = Nothing
End Sub

Sub ShowUpdateAgenda()

    sOutput = "<table border=0 cellpadding=5 cellspacing=0>"
    sOutput = sOutput & "<tr>"
    sOutput = sOutput & "<td style=""font-weight:bold; color:#336699;"">" & langAgendaSub & ":</td>"
    sOutput = sOutput & "<td><input name=subject type=text value=""" & sAgendaSub & """ size=50 maxlength=100></td>"
    sOutput = sOutput & "</tr>"
    sOutput = sOutput & "<tr>"
    sOutput = sOutput & "<td valign=""top"" style=""font-weight:bold; color:#336699;"">" & langAgendaDesc & ":</td>"
'	<textarea name="Sum" rows=5 cols=50
    sOutput = sOutput & "<td><textarea rows=5 cols=50 name=Description maxlength=250>" & sAgendaDesc & "</textarea></td>"
    sOutput = sOutput & "</tr>"
' Agenda Sort item functionality not decided yet! .. Hard coded in the add_agenda
'    sOutput = sOutput & "<tr>"
'	sOutput = sOutput & "<td style=""font-weight:bold; color:#336699;"">" & langAgendaSort & "</td>"
'   sOutput = sOutput & "<td><Input name=sort type=text value=""" & sAgendaSort & """ size =4 maxlength=4></td>"
'    sOutput = sOutput & "</tr>"
	'
	ShowUpdateForm
End Sub
%>

<% Sub ShowUpdateForm%>
        
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="messagehead" ID="Table1">
        <Form name="UpdateAgenda" method=post action="<%=sScriptName%>" ID="Form1">
		<input type=hidden name="action" value="update" ID="Hidden1">
		<input type=hidden name="mid" value=<%=smid%> ID="Hidden2">
		<input type=hidden name="AgendaID" value=<%=intAgendaID%> ID="Hidden3">	
		 <tr>
			<th width="100%" align=left>&nbsp;&nbsp;<%=langUpdateAgenda%>&nbsp;</th>
         </tr>
         <tr>
			<td colspan="2">
              <table border="0" cellpadding="0" cellspacing="0" ID="Table2">
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
  <script language="Javascript">
<!-- #include file="../scripts/modules.js" //-->
  </script>
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
    <%DrawTabs tabMeetings,1%>
  <table border="0" cellpadding="10" cellspacing="0" width="100%" ID="Table3">
    <tr>
      <td width="151" align="center"><img src="../images/icon_meeting.jpg"></td>
      <td><font size="+1"><b><%=langUpdateAgenda%></b>
      </font><br><img src="../images/spacer.gif"  height=16 width=16 align="absmiddle">&nbsp;
<!--
      <img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;
      <a href="../meetings"><%=langBack2MeetingsList%></a> -->
		</td>  
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
      <!-- #include file="quicklinks.asp" //-->
		<% Call DrawQuicklinks("",1) %> 
      </td>
      <td colspan="2" valign="top">
		    <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="meeting_view.asp?mid=<%=smid%>"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateAgenda.submit();"><%=langUpdate%></a></div>

         <%'Response.Write " Here is the show agendas .. smid = " & smid %>
		   <%'ShowAgendas%>
		   <%ShowUpdateAgenda%><%'GeneralInfo%>		
<!--  <div style="font-size:10px; padding-top:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="meeting_view.asp?mid=<%=smid%>"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateAgenda.submit();"><%=langUpdate%></a></div> -->
		<br><br>	
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		  <tr>	      
		   <%ShowItemsForm%>
		  </tr>
		</table>		

</body>
</html>
<%End Sub%>


<%Sub GeneralInfo%>

        <table width="100%" cellpadding="5" cellspacing="0" border="0" class="messagehead" ID="Table4">
				  <tr>
						<th align="left"><%=langGeneralInfo%></th>
				  </tr>	
				  <tr>
            <td>
              <table border="0" cellpadding="5" cellspacing="0" ID="Table5">
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langTopic%></td>
                  <td><%= sMtgTopic %> </td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langWhen%></td>
                  <td><%= sMtgTime %></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langWhere%></td>
                  <td><%= sMtgPlace %></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langReqBy%></td>
                  <td><%=sMtgReqBy%></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;" nowrap valign="top"><%=langSummary%></td>
                  <td><%= sMtgSummary %></td>
                </tr>
			  </table>
            </td>
          </tr>
		</table>
<%End Sub%>

<%Sub ShowItemsForm%>
      <td colspan="2" valign="top">
        <form name="DelAgendaItems" method=post action="deleteagendaitems.asp" method="post">
        <input type=hidden name="AgendaID" value=<%=intAgendaID%>>
        <input type=hidden name="mid" value=<%=smid%> ID="Hidden4">
		<%'If HasPermission("CanEditMeetings") Then %>
          <div style="font-size:10px; padding-bottom:5px;">
			  <img src="../images/newagendaitem.gif" align="absmiddle">&nbsp;
			  <a href="newagendaitem.asp?aid=<%=intAgendaID%>&mid=<%=smid%>">New Agenda Item</a>
			  &nbsp;&nbsp;&nbsp;&nbsp;
			  <img src="../images/small_delete.gif" align="absmiddle">&nbsp;
			  <a href="javascript:document.DelAgendaItems.submit();" ><%=langDelete%></a>
		  </div>
		<%'End If %>
          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th align=left>
              <% If HasPermission("CanEditMeetings") Then %>
              <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelAgendaItems', this.checked)"></th>
              <%Else%>
              &nbsp;
              <%End If%>
              </th>
              <th>&nbsp;</th>
              <th align="left" width="25%"><%=langDescription%></th>
              <th align="left" width="70%"><%=langURL%></th>
            </tr>
            <tr><%=sItems%></tr>
          </table>
        </form>
      </td>
<% End Sub%>
