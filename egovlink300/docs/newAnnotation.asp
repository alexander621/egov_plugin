<!-- #include file="../includes/common.asp" //-->

<%
' Check to see if we need to update this window
If request.querystring("task")="add" Then
	'Add annotation to the database
	
   	'---BEGIN: Update DB fields--------------------------------
      Set oCmd = Server.CreateObject("ADODB.Command")
      With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "NewAnnotation"
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("User", adInteger, adParamInput, 4, Session("UserID"))
	  .Parameters.Append oCmd.CreateParameter("Description", adVarChar, adParamInput, 5000, request("txtAnnotation"))
	  .Parameters.Append oCmd.CreateParameter("DocID", adInteger, adParamInput, 4, request("ID"))
	  .Execute
      End With
      Set oCmd = Nothing
      '---END: Update DB fields----------------------------------
	
	'Refresh back page with new task
    response.write "<script language=""Javascript"">window.opener.document.location.reload(true);</script>"
	'Close Window after adding task
	response.write "<script language=""Javascript"">window.close();</script>"
End If

 


%>

<html>
  <head>
    <title>
      New Document Annotation
    </title>
	  <link href="../global.css" rel="stylesheet" type="text/css">
  </head>
 
 <body>
   
  
  
  <table valign=top class=annotation>
  
	<tr>
	<td>
    <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="#" onClick="window.close();">Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.frmNewAnnotation.submit();">Create</a></div>
    </td>
	</tr>
	<tr>
    <td>
	 <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
              <tr>
                <th align="left" colspan="2"> <img class="menuimage" src="menu/images/annotate.gif" width="18" height="18" align="absmiddle">New Annotation</th>
              </tr>
              <tr><td>
	<form name="frmNewAnnotation" action="newAnnotation.asp?task=add" method="post">
    <textarea name="txtAnnotation" rows=15 cols=60></textarea>
	<input name="ID" type=hidden value="<%=request("ID")%>">
    </form>
	</td>
	</tr>
	</table>
	</td>
	</tr>
	<tr>
	<td>
	<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="#" onClick="window.close();" >Cancel</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.frmNewAnnotation.submit();">Create</a></div>
    </td>
	</tr>
    <tr><td align=center valign=top><hr class=annotation size="2" color="#000000" width="95%"><b><%=Now()%></b></td></tr>
	</table>
 </body>

 </html>

