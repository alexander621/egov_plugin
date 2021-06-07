<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: subscribe_confirm.asp
' AUTHOR: Steve Loar
' CREATED: 09/06/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Subscription confirmation.  Called from link in email sent to subscriber.
'
' MODIFICATION HISTORY
' 1.0   09/06/06   Steve Loar - Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

	Dim iUserid

	If request("u") <> "" Then 
		iUserid = CLng(request("u"))

		' Delete from to_user
		ClearMailList iUserid, "to_user" 

		' Copy from the temp list to prod
		CopyMailList iUserid 
		
		' Clear the temp list
		ClearMailList iUserid, "temp" 
	Else
		response.redirect "subscribe.asp"
	End If 

%>
<html>
<head>

	<title>E-Gov Services <%=sOrgName%> - Subscription Registration</title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />


</head>

<!--#Include file="../include_top.asp"-->

<!--BODY CONTENT-->

<tr>
	<td valign="top">
		<%  RegisteredUserDisplay( "../" ) %>	

<div id="content">
	<div id="centercontent">
		
		<div class="box_header4"><%=sOrgName%> Subscriptions</div>
		<div class="groupsmall2">
			<p>Thank you.  Your subscription choices have been 
			<% If request("s") <> "" Then%>
				saved. 
			<% Else %>
				confirmed.
			<% End If %>
			</p> 
		</div> <br />  <br />  <br />
	</div>
</div>

	<p>&nbsp;</p>
   
<!--#Include file="../include_bottom.asp"-->    
<!--#Include file="../includes/inc_dbfunction.asp"-->    


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub CopyMailList( iUserid )
'--------------------------------------------------------------------------------------------------
Sub CopyMailList( iUserid )
	Dim sSQL, oList

	sSQL = "Select distributionlistid FROM egov_class_distributionlist_temp WHERE userid = " & iUserid 

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	Do While Not oList.EOF 
		InsertMailList iUserid, oList("distributionlistid") 
		oList.movenext
	Loop

	oList.close
	Set oList = Nothing 

End Sub 
	

'--------------------------------------------------------------------------------------------------
' SUB InsertMailList( iUserid, iListid )
'--------------------------------------------------------------------------------------------------
Sub InsertMailList( iUserid, iListid )
	Dim sSql, oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")

	' Insert subscription
	oCmd.CommandText = "INSERT INTO egov_class_distributionlist_to_user ( userid, distributionlistid ) VALUES ( '" & iUserid & "', '" & iListid & "' )"
	oCmd.Execute
	Set oCmd = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB ClearMailList( iUserid, sList )
'--------------------------------------------------------------------------------------------------
Sub ClearMailList( iUserid, sList )
	Dim sSql, oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")

	If LCase(sList) = "temp" Then 
		sList = "egov_class_distributionlist_temp"
	Else
		sList = "egov_class_distributionlist_to_user"
	End If 

	' Insert subscription
	oCmd.CommandText = "Delete From " & sList & " where userid = " & iUserid 
	oCmd.Execute
	Set oCmd = Nothing

End Sub


%>