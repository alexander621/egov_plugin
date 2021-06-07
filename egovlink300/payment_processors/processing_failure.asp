<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaltemplate.asp
' AUTHOR: Steve Loar
' CREATED: 01/13/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Template for rentals public pages
'
' MODIFICATION HISTORY
' 1.0   01/13/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, sErrMsg, iGatewayErrorId

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

'If request("ge") <> "" Then 
	iGatewayErrorId = CLng(request("ge"))
	sErrMsg = GetGatewayErrorMsg( iGatewayErrorId, iOrgId )		' in ../includes/common.asp
'Else
'	response.redirect "../"
'End If 


%>

<html>
<head>

	<title><%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<p>
	<font class="pagetitle">Payment Processing Failure</font>
	<br />
</p><br /><br />
<p>
	<strong>
	We are sorry but we cannot process your payment at this time due to the following error. Please try again later.
	</strong>
</p>
<p>
	<strong>Error:</strong> <%=sErrMsg%>
</p><br /><br />


<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

