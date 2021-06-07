<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>


<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->


<html>
<head>


<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>


<link rel="stylesheet" href="../css/styles.css" type="text/css">
<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">
<link href="stylesheet" rel="../global.css" type="text/css">
<script language="Javascript" src="../scripts/modules.js"></script>
<script language="Javascript" src="../scripts/easyform.js"></script>
</head>


<!--#Include file="../include_top.asp"-->



<!--BEGIN PAGE CONTENT-->
PLACE CONTENT HERE
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------
%>