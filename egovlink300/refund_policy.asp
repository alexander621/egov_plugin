<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: refund_policy.asp
' AUTHOR: Steve Loar
' CREATED: 07/06/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   07/06/2007	Steve Loar	-  Initial Version
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>

<html>
<head>

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

</head>

<!--#Include file="include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "" ) %>

<div id="content">
	<div id="centercontent">

		<h5><%=sOrgName%> Refund Policy</h5><br /><br />
		<div id="refundpolicy">
		<%=GetOrgDisplay( iOrgId, "refund policy" ) %>
		</div>
		
	</div>
</div>

<!--END: PAGE CONTENT-->


<!--#Include file="include_bottom.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>

