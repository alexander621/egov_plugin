<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: subscribe_new.asp
' AUTHOR: Steve Loar
' CREATED: 09/06/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  New Subscription landing page from subscribe.asp.
'
' MODIFICATION HISTORY
' 1.0   09/06/06   Steve Loar - Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->

<html>
<head>
<title>E-Gov Services <%=sOrgName%> - Subscription Registration</title>
<link rel="stylesheet" href="../css/styles.css" type="text/css">
<link rel="stylesheet" href="../global.css" type="text/css">
<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">


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
			<p>You will receive a verification e-mail shortly that contains a link which you will need to 
			visit in order to confirm your subscription.</p>  
		</div> <br />  <br />  <br />
	</div>
</div>

	<P>&nbsp;</p>
   
<!--#Include file="../include_bottom.asp"-->    
<!--#Include file="../includes/inc_dbfunction.asp"-->    


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>