<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: subscription_none.asp
' AUTHOR: Steve Loar
' CREATED: 11/27/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This catches spam attack attempts.
'
' MODIFICATION HISTORY
' 1.0   11/27/2007	Steve Loar	-  Initial Version
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'Set up title
 if iorgid = 7 then
    lcl_title = sOrgName
 else
    lcl_title = "E-Gov Services " & sOrgName
 end if
%>
<html>
<head>
		<title><%=lcl_title%></title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

</head>

<!--#Include file="../include_top.asp"-->

<%
 'BEGIN: Page Content ---------------------------------------------------------
  RegisteredUserDisplay( "../" )

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "    <div id=""refundpolicy"">" & vbcrlf
  response.write "      <p>We are evaluating some new security features. If you see this message, your subscription request will not be processed.</p>" & vbcrlf
  response.write "      <p>If you are using a pop-up blocker such as Google Toolbar, please press your back button, enable Pop-ups, and re-submit your request.</p>" & vbcrlf
  response.write "      <p>Sorry for the inconvenience.</p>" & vbcrlf
  response.write "    </div>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: Page Content -----------------------------------------------------------
%>

<!--#Include file="../include_bottom.asp"-->  
