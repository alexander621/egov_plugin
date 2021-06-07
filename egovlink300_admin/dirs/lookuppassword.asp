<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
'response.end
	PageIsRequiredByLogin = True 
%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp"-->

<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: lookuppassword.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  page where admin user can lookup their password.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	???? - INITIAL VERSION
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

sLevel = "../" ' Override of value from common.asp

'Dim iorgid,iPaymentGatewayID,blnOrgRegistration,blnQuerytool,blnFaq
SetOrganizationParameters()

%>

<html>
<head>
	<title><%=langBSHome%></title>

	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="JavaScript">
	<!--
		function CheckEmail()
		{
			if (document.byemail.email.value == "")
			{
				alert("Please provide your Email");
				document.byemail.email.focus();
				return false;				
			}					
			return true;
		}

		function CheckUsername()
		{
			if (document.byusername.username.value == "")
			{
				alert("Please provide your Username");
				document.byusername.username.focus();
				return false;				
			}					
			return true;
		}
	//-->
	</script>


</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs 0,0%>
  <% ShowHeader sLevel %>

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" width="100%"  class="start" >
    <tr>
      <td valign="top" width='151'> &nbsp;
		 <!--<center> <img src='../images/icon_directory.jpg'></center>-->
	 <br>
	      
      </td>
      <td colspan="2" valign="top">

<font size="+1"><b>Reset Your Password</b></font>
	  <br><img src='../images/arrow_back.gif' align='absmiddle'> <a href='../login.asp'><%=langGoBack%></a>


<table cellpadding="5" cellspacing="0" width="200" border="0">
  <form method="post" name="byusername" Action="passwordreset.asp">
	<tr><td width="100%">Enter your username or email address</td></tr>
	<tr>
		<td width="100%">
			<input size="30" name="username">
			&nbsp; <a href="javascript:document.byusername.submit();" onclick="return CheckUsername();">
			<img src='../images/go.gif' border="0"><%=langGo%></a>
		</td>
	</tr>
</table>
<input type="hidden" name="emailstart" value="Y" />

</form>

 </td>
  <td width="200">&nbsp;</td>
    </tr>
 </table>

 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  
  

</body>
</html>
 



<%
'------------------------------------------------------------------------------------------------------------
' FUNCTION SETORGANIZATIONPARAMETERS()
'------------------------------------------------------------------------------------------------------------
Function SetOrganizationParameters()
	' SET DEFAULT RETURN VALUE
	iReturnValue = 1

	' BUILD CURRENT URL
	If request.servervariables("HTTPS") = "on" Then
		sProtocol = "https://"
	Else
		sProtocol = "http://"
	End If
	sSERVER = request.servervariables("SERVER_NAME")
	sCurrent = sProtocol & sSERVER & "/" & GetVirtualDirectyName()


	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT * FROM Organizations WHERE OrgEgovWebsiteURL='" & sCurrent & "'"

	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oOrgInfo.EOF Then
		iOrgID = oOrgInfo("OrgID")
		sOrgName = oOrgInfo("OrgName")
		sHomeWebsiteURL = oOrgInfo("OrgPublicWebsiteURL")
		sEgovWebsiteURL = oOrgInfo("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oOrgInfo("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oOrgInfo("OrgTopGraphicRightURL")
		sWelcomeMessage = oOrgInfo("OrgWelcomeMessage")
		sActionDescription = oOrgInfo("OrgActionLineDescription")
		sPaymentDescription = oOrgInfo("OrgPaymentDescription")
		iHeaderSize = oOrgInfo("OrgHeaderSize")
		sTagline = oOrgInfo("OrgTagline")
		iPaymentGatewayID = oOrgInfo("OrgPaymentGateway")
		blnOrgAction = oOrgInfo("OrgActionOn")
		blnOrgPayment = oOrgInfo("OrgPaymentOn")
		blnOrgDocument = oOrgInfo("OrgDocumentOn")
		blnOrgCalendar = oOrgInfo("OrgCalendarOn")
		sorgVirtualSiteName = oOrgInfo("orgVirtualSiteName")
		blnOrgRegistration = oOrgInfo("orgRegistration")
		blnQuerytool = oOrgInfo("orgQueryTool")
		blnFaq = oOrgInfo("orgFaqOn")
	End If
	Set oOrgInfo = Nothing 

	If NOT ISNULL(iOrgID) Then 
		iReturnValue = iOrgID
	End If

	' RETURN VALUE
	SetOrganizationParameters = iReturnValue
	
End Function


'------------------------------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()

	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 0) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function

%>
