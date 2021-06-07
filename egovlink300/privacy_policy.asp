<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
	Dim sError 

	If OrgHasFeature( iOrgId, "no privacy policy" ) Then	
		response.redirect "./"
	End If 

%>

<html>
<head>
<title>E-Gov Services <%=sOrgName%></title>

<link rel="stylesheet" href="css/styles.css" type="text/css">
<link rel="stylesheet" href="global.css" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">

<script language="Javascript" src="scripts/modules.js"></script>
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
</script>
</head>


<!--#Include file="include_top.asp"-->


<TR><TD VALIGN=TOP>

<% if iOrgId <> 153 then %>
<!-- start privacy policy code -->
<p><h3>Privacy Policy</h3></p>
<p><b>The <%=sOrgName%><% if iOrgId = 187 then response.write " and ""Make it Moon""" %> respects your privacy.</b></p>
<p>Any personal information you provide to us including and similar to your name, address, telephone number and e-mail address will not be released, sold, or rented to any entities or individuals outside of the <%=sOrgName%> except as required by law.</p>

<p><b>External Sites.</b></p>
<p>The <%=sOrgName%> is not responsible for the content of external internet sites. You are advised to read the privacy policy of any external sites before disclosing any personal information.</p>

<p><b>Cookies </b></p>
<p>A "cookie" is a small data text file that is placed in your browser and allows electronic commerce link to recognize you each time you visit this site(customisation etc). Cookies themselves do not contain any personal information, and the <%=sOrgName%>  does not use cookies to collect personal information. Cookies may also be used by 3rd party content providers such as newsfeeds.</p>

<p><b>Remember The Risks Whenever You Use The Internet </b></p>
<p>While we do our best to protect your personal information, we cannot guarantee the security of any information that you transmit to the <%=sOrgName%> and you are solely responsible for maintaining the secrecy of any passwords or other account information. In addition other Internet sites or services that may be accessible through the <%=sOrgName%> E-government website have separate data and privacy practices independent of us, and therefore we disclaim any responsibility or liability for their policies or actions. </p>
<p>Please contact those vendors and others directly if you have any questions about their privacy policies.</p>
<!-- end code -->
<% end if %>


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="include_bottom.asp"-->  
