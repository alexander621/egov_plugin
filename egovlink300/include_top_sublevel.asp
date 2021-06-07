<%
' GET INFORMATION FOR CURRENT INSTITUTION
Dim sTopGraphicLeftURL,sHomeWebsiteURL,sTopGraphicRighURL,sEgovWebsiteURL 
Dim sWelcomeMessage,sTagline
iOrgID = SetOrganizationParameters()

If bCustomButtonsOn = 1 Then
	sImgDir = "custom/images/" & sorgVirtualSiteName & "/"
Else
  sImgDir = "/img/"
End If
%>

<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 topmargin=0>


<table cellspacing=0 cellpadding=0 border=0 bordercolor=green width=100% height=100%>


<!--BEGIN: TOP GRAPHIC-->
  <tr>
    <td width=100% align=left valign=top background="<%=sTopGraphicRighURL%>"><a href="default.asp"><img name="City Logo" src="<%=sTopGraphicLeftURL%>" border="0" alt="Click here to return to the E-Government Services start page"></a></td>
    <td width=1 height=57 background="<%=sTopGraphicRighURL%>"><img src="<%=sEgovWebsiteURL%>/img/clearshim.gif" border="0" width=1 height=57></td>
  </tr>
<!--END: TOP GRAPHIC-->

<!--BEGIN: BUTTON ROW-->
  <tr>
    <td background="<%=sEgovWebsiteURL%>/img/button_finish.gif">
	<table cellspacing=0 cellpadding=0 height=24 border=0 bordercolor=red>
		<tr>
			<td><a href="<%=sHomeWebsiteURL%>"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_city.gif" border="0"></a></td>
	    	<td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td>
	    	<td><a href="<%=sEgovWebsiteURL%>/"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_egov.gif" border="0"></a></td>
	    	<td><img src="<%=sEgovWebsiteURL%>/img/button_line.gif" border="0"></td>
	    	<td><a href="<%=sEgovWebsiteURL%>/action.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_action.gif" border="0"></a></td>
	    	<td><img src="<%=sEgovWebsiteURL%>/img/button_line.gif" border="0"></td>
	    	<td><a href="<%=sEgovWebsiteURL%>/events/calendar.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_calendar.gif" border="0"></a></td>
	    	<td><img src="<%=sEgovWebsiteURL%>/img/button_line.gif" border="0"></td>
	    	<td><a href="<%=sEgovWebsiteURL%>/docs/menu/home.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_docs.gif" border="0"></a></td>
	    	<td><img src="<%=sEgovWebsiteURL%>/img/button_line.gif" border="0"></td>
	    	<td><a href="<%=sEgovWebsiteURL%>/payment.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_permits.gif" border="0"></a></td>
  	</table>
	</td>
    <td width=1 height=24 background="<%=sEgovWebsiteURL%>/img/button_finish.gif"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>clearshim.gif" border="0" width=1 height=24></td>
  </tr>
  <!--END: BUTTON ROW-->

  <!--BEGIN: BUTTON ROW SHADOW-->
  <tr>
    <td background="<%=sEgovWebsiteURL%>/img/horiz_shadow_14px.gif"><img src="<%=sEgovWebsiteURL%>/img/horiz_shadow_14px.gif" border="0" height=14></td>
    <td width=1 height=14 background="<%=sEgovWebsiteURL%>/img/horiz_shadow_14px.gif"><img src="<%=sEgovWebsiteURL%>/img/clearshim.gif" border="0" width=1 height=14></td>
  </tr>
  <!--END: BUTTON ROW SHADOW-->
  


  <!--BEGIN: MAIN BODY CONTENT-->
  <tr>
    <td valign=top class=indent20><p class=verdana>

  


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTION AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


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
	sURL = REPLACE(lcase(request.servervariables("URL")),lcase("/" & GetPageName()),"")
	sCurrent = sProtocol & sSERVER & sURL


	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT * FROM Organizations WHERE OrgEgovWebsiteURL='" & sCurrent & "'"
	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oOrgInfo.EOF Then
		iOrgID = oOrgInfo("OrgID")
		sHomeWebsiteURL = oOrgInfo("OrgPublicWebsiteURL")
		sEgovWebsiteURL = oOrgInfo("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oOrgInfo("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oOrgInfo("OrgTopGraphicRightURL")
		sWelcomeMessage = oOrgInfo("OrgWelcomeMessage")
		sTagline = oOrgInfo("OrgTagline")
	End If
	Set oOrgInfo = Nothing 

	If NOT ISNULL(iOrgID) Then 
		iReturnValue = iOrgID
	End If

	' RETURN VALUE
	SetOrganizationParameters = iReturnValue
	
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETPAGENAME()
'------------------------------------------------------------------------------------------------------------
Function GetPageName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	For Each arr in strURL 
		sReturnValue = arr 
	Next 
	
	GetPageName = sReturnValue

End Function



%>