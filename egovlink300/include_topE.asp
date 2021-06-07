<%
If bCustomButtonsOn Then
	sImgDir = "/custom/images/" & sorgVirtualSiteName & "/"
Else
  sImgDir = "/img/"
End If
%>

<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 topmargin=0>


<table cellspacing=0 cellpadding=0 border=1 bordercolor=red >


<!--BEGIN: TOP GRAPHIC-->
  <tr>
    <td COLSPAN=2 height=<%=iHeaderSize%> width=100% valign=top background="<%=sTopGraphicRighURL%>">
		<a href="<%=sHomeWebsiteURL%>"><img name="City Logo" src="<%=sTopGraphicLeftURL%>" border="0" alt="Click here to return to the E-Government Services start page"></a>
	</td>
    <td width=1  height=<%=iHeaderSize%> background="<%=sTopGraphicRighURL%>">
		<img src="<%=sEgovWebsiteURL%>/img/clearshim.gif" border="0" width=1 height=<%=iHeaderSize%> >
	</td>
  </tr>
<!--END: TOP GRAPHIC-->
  <tr>



<%If blnMenuOn Then %>
    <td background="<%=sEgovWebsiteURL%><%=sImgDir%>button_finish.gif" VALIGN=TOP>
				<table cellspacing=0 cellpadding=0 height=24 border=1 bordercolor=gold>
				
					<!--MAIN CITY WEBSITE -->
					<tr><td><a href="<%=sHomeWebsiteURL%>"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_city.gif" border="0"></a></td></tr>
			    <tr><td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td></tr>
		
					<!--EGOVLINK CITY HOME -->
			    	<tr><td><a href="<%=sEgovWebsiteURL%>/"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_egov.gif" border="0"></a></td></tr>
			    	<tr><td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td></tr>
		
					<!--ACTION LINE TAB -->
					<%If blnOrgAction Then%>
			    	<tr><td><a href="<%=sEgovWebsiteURL%>/action.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_action.gif" border="0"></a></td></tr>
			    	<tr><td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td></tr>
					<%End If%>
		
					<!--CALENDAR TAB -->
					<%If blnOrgCalendar Then%>
			    	<tr><td><a href="<%=sEgovWebsiteURL%>/events/calendar.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_calendar.gif" border="0"></a></td></tr>
			    	<tr><td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td></tr>
					<%End If%>
		
					<!--DOCUMENT TAB -->
					<%If blnOrgDocument Then%>
			    <tr><td><a href="<%=sEgovWebsiteURL%>/docs/menu/home.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_docs.gif" border="0"></a></td></tr>
					<tr><td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td></tr>
					<%End IF%>
		
					<!--PAYMENT TAB -->
					<%If blnOrgPayment Then%>
			    	<tr><td><a href="<%=sEgovWebsiteURL%>/payment.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_permits.gif" border="0"></a></td></tr>
			    	<tr><td><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_line.gif" border="0"></td></tr>
					<%End If%>
					
					<!--FAQ TAB -->
					<%If blnOrgFaq Then%>
			    	<tr><td><a href="<%=sEgovWebsiteURL%>/faq.asp"><img src="<%=sEgovWebsiteURL%><%=sImgDir%>button_faq.gif" border="0"></a></td></tr>
					<%End If%>
				
		  	</table>
	</td>

<%Else%>
  <td background="<%=sEgovWebsiteURL%><%=sImgDir%>button_finish.gif">
			&nbsp;
	</td>
<%End If%>
  
  
  

  <!--BEGIN: MAIN BODY CONTENT-->
    <td valign=top><TABLE BORDER=2 BORDERCOLOR=PINK>
 


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
	sCurrent = sProtocol & sSERVER & "/" & GetVirtualDirectyName()


	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT * FROM Organizations INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID WHERE OrgEgovWebsiteURL='" & sCurrent & "'"

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
		blnOrgFaq = oOrgInfo("OrgFaqOn")
		sorgVirtualSiteName = oOrgInfo("orgVirtualSiteName")
		sOrgActionName =  oOrgInfo("OrgActionName")
		sOrgPaymentName =  oOrgInfo("OrgPaymentName")
		sOrgCalendarName =  oOrgInfo("OrgCalendarName")
		sOrgDocumentName =   oOrgInfo("OrgDocumentName")
		'sOrgFaqName =  oOrgInfo("OrgFaqName")
		sOrgRegistration = oOrgInfo("OrgRegistration")
		blnCalRequest = oOrgInfo("OrgRequestCalOn")
		iCalForm =  oOrgInfo("OrgRequestCalForm")
		sHomeWebsiteTag = oOrgInfo("OrgPublicWebsiteTag")
		sEgovWebsiteTag = oOrgInfo("OrgEgovWebsiteTag")
		bCustomButtonsOn = oOrgInfo("OrgCustomButtonsOn")
		iTimeOffset = oOrgInfo("gmtoffset")
		blnMenuOn = oOrgInfo("orgDisplayMenu")
		blnFooterOn =	oOrgInfo("orgDisplayFooter")
	End If

	oOrgInfo.Close
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


'------------------------------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function




%>