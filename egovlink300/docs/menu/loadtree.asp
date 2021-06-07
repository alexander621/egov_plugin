<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="URLDecode.asp" //-->
<!-- #include file="loadfolder_db.inc" //-->


<%
Dim strPath, strList
strPath = URLDecode( Request("path") )


' GET CITY DOCUMENT LOCATION
sLocationName = "pikeville"
strList = LoadFolder("/public_documents/" & sLocationName, strPath  )	



' CHECK TO SEE IF THE PATH IS A VALID FILE
'If instr(1,strPath,".") < 1 Then
	'Response.write "Not Valid."
	'Response.end
'End If
%>


<html>
<head>
  <script>
    function window_load() {
      parent.loadFrame();
    }
    window.onload = window_load;
  </script>
</head>

<body>
<div id=chunk>
  <%
  If strList = "" Then
    Response.Write "&nbsp;<font color=""#003366""><i>Permission Denied</i></font>"
  Else
    Response.Write strList
  End If

 ' response.write "strPath="&strPath
  %>
</div>
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
	sCurrent = sProtocol & sSERVER & GetVirtualDirectyName()


	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT * FROM Organizations WHERE OrgEgovWebsiteURL='" & sCurrent & "'"
	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSQL, Application("DSN"), 3, 1
	
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


'------------------------------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = sReturnValue

End Function

%>

