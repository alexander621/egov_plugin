<%@ Language=VBScript %>
<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer = True

' SET VARIABLES
strDir = session("MyDoc")
' It is necessary to pass these because the upload form submits to 
' file multipart/form-data file, so it is necessary to carry these
' values thru the querystring

strTitle=request.QueryString("strTitle") 
strMessage=request.QueryString("Message") 


'---BEGIN: Update DB fields for Document(DIRECT CONTENT) --------------------------------
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "NewDocument"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
		.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
		.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, null)
		.Execute
	End With
	Set oCmd = Nothing
'---END: Update DB fields----------------------------------

Response.Redirect "addarticle.asp?strTitle=" & strTitle & "&task=ADD&method=upload&Message=" & strMessage 
%>

<html>
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
</head>
<body>



</body>
</html>




