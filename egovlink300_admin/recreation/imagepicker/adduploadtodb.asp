<%@ Language=VBScript %>
<!-- #include file="../../includes/common.asp" //-->
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
		.Parameters.Append oCmd.CreateParameter("@RETURN_VALUE", 3, 4,4)
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
		.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
		.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, null)
		.Execute
	End With
	iNewID = oCmd.Parameters.Item("@RETURN_VALUE").Value
	Set oCmd = Nothing
	'response.write iNewID
	'response.end
'---END: Update DB fields----------------------------------

'Response.Redirect "addarticle.asp?strTitle=" & strTitle & "&task=ADD&method=upload&Message=" & strMessage 
%>

<html>
<head>
<script>
	function sendValuesBack() {
      		var objParent=window.opener;
	  	objParent.addItem.itemID.value='<%=iNewID%>';
	  	objParent.addItem.link.value='<%=strTitle%>';
	  	if(objParent.addItem.title.value=='')objParent.addItem.title.value='<%=strTitle%>';
		window.close();
	}
</script>
<meta name="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
</head>
<body onload="javascript:sendValuesBack();">



</body>
</html>




