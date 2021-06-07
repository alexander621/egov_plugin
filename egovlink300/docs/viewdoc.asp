<% response.end %>

<!-- #include file="../includes/common.asp" //-->
<%



Dim sSql, oRst, sPath, sLink, iNewWindow, docid

docid = Request.QueryString("did")

If docid & "" = "" Then
  Response.Write "Invalid request format."
  Response.End
End If


sSql = "EXEC GetDocumentPathByID 5,176," & docid

response.write sSQL
response.end


Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .ActiveConnection = Application("DSN")
  .CursorLocation = adUseClient
	.CursorType = adOpenStatic
	.LockType = adLockReadOnly
	.Open sSql
  .ActiveConnection = Nothing
End With

If Not oRst.EOF Then
  sPath = oRst("DocumentURL") & ""
  sLink = oRst("LinkURL") & ""
  iNewWindow = oRst("LinkTargetsNew")
  
	'--------------------------------------------------------------------------------------------------
	' BEGIN: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
		iSectionID = 44
		If sLink <> "" Then
			sDocumentTitle = RIGHT(sPath,Len(sPath)- InstrRev(sPath,"/"))
			If sDocumentTitle = "" Then
				sDocumentTitle = "UNKNOWN DOCUMENT"
			End If
		Else
			sDocumentTitle = RIGHT(sLink,Len(sLink)- InstrRev(sLink,"/"))
			If sDocumentTitle = "" Then
				sDocumentTitle = "UNKNOWN LINK"
			End If
		End If
		
		sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
		datDate = Date()	
		datDateTime = Now()
		sVisitorIP = request.servervariables("REMOTE_ADDR")
	'	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
	'--------------------------------------------------------------------------------------------------
	' END: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------

  
  oRst.Close


End If
Set oRst = Nothing

If sLink <> "" Then
  Response.Redirect sLink
Else
  Response.Redirect sPath
End If
%>