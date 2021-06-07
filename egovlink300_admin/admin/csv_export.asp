<%
' SET UP PAGE OPTIONS
sDate = Month(Date()) & Day(Date()) & Year(Date())
server.scripttimeout = 4800
Response.ContentType = "application/msexcel"
Response.AddHeader "Content-Disposition", "attachment;filename=EGOV_EXPORT_" & sDate & ".CSV"


' CREATE CSV FILE FOR DOWNLOAD
Call CreateDownload()
%>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' FUNCTION CREATEDOWNLOAD()
'------------------------------------------------------------------------------------------------------------
Sub CreateDownload()
	
	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = session("DISPLAYQUERY")
	oSchema.Open sSQL, Application("DSN"), 3, 1

	If NOT oSchema.EOF Then
	
		' WRITE COLUMN HEADINGS
		For Each fldLoop in oSchema.Fields
			response.write  chr(34) & fldLoop.Name & chr(34) & ","
		Next
		response.write  chr(34) & "Document Size (Bytes)" & chr(34)
		response.write vbcrlf
		response.flush

		' WRITE DATA
		Do While NOT oSchema.EOF
			For Each fldLoop in oSchema.Fields
				response.write chr(34) & trim(fldLoop.Value) & chr(34) & ","
			Next
			response.write  chr(34) & GetTotalDocumentsSize(oSchema("Org ID") ) & chr(34)
			response.write vbcrlf
			response.flush

			oSchema.MoveNext
		Loop
	
	Else

		' NO DATA

	End If

	Set oSchema = Nothing

End Sub


'-------------------------------------------------------------------------------------------------
' Function GetTotalDocumentsSize( sOrgId )
'-------------------------------------------------------------------------------------------------
Function GetTotalDocumentsSize( ByVal sOrgId )
	Dim sSql, oRs

	sSql = "SELECT SUM(documentsize) AS documentsize "
	sSql = sSql & " FROM documents WHERE orgid = " & sOrgId
	sSql = sSql & " GROUP BY orgid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
'		If CDbl(oRs("documentsize")) > CDbl(1024) Then 
'			GetTotalDocumentsSize = FormatNumber((CDbl(oRs("documentsize")) / CDbl(1024)),0)  & " KB"
'		Else
			GetTotalDocumentsSize = FormatNumber(oRs("documentsize"),0,,,0) '& " Bytes"
'		End If 
	Else
		GetTotalDocumentsSize = "0" 'Bytes"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
