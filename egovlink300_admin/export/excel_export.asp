<%
' SET UP PAGE OPTIONS
sDate = Month(Date()) & Day(Date()) & Year(Date())
server.scripttimeout = 4800
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=EGOV_EXPORT_" & sDate & ".xls"


' CREATE FILE FOR DOWNLOAD
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
	
		response.write "<html><body><table border=""1"">"
		response.write "</tr>"
		' WRITE COLUMN HEADINGS
		For Each fldLoop in oSchema.Fields
			response.write  "<th>" & fldLoop.Name & "</th>"
		Next
		response.write "</tr>"
		response.flush

		' WRITE DATA
		Do While NOT oSchema.EOF
			response.write "<tr>"
			For Each fldLoop in oSchema.Fields
				sFieldValue = trim(fldLoop.Value)
				
				' REMOVE LINE BREAKS
				If NOT ISNULL(sFieldValue) Then
					sFieldValue = replace(sFieldValue,chr(10),"")
					sFieldValue = replace(sFieldValue,chr(13),"")
				End If

				response.write "<td>" & sFieldValue & "</td>"
			Next
			response.write "</tr>"
			response.flush

			oSchema.MoveNext
		Loop

		response.write "</table></body></html>"
		response.flush
	Else

		' NO DATA

	End If

	oSchema.Close
	Set oSchema = Nothing

End Sub
%>
