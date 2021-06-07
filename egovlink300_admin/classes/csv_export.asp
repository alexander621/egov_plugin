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
		response.write vbcrlf
		response.flush

		' WRITE DATA
		Do While NOT oSchema.EOF
			For Each fldLoop in oSchema.Fields
				response.write chr(34) & trim(fldLoop.Value) & chr(34) & ","
			Next
			response.write vbcrlf
			response.flush

			oSchema.MoveNext
		Loop
	
	Else

		' NO DATA

	End If

	Set oSchema = Nothing

End Sub
%>
