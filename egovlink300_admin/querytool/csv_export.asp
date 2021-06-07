<%
server.scripttimeout = 4800
Response.ContentType = "application/msexcel"
Response.AddHeader "Content-Disposition", "attachment;filename=report.csv"
DrawFieldSelection()
%>


<%
Function DrawFieldSelection()
	
	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = session("FULLQUERY")

	if sSQL = "" then sSQL = "SELECT * FROM Consumer WHERE 1=2"
	'oSchema.Open sSQL, request.cookies("DSN"), 3, 1
	If INSTR(UCASE(sSQL),"QDF2") <> 0 Then
		'oSchema.Open sSQL, request.cookies("DSN"), 3, 1
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	Else
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	End If


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

	End If

	Set oSchema = Nothing

End Function
%>
