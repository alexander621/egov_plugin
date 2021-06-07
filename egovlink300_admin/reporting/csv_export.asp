<!--#include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CSV_EXPORT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/2007
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  CREATES CSV FILE FROM PASSED SQL STATEMENT
'
' MODIFICATION HISTORY
' 1.0   1/10/2007	JOHN STULLENBERGER - INITIAL VERSION
' 1.2	7/14/2010	Steve Loar - Made into and Exel export with formatted money columns
' 1.3	7/22/2010	Steve Loar - changes for Point and Pay displays
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oSchema, sProcessingRoute

' SET UP PAGE OPTIONS
sName = Replace(replace(Replace(replace(replace(Now(),":",""),"/",""),"AM",""),"PM","")," ","") ' NAME BASED ON DATETIME STRING
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=EGOV_EXPORT_" & sName & ".xls"


sSql = session("DISPLAYQUERY")

Set oSchema = Server.CreateObject("ADODB.Recordset")
oSchema.Open sSql, Application("DSN"), 3, 1

If Not oSchema.EOF Then
	response.write "<html>"
	
	response.write vbcrlf & "<style>  "
	response.write " .moneystyle "
	response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
	response.write vbcrlf & "</style>"

	response.write "<body><table border=""1"">"

	sProcessingRoute = LCase(GetProcessingRoute())
	
	' WRITE COLUMN HEADINGS
	response.write "<tr>"
	For Each fldLoop in oSchema.Fields
		response.write "<th>"
		If LCase(fldLoop.Name) = "transaction id" Then
			If sProcessingRoute = "pointandpay" Then
				response.write "Order Number"
			Else
				response.write fldLoop.Name
			End If 
		Else
			response.write fldLoop.Name
		End If 
		response.write "</th>"
	Next
	response.write "</tr>"
	response.flush

	sProcessingRoute = LCase(GetProcessingRoute())
	' WRITE DATA
	Do While Not oSchema.EOF
		response.write "<tr>"
		For Each fldLoop in oSchema.Fields
			response.write "<td"
				If (InStr(fldLoop.Value,".") > 0 And IsNumeric(fldLoop.Value))  Or (fldLoop.Type = 6) Then 
					' If this is a money field then show the decimals 
					response.write " class=""moneystyle"""
				End If 
			response.write ">" & Trim(fldLoop.Value) & "</td>"
		Next
		response.write "</tr>"
		response.flush
		oSchema.MoveNext
	Loop

	response.write "</table></body></html>"

Else

	' NO DATA

End If

Set oSchema = Nothing



%>
