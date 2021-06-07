<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitlistexport.asp
' AUTHOR: Steve Loar
' CREATED: 05/29/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Export of permits, dumped to excel
'
' MODIFICATION HISTORY
' 1.0   05/29/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim sSql, oRs, sDate, sSearch, bIsArchive

	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Permits_Export_" & sDate & ".xls"

	sSql = session("PermitListSql")
	'response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<html><body><table border=""1"" cellpadding=""2"">"
		response.write vbcrlf & "<tr height=""30""><th>Permit #</th><th>Permit Type</th><th>Address/Location</th><th>Listed Owner</th><th>Applicant</th><th>Status</th><th>Status Date</th></tr>"
		response.flush

		Do While Not oRs.EOF
			If oRs("isarchive") Then
				bIsArchive = True
			Else
				bIsArchive = False 
			End If 

			response.write "<tr height=""26"">"
			response.write "<td align=""center"" nowrap=""nowrap"">"
			If Not bIsArchive Then
				response.write "&nbsp;" & GetPermitNumber( oRs("permitid") )
			Else
				response.write "&nbsp;" & oRs("permitnumberdisplay")
			End If 
			response.write "</td>"

			response.write "<td align=""left"" nowrap=""nowrap"">"
			response.write oRs("permittype") & " &ndash; " & oRs("permittypedesc")
			response.write "</td>"

			response.write "<td align=""left"">"

			Select Case oRs("locationtype")

				Case "address"
					response.write "&nbsp;" & oRs("residentstreetnumber")
					If oRs("residentstreetprefix") <> "" Then
						response.write " " & oRs("residentstreetprefix")
					End If 
					response.write " " & oRs("residentstreetname")
					If oRs("streetsuffix") <> "" Then
						response.write " " & oRs("streetsuffix")
					End If 
					If oRs("streetdirection") <> "" Then
						response.write " " & oRs("streetdirection")
					End If 
					response.write " " & oRs("residentunit")

				Case "location"
					response.write oRs("permitlocation")

				Case Else 
					response.write "&nbsp;"

			End Select

			response.write "</td>"

			response.write "<td align=""left"" nowrap=""nowrap"">"
			response.write "&nbsp;" & oRs("listedowner")
			response.write "</td>"

			response.write "<td align=""left"" nowrap=""nowrap"">"
			If Not bIsArchive Then
				response.write GetPermitApplicantName( oRs("permitid") )
			Else
				response.write GetArchiveContractor( oRs("permitid") )
			End If 
			response.write "</td>"

			If oRs("isonhold") Or oRs("isvoided") Or oRs("isexpired") Then 
				response.write "<td align=""center"">"
				If oRs("isonhold") Then 
					response.write "On Hold"
				Else
					If oRs("isvoided") Then 
						response.write "Voided"
					Else
						response.write "Expired"
					End If 
				End If 
				response.write "</td>"
				response.write "<td align=""center"">"
				If oRs("isexpired") And Not IsNull(oRs("expirationdate")) Then 
					response.write FormatDateTime(oRs("expirationdate"),2)
				Else 
					response.write GetLastLogDate( oRs("permitid") )   ' in permitcommonfunctions.asp
				End If 
				response.write "</td>"
			Else
				response.write "<td align=""center"">"
				response.write oRs("permitstatus")
				response.write "</td>"

				response.write "<td align=""center"">"
				Select Case oRs("statusdatedisplayed") 
					Case "applieddate"
						response.write FormatDateTime(oRs("applieddate"),2)
					Case "releaseddate"
						response.write FormatDateTime(oRs("releaseddate"),2)
					Case "approveddate"
						response.write FormatDateTime(oRs("approveddate"),2)
					Case "issueddate"
						response.write FormatDateTime(oRs("issueddate"),2)
					Case "completeddate"
						response.write FormatDateTime(oRs("completeddate"),2)
				End Select 
				response.write "</td>"
			End If 

			response.write "</tr>"
			response.flush
			oRs.MoveNext
		Loop 
		response.write "</table></body></html>"
		response.flush
	End If 

	oRs.Close 
	Set oRs = Nothing 



%>

<!-- #include file="permitcommonfunctions.asp" //-->
