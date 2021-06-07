<!-- #include file="../egovlink300_global/includes/inc_email.asp" //-->
<%
	' This catches the return of promotional emails
	' Steve Loar - May 2006

	Dim iLeadId, oCmd, sSalesName, sSalesEmail, sLeadName, sLeadEmail, sCityState, bEmailSales
	Dim sHTMLBody, sSubject

	' If they came here with a leadid
	If request("leadid") <> "" And IsNumeric(request("leadid")) Then 
		iLeadId = clng(request("leadid"))
		' Update the click date and count
		UpdateLeadClick iLeadId

		' Now Get the contact Info and send an email
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "GetSalesLeadInfo"
			.CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iLeadId", 3, 1, 4, iLeadId)
			.Parameters.Append oCmd.CreateParameter("@SalesName", 200, 2, 255)
			.Parameters.Append oCmd.CreateParameter("@SalesEmail", 200, 2, 255)
			.Parameters.Append oCmd.CreateParameter("@LeadName", 200, 2, 255)
			.Parameters.Append oCmd.CreateParameter("@LeadEmail", 200, 2, 255)
			.Parameters.Append oCmd.CreateParameter("@CityState", 200, 2, 255)
			.Parameters.Append oCmd.CreateParameter("@EmailSales", 11, 2, 1)
			.Execute
		End With

		sSalesName = oCmd.Parameters("@SalesName").Value
		sSalesEmail = oCmd.Parameters("@SalesEmail").Value
		sLeadName = oCmd.Parameters("@LeadName").Value
		sLeadEmail = oCmd.Parameters("@LeadEmail").Value
		sCityState = oCmd.Parameters("@CityState").Value
		bEmailSales = oCmd.Parameters("@EmailSales").Value
		
		Set oCmd = Nothing

		'response.write "bEmailSales = " & bEmailSales
		If bEmailSales Then
			' email the sales person if flagged to do so
			sSubject = "A Lead has responsed to an email"
			sHTMLBody = "<p>" & sLeadName & " of " & sCityState & " " &   sLeadEmail & " clicked on website link in email</p>"
			subSendEmail sSalesEmail, sSalesName, sSalesEmail, sSubject, sHTMLBody 
		End If 
	End If 

	' Take them to the main site
	response.redirect "http://www2.egovlink.com"


'--------------------------------------------------------------------------------------------------
' Sub UpdateLeadClick( iLeadId )
'--------------------------------------------------------------------------------------------------
Sub UpdateLeadClick( iLeadId )
	' Update the sales lead as sent
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "UpdateSalesLeadClick"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iLeadId", 3, 1, 4, iLeadId)
		.Execute
	End With

	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub subSendEmail( sToEmail, sFromName, sFromEmail, sSubject, sHTMLBody )
'--------------------------------------------------------------------------------------------------
Sub subSendEmail( sToEmail, sFromName, sFromEmail, sSubject, sHTMLBody )
	sendEmail "", sToEmail, "", sSubject, sHTMLBody, "", "N"

End Sub


%>
