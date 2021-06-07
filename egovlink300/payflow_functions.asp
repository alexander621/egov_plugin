<%
' TESTING LINE

ProcessCreditCardTransaction "John Stullenberger","5555555555554444","1207","100.00","1935" 

Function ProcessCreditCardTransaction(sName,sCardNumber,sExpiration,sAmount,sOrderID)
		
		' CREATE THE PAYFLOW COM CLIENT COMPONENT
		Set client = Server.CreateObject("PFProCOMControl.PFProCOMControl.1")
		Set oReturnCodes = CreateObject("Scripting.Dictionary")

		' BUILD PARAMETER LIST
		parmList = "TRXTYPE=S" ' SALE TRANSACTION - IMMEDIATELY FUND WITHDRAWAL
		parmList = parmList + "&TENDER=C" ' CREDIT CARD TRANSACTION
		parmList = parmList + "&COMMENT1=WWW.EGOVLINK.COM PURCHASE" ' COMMENT
		parmList = parmList + "&ACCT=" + sCardNumber ' SET CREDIT CARD NUMBER
		parmList = parmList + "&EXPDATE=" + sExpiration' SET CREDIT CARD EXP DATE
		parmList = parmList + "&NAME=" + sName' SET CUSTOMER's FULL NAME
		parmList = parmList + "&AMT=" + sAmount ' SET AMOUNT TO BE CHARGED
		parmList = parmList + "&PWD=ecmd_4303"  ' SET PAYFLOW PASSWORD
		parmList = parmList + "&USER=egovlink"  ' SET PAYFLOW USER
		parmList = parmList + "&VENDOR=egovlink" ' SET PAYFLOW VENDOR
		parmList = parmList + "&PARTNER=verisign" ' SET PAYFLOW PARTNER
	
		
		' CREATE TRANSACTION AND COMMUNICATE WITH PAYFLOW SERVER
		oTransaction = client.CreateContext("test-payflow.verisign.com", 443, 30, "", 0, "", "")
		sReturnCodes = client.SubmitTransaction(oTransaction, parmList, Len(parmList))
		client.DestroyContext (oTransaction)

		' PROCESS RETURN CODES 
		Do while Len(sReturnCodes) <> 0

			' GET NAME VALUE PAIR
			if InStr(sReturnCodes,"&") Then
				varString = Left(sReturnCodes, InStr(sReturnCodes , "&" ) -1)
			else
				varString = sReturnCodes
			end if
			
			' GET VALUES FOR PAIR FROM STRING
			name = Left(varString, InStr(varString, "=" ) -1) ' GET RETURN CODE NAME
			value = Right(varString, Len(varString) - (Len(name)+1)) ' GET RETURN CODE VALUE
			
			' ADD ITEMS TO DICTIONARY
			oReturnCodes.Add name,value
			'response.write name & value
			' SKIP PROCESSING & IN RETURN CODE STRING
			if Len(sReturnCodes) <> Len(varString) Then 
				sReturnCodes = Right(sReturnCodes, Len(sReturnCodes) - (Len(varString)+1))
			else
				sReturnCodes = ""
			end if

		Loop


		' PROCESS TRANSACTION RESULT
		If  clng(oReturnCodes.Item("RESULT")) = 0 Then
				' TRANSACTION SUCCEEDED
				approved = TRUE
				response.write  sOrderID & ":" & oReturnCodes.Item("AUTHCODE")  & ":" & oReturnCodes.Item("PNREF")  & ":" & oReturnCodes.Item("RESPMSG")  & ":" & sAmount 
				'ProcessSuccessfulTransaction sOrderID,oReturnCodes.Item("AUTHCODE"),oReturnCodes.Item("PNREF"),oReturnCodes.Item("RESPMSG"),sAmount 
		ElseIf clng(oReturnCodes.Item("RESULT")) < 0 Then
				' COMMUNICATION ERROR
				approved = FALSE

				' DISPLAY COMMUNICATION MESSAGE TO CUSTOMER
				 response.write "<div class=payflowmsgfail>Your credit card purchase was unable to processed because of a network communication error. Please try your transaction again later.<blockquote><font color=#000000>DSI Order Number:</font> " &sOrderID& "<br><font color=#000000>Payment Reference Number:</font> "&oReturnCodes.Item("PNREF")&" <br><font color=#000000>Description:</font> ("&oReturnCodes.Item("RESULT")&") - "&oReturnCodes.Item("RESPMSG")&" </blockquote></div>"
		ElseIf clng(oReturnCodes.Item("RESULT")) > 0 Then
				' TRANSACTION FAILED
				response.write oReturnCodes.Item("RESULT") & ":" & oReturnCodes.Item("PNREF")  & ":" & oReturnCodes.Item("RESPMSG")
				'ProcessDeclinedTransaction sOrderID,oReturnCodes.Item("RESULT"),oReturnCodes.Item("PNREF"),oReturnCodes.Item("RESPMSG")
				approved = FALSE
		End If
		
		' DESTORY OBJECTS
		Set client = Nothing
		Set oReturnCodes = Nothing
End Function


Function ProcessSuccessfulTransaction(sOrderID,sAUTHCODE,sPNREF,sRESPMSG,sAmount)
	
	' UPDATE DATABASE WITH RESULT
	Set oUpdate = Server.CreateObject("ADODB.Recordset")
	sSQL = "UPDATE orders SET datebilled='" & now() & "', status='COMPLETE',payflow_authorizationcode='"&sAUTHCODE&"',payflow_result=0,payflow_respmsg='"&sRESPMSG&"',payflow_pnref='"&sPNREF&"' WHERE orderid=" & sOrderID
	oUpdate.Open sSQL, Application("DSN") , 3, 1
	Set oUpdate = Nothing

	' DISPLAY SUCCESS MESSAGE TO CUSTOMER
    response.write "<div class=payflowmsg>Your credit card purchase was approved. Keep the following information for your records:<blockquote><font color=#000000>DSI Order Number:</font> " &sOrderID& "<br><font color=#000000>Payment Reference Number:</font> "&sPNREF&" <br><font color=#000000>Authorization Code:</font> "&sAUTHCODE&"<br><font color=#000000>Amount Charged: </font> "&formatcurrency(sAmount,2)&" </blockquote></div>"

End Function


Function ProcessDeclinedTransaction(sOrderID,sRESULT,sPNREF,sRESPMSG)

	' UPDATE DATABASE WITH RESULT
	Set oUpdate = Server.CreateObject("ADODB.Recordset")
	sSQL = "UPDATE orders SET datebilled='" & now() & "', status='INCOMPLETE',payflow_authorizationcode='XXXX-XXXX-XXXX',payflow_result=0,payflow_respmsg='"&sRESPMSG&"',payflow_pnref='"&sPNREF&"' WHERE orderid=" & sOrderID
	oUpdate.Open sSQL, Application("DSN") , 3, 1
	Set oUpdate = Nothing

	' DISPLAY DECLINED MESSAGE TO CUSTOMER
    response.write "<div class=payflowmsgfail>Your credit card purchase was declined for the following reason:<blockquote><font color=#000000>DSI Order Number:</font> " &sOrderID& "<br><font color=#000000>Payment Reference Number:</font> "&sPNREF&" <br><font color=#000000>Description:</font> ("&sRESULT&") - "&sRESPMSG&" </blockquote></div>"

End Function


Function ProcessCommunicationError()

	' UPDATE DATABASE WITH RESULT
	Set oUpdate = Server.CreateObject("ADODB.Recordset")
	sSQL = "UPDATE orders SET datebilled='" & now() & "', status='INCOMPLETE',payflow_authorizationcode='XXXX-XXXX-XXXX',payflow_result=0,payflow_respmsg='"&sRESPMSG&"',payflow_pnref='"&sPNREF&"' WHERE orderid=" & sOrderID
	oUpdate.Open sSQL, Application("DSN") , 3, 1
	Set oUpdate = Nothing

	' DISPLAY COMMUNICATION MESSAGE TO CUSTOMER
    response.write "<div class=payflowmsgfail>Your credit card purchase was unable to processed because of a network communication error:<blockquote><font color=#000000>DSI Order Number:</font> " &sOrderID& "<br><font color=#000000>Payment Reference Number:</font> "&sPNREF&" <br><font color=#000000>Description:</font> ("&sRESULT&") - "&sRESPMSG&" </blockquote></div>"
	
End Function
%>
