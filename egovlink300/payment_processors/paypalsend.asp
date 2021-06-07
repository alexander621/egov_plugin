<!--#include file="../includes/common.asp" //-->
<!--#include file="../include_top_functions.asp" //-->
<%
'**************************************************************************
' This work product is provided "AS IS" and without warranty. 
'**************************************************************************

'Please refer to the posting regarding the HTTPS Interface on the PayPal Developer's
'Forum for the Payflow Gateway.  Go to http://www.paypaldeveloper.com and "Jump to" 
'Payflow Gateway.
'
'You will also need to reference the Payflow Pro Developer's Guide are that available
'from the Integration Center at https://www.paypal.com/IntegrationCenter/ic_payflowpro.html.

'Things you should implement are:
' 
'1. Check for DUPLICATE transaction and react accordingly.
'2. Build in some re-try logic in case you do not receive a response.  You will
'   need to take into consideration the Request ID. See the HTTPS Interface posting for
'   more information on the Request ID and DUPLICATE response.
' 
' MODIFICATION HISTORY
' 1.0	05/28/2009	Steve Loar - Initial Version
'
'**************************************************************************

Dim objWinHttp, strHTML, parmList, requestID, sGatewayHost, sFeature, iPaymentControlNumber, iOrgId, sOrgFeature

If request("cardNum") = "" Then
	' User did not enter a credit card number or this is a spider so just send back a communication error
	response.write "RESULT=-1&RESPMSG=MISSING PAYMENT INFORMATION!"
	response.End 
End If 

If request( "feature" ) <> "" Then 
	sFeature = request( "feature" )
Else 
	sFeature = "default"
End If 

'Build the parameter list
'
'This a very, very basic implementation to just how how you can post data.  What data 
'you decide to send and how your react to the response is a business decision that you
'must make.
parmList = "TENDER=C" ' Always Credit Card
parmList = parmList + "&ACCT=" + request("cardNum") 
parmList = parmList + GetVerisignOptions( iorgid, sGatewayHost, sFeature )
parmList = parmList + "&EXPDATE=" + request("cardExp")
If request("cvv2") <> "" Then 
	parmList = parmList + "&CVV2=" + Trim(request("cvv2"))
End If 
parmList = parmList + "&AMT=" + FormatNumber(request("amount"),2,,,0)
parmList = parmList + "&TRXTYPE=S" ' Always Sale
parmList = parmList + "&NAME=" + request("sjname") 
parmList = parmList + "&STREET=" + request("StreetAddress")
if request("city") <> "" then parmList = parmList + "&CITY=" + request("city")
if request("state") <> "" then parmList = parmList + "&STATE=" + request("state")
parmList = parmList + "&ZIP=" + request("ZipCode")
parmList = parmList + "&CUSTREF=" + request("ordernumber")
parmList = parmList + "&COMMENT1=" & request("comment1")
parmList = parmList + "&COMMENT2=" & request("comment2")
parmList = parmList + "&TIMEOUT=120"

If sGatewayHost <> "" Then 
	'Open Session
	Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

	' Set timeouts of resolve(0), connection(60000), send(30000), receive(30000) in milliseconds. 0 = infinite
	objWinHttp.SetTimeouts 0, 120000, 60000, 120000

	objWinHttp.Open "POST", sGatewayHost, False


	'Build the headers
	'
	'The required headers are Content-Type, X-VPS-Timeout and X-VPS-Request-ID.  All other headers are for
	'your use and you'd populate them based on your environment.
	'
	objWinHttp.setRequestHeader "Content-Type", "text/namevalue"  ' for XML, use text/xml
	objWinHttp.SetRequestHeader "Host", sEgovWebsiteURL   ' your website
	objWinHttp.SetRequestHeader "X-VPS-Timeout", "120"
	'objWinHttp.SetRequestHeader "X-VPS-VIT-Client-Architecture", "x86"  ' set to your environment
	objWinHttp.SetRequestHeader "X-VPS-VIT-Client-Type", "ASP/Classic"
	'objWinHttp.SetRequestHeader "X-VPS-VIT-Client-Version", "0.0.1"     ' your application version
	'objWinHttp.SetRequestHeader "X-VPS-VIT-Integration-Product", "Homegrown"   ' your application 
	'objWinHttp.SetRequestHeader "X-VPS-VIT-Integration-Version", "0.0.1"
	'objWinHttp.SetRequestHeader "X-VPS-VIT-OS-Name", "windows"      ' your OS
	'objWinHttp.SetRequestHeader "X-VPS-VIT-OS-Version", "2002_SP2"  ' your OS version

	' Need to generate a unique id for the request id
	requestID = generateRequestID(32)
	objWinHttp.SetRequestHeader "X-VPS-Request-ID", requestID

	If request("paymentcontrolnumber") <> "" Then
		iPaymentControlNumber = request("paymentcontrolnumber")
		iOrgId = request("orgid")
		sOrgFeature = request("orgfeature") & ": paypalsend"

		'makePaymentLogEntry iPaymentControlNumber, iOrgId, sOrgFeature, parmList
		makePaymentLogEntry iPaymentControlNumber, iOrgId, sOrgFeature, sGatewayHost
	End If 

	'if iOrgId = 5 then
		Const WinHttpRequestOption_SecureProtocols = 9
		const SecureProtocol_TLS1_2 = 2048    'TLS 1.2
		objWinHttp.Option(WinHttpRequestOption_SecureProtocols) = SecureProtocol_TLS1_2
		makePaymentLogEntry iPaymentControlNumber, iOrgId, sOrgFeature, "Using TLS 1.2"
	'end if

	'Send Parameter List
	objWinHttp.Send parmList

	' Get the text of the response.
	transResponse = objWinHttp.ResponseText

	Set objWinHttp = Nothing

	response.write transResponse
Else
	' client is not set up correctly so just send back a communication error
	response.write "RESULT=-1&RESPMSG=MISSING GATEWAY HOST URL"
End If 


'------------------------------------------------------------------------------
' Function generateRequestID( tmpLength )
'------------------------------------------------------------------------------
Function generateRequestID( tmpLength )

	Randomize Timer

  	Dim tmpCounter, tmpGUID
  	Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

  	For tmpCounter = 1 To tmpLength
    	tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  	Next

  	generateRequestID = tmpGUID

End Function


'------------------------------------------------------------------------------
' string GetVerisignOptions( iOrgId, sGatewayHost, sFeature )
'------------------------------------------------------------------------------
Function GetVerisignOptions( ByVal iOrgId, ByRef sGatewayHost, ByVal sFeature )
	Dim sSql, oRs

	GetVerisignOptions = ""

	sSql = "SELECT vendor, [user], password, partner, ISNULL(gatewayhost,'') AS gatewayhost "
	sSql = sSql & " FROM egov_verisign_options WHERE orgid = "  & iOrgId
	sSql = sSql & " AND feature = '" & Track_DBsafe(sFeature) & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetVerisignOptions = "&PWD=" & oRs("password")  ' SET PAYFLOW PASSWORD
		GetVerisignOptions = GetVerisignOptions & "&USER=" & oRs("user")  ' SET PAYFLOW USER
		GetVerisignOptions = GetVerisignOptions & "&VENDOR=" & oRs("vendor") ' SET PAYFLOW VENDOR
		GetVerisignOptions = GetVerisignOptions & "&PARTNER=" & oRs("partner") ' SET PAYFLOW PARTNER
		sGatewayHost = oRs("gatewayhost")
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' makePaymentLogEntry iPaymentControlNumber, iOrgId, sFeature, sLogEntry
'------------------------------------------------------------------------------
Sub makePaymentLogEntry( ByVal iPaymentControlNumber, ByVal iOrgId, ByVal sOrgFeature, ByVal sLogEntry )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( " & iPaymentControlNumber
	sSql = sSql & ", " & iOrgId & ", 'public', '" & sOrgFeature & "', '" & Track_DBsafe( sLogEntry ) & "' )"

	RunSQLStatement sSql 

End Sub 

%>
