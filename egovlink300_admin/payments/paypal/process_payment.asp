<!-- #include file="../../../egovlink300_global/includes/inc_email.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
' 
'--------------------------------------------------------------------------------------------------
' FILENAME:  PROCESS_PAYMENT.ASP
' WRITTEN BY:  JOHN STULLENBERGER
' CREATED: 9/2/2004
' COPYRIGHT:  COPYRIGHT 2004 EC LINK, INC.  ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
' THIS PAGE RECEIVED ORDER INFORMATION POSTED FROM PAY PAY, VERIFIES, AND UPDATE EGOV DATABASE.
' 
' MODIFICATION HISTORY:
' 1.0  9/1/04	JOHN STULLENBERGER - INITIAL VERSION.
' 3.0  6/9/05	JOHN STULLENBERGER - REVISIONS FOR SINGLE CODE BASE (INSTANCE BASED SETUPS)			
'--------------------------------------------------------------------------------------------------
' 
'--------------------------------------------------------------------------------------------------
Dim iPaymentControlNumber, iOrgId

iOrgId = GetOrgIdByURL()

' LOG START OF SCRIPT
'AddtoLog " "
'AddtoLog "-----------------------------------------------------------------------------------------"
'AddtoLog "PAYMENT PROCESSING SCRIPT STARTED..."
iPaymentControlNumber = CreatePaymentsControlRow( "PAYPAL PAYMENT PROCESSING SCRIPT STARTED.", iOrgId )


' INITIALIZE VARIABLES
Dim Item_name, Item_number, Payment_status, Payment_amount
Dim Txn_id, Receiver_email, Payer_email
Dim objHttp, str


' READ POST FROM PAYPAL SYSTEM AND ADD 'CMD' FORM VALUE
'AddtoLog " - GETTING PAYMENT INFORMATION FROM PAY PAL"
AddToPaymentsLog iPaymentControlNumber, "GETTING PAYMENT INFORMATION FROM PAY PAL", iOrgId
str = "cmd=_notify-validate&" &  Request.Form


' POST BACK TO PAYPAL SYSTEM TO VALIDATE TRANSACTION
'AddtoLog " - VALIDATING PAYMENT INFORMATION WITH PAY PAL"
AddToPaymentsLog iPaymentControlNumber, "VALIDATING PAYMENT INFORMATION WITH PAY PAL", iOrgId
Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'AddtoLog " - Sending validation to PAY PAL"
AddToPaymentsLog iPaymentControlNumber, "Sending validation to PAY PAL", iOrgId
'session("paypalstr") = str
AddToPaymentsLog iPaymentControlNumber, str, iOrgId
objHttp.Send str
'session("paypalstr") = ""
'AddtoLog " - Completed Sending validation to PAY PAL"
AddToPaymentsLog iPaymentControlNumber, "Completed Sending validation to PAY PAL", iOrgId
'AddToPaymentsLog iPaymentControlNumber, "Form contains: " & str, iOrgId


' ASSIGN POSTED VARIABLES TO LOCAL VARIABLES
Item_name = request("item_name")
Item_number = request("item_number")
Payment_status = request("payment_status")
Payment_amount = request("mc_gross")
Payment_currency = request("mc_currency")
Txn_id = request("txn_id")
Receiver_email = request("receiver_email")
Payer_email = request("payer_email")


' CHECK NOTIFICATION VALIDATION
if (objHttp.status <> 200 ) then
	
	' HTTP POST FAILED
	'AddtoLog Txn_id & " - HTTP POST OPERATION PROCESS FAILED - UNABLE TO VERIFY TRANSACTION -" & objHttp.status
	AddToPaymentsLog iPaymentControlNumber, Txn_id & " HTTP POST OPERATION PROCESS FAILED - UNABLE TO VERIFY TRANSACTION - " & objHttp.status, iOrgId

elseif (objHttp.responseText = "VERIFIED") then
	
	' NOTIFICATION VERIFIED
	'AddtoLog Txn_id & " - HTTP POST OPERATION PROCESS VERIFIED"
	AddToPaymentsLog iPaymentControlNumber, "INFORMATION VERIFIED BY PAYPAL", iOrgId
	AddToPaymentsLog iPaymentControlNumber, "TransactionId = " & Txn_id, iOrgId
	AddToPaymentsLog iPaymentControlNumber, "Amount = " & Payment_amount, iOrgId
	AddToPaymentsLog iPaymentControlNumber, "Payer = " & Request("first_name") & " " & Request("last_name"), iOrgId  
	AddToPaymentsLog iPaymentControlNumber, "Payer Email = " & Payer_email, iOrgId  
	AddToPaymentsLog iPaymentControlNumber, "Option 1 = " & Request("option_selection1"), iOrgId  
	AddToPaymentsLog iPaymentControlNumber, "ServiceId = " & Item_number, iOrgId  

	' UPDATE SQL DATABASE
	'AddtoLog Txn_id & " - START UpdateDataBase"
	AddToPaymentsLog iPaymentControlNumber, "Start UpdateDataBase", iOrgId
	UpdateDataBase Txn_id 
	'AddtoLog Txn_id & " - END UpdateDataBase"
	AddToPaymentsLog iPaymentControlNumber, "End UpdateDataBase", iOrgId
	
ElseIf (objHttp.responseText = "INVALID") Then 
	
	' NOTIFICATION REPORTED AS INVALID - LOG FOR MANUAL INVESTIGATION
	'AddtoLog Txn_id & " - INVALID"
'	AddtoLog Txn_id & " - " & str
'	AddtoLog Txn_id & " - Item_name = " & request("item_name")
'	AddtoLog Txn_id & " - Item_number = " & request("item_number")
'	AddtoLog Txn_id & " - Payment_status = " & request("payment_status")
'	AddtoLog Txn_id & " - option_selection1 = " & Request("option_selection1")
	AddToPaymentsLog iPaymentControlNumber, Txn_id & " - INVALID", iOrgId
'	AddToPaymentsLog iPaymentControlNumber, Txn_id & " - " & str, iOrgId
'	AddToPaymentsLog iPaymentControlNumber, Txn_id & " - Item_name = " & request("item_name"), iOrgId
'	AddToPaymentsLog iPaymentControlNumber, Txn_id & " - Item_number = " & request("item_number"), iOrgId
'	AddToPaymentsLog iPaymentControlNumber, Txn_id & " - Payment_status = " & request("payment_status"), iOrgId
'	AddToPaymentsLog iPaymentControlNumber, Txn_id & " - option_selection1 = " & Request("option_selection1"), iOrgId
Else 
	
	' UNKNOW ERROR
	'AddtoLog "UNKNOWN ERROR"
	AddToPaymentsLog iPaymentControlNumber, "UNKNOWN ERROR", iOrgId

end if


' DESTROY OBJECTS
set objHttp = nothing


' LOG END OF SCRIPT
'AddtoLog "PAYMENT SCRIPT Successfully ENDED."
AddToPaymentsLog iPaymentControlNumber, "PAYPAL PAYMENT SCRIPT SUCCESSFULLY ENDED.", iOrgId



'----------------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' FUNCTION ADDTOLOG(STEXT)
'----------------------------------------------------------------------------------------
Sub AddtoLog(sText)
    ' WRITES SUPPLIED TEXT TO FILE WITH DATETIME
	Set oFSO = Server.Createobject("Scripting.FileSystemObject")
    Set oFile = oFSO.GetFile("C:\wwwroot\www.cityegov.com\egovlink300_admin\payments\paypal\payment.log")
    Set oText = oFile.OpenAsTextStream(8)
    oText.WriteLine (Now() & Chr(9) & sText)
    oText.Close
    
    Set oText = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing
End Sub 


'----------------------------------------------------------------------------------------
' void UpdateDatabase iPaymentRefID
'----------------------------------------------------------------------------------------
Sub UpdateDatabase( ByVal iPaymentRefID )
	Dim iOrgId, blnAddNew, oPayment, sSQL, iPaymentID, lPos, ipaymentserviceid

	'AddtoLog iPaymentRefID & " - UPDATING PAYMENT TABLE..." 
	'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - UPDATING PAYMENT TABLE", iOrgId
	
	blnAddNew = False

	' INSERT FORM INFORMATION INTO DATABASE	
	sSQL = "SELECT * FROM egov_payments WHERE paymentrefid = '" & iPaymentRefID & "'"

	Set oPayment = Server.CreateObject("ADODB.Recordset")
	oPayment.CursorLocation = 3
	oPayment.Open sSQL, Application("DSN"), 1, 3
	
	' CHECK FOR NEW RECORD
	If oPayment.EOF Then
		'AddtoLog iPaymentRefID & " - Not found in PAYMENT TABLE, adding new record..."
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Not found in PAYMENT TABLE, adding new record", iOrgId
		' ADD NEW RECORD
		oPayment.AddNew
		oPayment("paymentrefid") = Request("txn_id")
		oPayment("paymentinfoid") = Request("option_selection1")
		'AddtoLog iPaymentRefID & " - Option 1 = " & Request("option_selection1")
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Option 1 = " & Request("option_selection1"), iOrgId
		
		' The item_number is passed to PayPal and is hard coded in the table egov_PayPalOptions
		' The Service Id is before the "~" so grab them, convert to an Int and divide by 100 to get the Service Id
		' The OrgId is after the "~"  
		'AddtoLog iPaymentRefID & " - Item Number = " & Request("item_number")
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Item Number = " & Request("item_number"), iOrgId
		If Len(Request("item_number")) > 3 Then
			lPos = InStr(Request("item_number"),"~")
			If lpos > 0 Then 
				ipaymentserviceid = clng(LEFT(Request("item_number"),(lPos - 1)))
				ipaymentserviceid = ipaymentserviceid / 100
				iOrgId = Mid(Request("item_number"), (lPos + 1))
				' Handle the rare extra ~ problem
				iOrgId = Replace(iOrgId,"~","")
				iOrgId = clng(iOrgId) 
				oPayment("orgid") = iOrgId
				'iOrgId = clng(Mid(Request("item_number"),(lPos + 1)))
			Else
				' Item_number is being passed incorrectly
				ipaymentserviceid = 0
				oPayment("orgid") = 9999
				iOrgId = 9999
			End if
		Else
			' Item_number is being passed incorrectly
			ipaymentserviceid = 0
			oPayment("orgid") = 9999
			iOrgId = 9999
		End If
		
		' We now pass the orgid, so this is not needed
		' oPayment("orgid") = GetOrgid(ipaymentserviceid) ' change this function to get from item_number

		oPayment("paymentserviceid") = ipaymentserviceid
		oPayment("userid") = AddUserInformation()
		blnAddNew = True
	Else 
		'AddtoLog iPaymentRefID & " - Found in PAYMENT TABLE, updating existing record..."
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Found in PAYMENT TABLE, updating existing record", iOrgId
		' We still need an orgid
		If Len(Request("item_number")) > 3 Then
			'AddtoLog iPaymentRefID & " - Item Number = " & Request("item_number")
			'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Item Number = " & Request("item_number"), iOrgId
			lPos = InStr(Request("item_number"),"~")
			If lpos > 0 Then 
				iOrgId = Mid(Request("item_number"), (lPos + 1))
				' Handle the rare extra ~ problem
				iOrgId = Replace(iOrgId,"~","")
				iOrgId = clng(iOrgId) 
			Else
				iOrgId = 9999
			End if
		Else
			' Item_number is being passed incorrectly
			iOrgId = 9999
		End If
	End If
	'AddtoLog iPaymentRefID & " - OrgId: " & IOrgId
	'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - OrgId: " & IOrgId, iOrgId

	' UPDATE PAYMENT DATABASE FIELDS
	oPayment("paymentamount") = Request("mc_gross")
	oPayment("paymentstatus") = Request("payment_status")
	oPayment.Update
	iPaymentID = oPayment("paymentid") 
	'iOrgId = oPayment("orgid")
	'AddtoLog iPaymentRefID & " - Payment Data Saved. PaymentId: " & iPaymentID
	'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Payment Data Saved. PaymentId: " & iPaymentID, iOrgId

	oPayment.Close
	Set oPayment = Nothing

	'AddtoLog iPaymentRefID & " - oPayment Recordset closed"
	'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - oPayment Recordset closed", iOrgId

	' ADD PAYMENT DETAILS FOR NEW RECORDS
	If blnAddNew Then
		AddPaymentDetail iPaymentID 
	Else 
		'AddtoLog iPaymentRefID & " - Skipped Payment Details."
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " - Skipped Payment Details", iOrgId
	End If

	' SEND EMAIL TO CITY ADMIN
	If OrgHasFeature( iOrgId, "payment emails" ) Then 
'	If iOrgId = 30 Or iOrgId = 67 Then
		' City of Eden,NC, and Winslow,AZ Only right now
		'AddtoLog iPaymentRefID & " - Sending email to city admin."
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " -  Sending email to city admin", iOrgId
		SendPaymentEmail iPaymentID, iOrgId 
	Else 
		'AddtoLog iPaymentRefID & " - Skipped Sending Email."
		'AddToPaymentsLog iPaymentControlNumber, iPaymentRefID & " -   Skipped Sending Email", iOrgId
	End If 

End Sub 


'----------------------------------------------------------------------------------------
' FUNCTION ADDUSERINFORMATION()
'----------------------------------------------------------------------------------------
Function AddUserInformation()
	Dim sSql, oUser, iReturnValue

	'AddtoLog Request("txn_id")& " - Inserting USER INFORMATION for " & request("first_name") & " " & request("last_name") & "..."
	'AddToPaymentsLog iPaymentControlNumber, Request("txn_id")& " - Inserting USER INFORMATION for " & request("first_name") & " " & request("last_name"), iOrgId

	' INSERT FORM INFORMATION INTO DATABASE	
	iReturnValue = 0
	sSql = "SET NOCOUNT ON;Insert into egov_users (userfname, userlname, useremail) VALUES ('"
	sSql = sSql & dbsafe(request("first_name")) & "','" & dbsafe(request("last_name")) & "','"
	sSql = sSql & dbsafe(request("payer_email")) & "');SELECT @@IDENTITY AS ROWID;"
	
	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.CursorLocation = 3

	oUser.Open sSql, Application("DSN"), 1, 3
	iReturnValue = oUser("ROWID")

'	oUser.Open "egov_users", Application("DSN") , 1, 2, 2
'	oUser.AddNew
'	oUser("userfname") = dbsafe(request("first_name"))
'	oUser("userlname") = dbsafe(request("last_name"))
'	oUser("useremail") = request("payer_email")
'	oUser.Update
'	iReturnValue = oUser("userid")

	oUser.Close
	Set oUser = Nothing

	AddUserInformation = iReturnValue

End Function


'----------------------------------------------------------------------------------------
' FUNCTION ADDPAYMENTDETAIL(IPAYMENTID)
'----------------------------------------------------------------------------------------
Sub AddPaymentDetail(iPaymentID)
	Dim sDetails, oDetails

	'AddtoLog request("txn_id")& " - UPDATING PAYMENT DETAIL INFORMATION..."
	'AddToPaymentsLog iPaymentControlNumber, request("txn_id")& " - UPDATING PAYMENT DETAIL INFORMATION", iOrgId
	
	sDetails = ""
	For Each oField IN Request.Form
		sDetails = sDetails & oField & " : " & request(oField) & "</br>"
	Next 

	Set oDetails = Server.CreateObject("ADODB.Recordset")
	oDetails.CursorLocation = 3
	oDetails.Open "egov_paymentdetails", Application("DSN"), 1, 2, 2

	oDetails.AddNew
	oDetails("paymentid") = iPaymentID
	oDetails("paymentsummary") = dbsafe(sDetails)

	oDetails.Update

	oDetails.Close
	Set oDetails = Nothing 

End Sub 


'----------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'----------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'----------------------------------------------------------------------------------------
' FUNCTION SENDPAYMENTEMAIL(IPAYMENTID, iOrgId)
'----------------------------------------------------------------------------------------
sub SendPaymentEmail(iPaymentID,iOrgId)
	Dim oCdoMail, oCdoConf, sMsg, oPayment, sSql

	' CONNECT TO DATABASE AND GET PAYMENT INFORMATION
	sSql = "SELECT * FROM dbo.egov_payment_list where paymentid = " & iPaymentID & " and orgid = " & iOrgId

	Set oPayment = Server.CreateObject("ADODB.Recordset")
	oPayment.Open sSql, Application("DSN"), 3, 1
	
	If oPayment("assigned_email") = "" or isNull(oPayment("assigned_email")) then
		'adminEmailAddr = "codes@pikevillecity.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
		' Do not bother to send a mail message if there is no one to send to
		
	Else
		adminEmailAddr = oPayment("assigned_email") ' ASSIGNED ADMIN USER EMAIL
 
		' SEND EMAIL TO CITY ADMIN
		'Set oCustomFields = Server.CreateObject("ADODB.Recordset")
		'oCustomFields.CursorLocation = 3
		'cSQL = "Select payment_information FROM egov_payment_list WHERE paymentid=" & iPaymentID & " and orgid = " & iOrgId
		'oCustomFields.Open cSQL, Application("DSN") , 1, 3
		
		sMsg = sMsg & "This automated message was sent by the E-Gov web site. Do not reply to this message." & vbcrlf 
		sMsg = sMsg & " " & vbcrlf 
		sMsg = sMsg & "A payment was received on " & Date() & "." & vbcrlf 
		sMsg = sMsg & " " & vbcrlf  
		sMsg = sMsg & "---------------------------------------------------------------------------------------------------" & vbcrlf 
		sMsg = sMsg & " " & vbcrlf 
		sMsg = sMsg & "Item_name = " & Request.Form("item_name")   & vbcrlf
		sMsg = sMsg & "Item_number = " & Request.Form("item_number")   & vbcrlf
		sMsg = sMsg & "Payment_status = " & Request.Form("payment_status")   & vbcrlf
		sMsg = sMsg & "Payment_amount = " & Request.Form("mc_gross")   & vbcrlf
		sMsg = sMsg & "Payment_currency = " & Request.Form("mc_currency")   & vbcrlf
		sMsg = sMsg & "Txn_id = " & Request.Form("txn_id") & " " & vbcrlf
		sMsg = sMsg & "Receiver_email = " & Request.Form("receiver_email")   & vbcrlf
		sMsg = sMsg & "Payer_email = " & Request.Form("payer_email")   & vbcrlf
		
		'to strip out html from the additional information
		If IsNull(oPayment("payment_information")) Then
			customPaymentInfo = "No Payment Information was provided."
		Else 
			customPaymentInfo = oPayment("payment_information")
			customPaymentInfo = Replace(customPaymentInfo,"custom_","")
			' added 2 of the 3 lines below to handle different variations of the break tag - Steve Loar
			customPaymentInfo = Replace(customPaymentInfo,"</br>",vbcrlf)
			customPaymentInfo = Replace(customPaymentInfo,"<br />",vbcrlf)
			customPaymentInfo = Replace(customPaymentInfo,"<br/>",vbcrlf)
		End If 

		sMsg = sMsg & customPaymentInfo & vbcrlf


		sendEmail "", adminEmailAddr, "", "E-Gov - Payment Received", "", sMsg, "Y"


	End If
	oPayment.close
	Set oPayment = Nothing 

End Sub 



'----------------------------------------------------------------------------------------
' FUNCTION GETORGID(IPAYMENTSERVICEID)
'----------------------------------------------------------------------------------------
Function GetOrgid(iPaymentServiceid)
	Dim oOrg, iReturnValue

	iReturnValue = 0
	
	Set oOrg = Server.CreateObject("ADODB.Recordset")
	oOrg.Open "Select orgid FROM egov_organizations_to_paymentservices WHERE paymentservice_id=" & iPaymentServiceid , Application("DSN"), 1, 3
	iReturnValue = oOrg("orgid")

	oOrg.Close
	Set oOrg = Nothing

	GetOrgid = iReturnValue

End Function

'--------------------------------------------------------------------------------------------------
' FUNCTION OrgHasFeature( sFeature )
'--------------------------------------------------------------------------------------------------
Function OrgHasFeature( iOrgId, sFeature )
	Dim sSql, oFeatureAccess, blnReturnValue

	' SET DEFAULT
	OrgHasFeature = False

	' LOOKUP passed FEATURE FOR the current ORGANIZATION 
	sSql = "SELECT count(FO.featureid) as feature_count FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid and orgid = " & iOrgId & " AND F.feature = '" & sFeature & "' "
	Set oFeatureAccess = Server.CreateObject("ADODB.Recordset")
	oFeatureAccess.Open  sSQL, Application("DSN"), 3, 1
	
	If clng(oFeatureAccess("feature_count")) > 0 Then
		' the ORGANIZATION HAS the FEATURE
		OrgHasFeature = True
	End If
	
	oFeatureAccess.close 
	Set oFeatureAccess = Nothing

End Function


'------------------------------------------------------------------------------
' Function CreatePaymentsControlRow( sLogEntry, iOrgId )
'------------------------------------------------------------------------------
Function CreatePaymentsControlRow( sLogEntry, iOrgId )
	Dim sSql, iPaymentControlNumber

	sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iOrgId & ", 'Admin', 'Payments', '" & dbready_string(sLogEntry,500) & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

	CreatePaymentsControlRow = iPaymentControlNumber

End Function 


'------------------------------------------------------------------------------
' Sub AddToPaymentsLog( iPaymentControlNumber, sLogEntry, iOrgId )
'------------------------------------------------------------------------------
Sub AddToPaymentsLog( iPaymentControlNumber, sLogEntry, iOrgId  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & iOrgId & ", 'Admin', 'Payments', '" & dbready_string(sLogEntry,1000) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

End Sub 


'------------------------------------------------------------------------------
' Function dbready_string( p_value, p_length )
'------------------------------------------------------------------------------
Function dbready_string( p_value, p_length )
	Dim lcl_return

	lcl_return = ""

	If p_value <> "" And p_length <> "" Then 
		lcl_return = trim(p_value)
		lcl_return = replace(lcl_return,"<","&lt;")
		lcl_return = replace(lcl_return,">","&gt;")

		'Verify the length
		If Len(lcl_return) > p_length Then 
			lcl_return = mid(lcl_return,1,p_length)
		End If 
		lcl_return = replace(lcl_return,"'","''")
	End If 

	dbready_string = lcl_return

End Function 


'-------------------------------------------------------------------------------------------------
' Function RunInsertStatement( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunInsertStatement( sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSQL, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing

	RunInsertStatement = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' Sub RunSQLStatement( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQLStatement( sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub


'----------------------------------------------------------------------------------------
' Function GetOrgIdByURL()
'----------------------------------------------------------------------------------------
Function GetOrgIdByURL()
	Dim oOrgInfo, sSql 

	' BUILD CURRENT URL
	If request.servervariables("HTTPS") = "on" Then
		sProtocol = "https://"
	Else
		sProtocol = "http://"
	End If
	sSERVER = request.servervariables("SERVER_NAME")
	sCurrent = sProtocol & sSERVER & "/" & GetVirtualDirectyName()

	' LOOKUP CURRENT URL IN DATABASE
	sSQL = "SELECT orgid FROM Organizations WHERE OrgEgovWebsiteURL = '" & sCurrent & "'"

	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oOrgInfo.EOF Then
		GetOrgIdByURL = oOrgInfo("OrgID")
	Else
		GetOrgIdByURL = 0
	End If

	oOrgInfo.Close 
	Set oOrgInfo = Nothing 

End Function


'----------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'----------------------------------------------------------------------------------------
Function GetVirtualDirectyName()

	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 0) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = Replace(sReturnValue, "/", "")

End Function

%>


