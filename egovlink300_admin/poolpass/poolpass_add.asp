<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" //-->
<html>
<head>
<script language="javascript">
<!--
function SubmitPayment() {
  document.frmpayment.submit();
}
//-->
</script>
<%
Dim iPoolPassId, Item, nAmount, sDescription, sType, oMembership

'Determine if this is a purchase or a renewal.
'  - If the "poolpassid" is NULL then it is a "purchase".
'  - Otherwise, if there is a value then it is a "renewal".
 if trim(request("poolpassid")) <> "" then
    sPoolPassID       = trim(request("poolpassid"))
    lcl_isRenewedPass = "Y"

    getPoolPassInfo sPoolPassID, iUserId, iRateId, iMembershipId, iPeriodId, lcl_isSeasonal

   'Using this global function gives up the ResidentType which we need for our "sUserType" variable.
    getRateInfo iRateId, lcl_rate_description, lcl_rate_residenttype

    sUserType = lcl_rate_residenttype

 else
    sPoolPassID        = ""
    lcl_isRenewedPass  = "N"
   	iUserId            = request("userid")
   	sUserType          = request("usertype")
   	iRateId            = request("rateid")
   	iMembershipId      = CLng(request("iMembershipId"))
   	iPeriodId          = request("iperiodid")
    lcl_isSeasonal     = request("isSeasonal")
 end if

 lcl_notes = DBSafe(request("purchasenotes"))
 lcl_adminid = DBSafe(request.cookies("User")("userid"))

 Set oMembership = New classMembership

 oMembership.MembershipId = iMembershipId
	sRateName          = ""
	sMessage           = ""
	iMaxSignUps        = 1
	nAmount            = 001.00
	iFamilyCount       = 0
 sDescription       = ""
 sType              = ""
 lcl_startdate      = request("startdate")
 lcl_isPunchcard    = request("isPunchcard")
 lcl_punchcardlimit = request("punchcard_limit")

 'nAmount = GetRateAmount( request("rateid"), sDescription, sType, Session("OrgID") )
 getPoolPassRateAmount request("rateid"), nAmount, sDescription, sType

 'iPassId = oMembership.MembershipPurchase( request("iuserid"), request("rateid"), nAmount, request("paymenttype"), request("paymentlocation"), request("imembershipid"), request("iperiodid") )
 oMembership.MembershipPurchase iUserID, iRateID, nAmount, request("paymenttype"), request("paymentlocation"), iMembershipID, _
                                iPeriodID, sPoolPassID, lcl_startdate, iPassID

if session("orgid") = "26" then
 sSQL = "UPDATE egov_poolpasspurchases SET note = '" & lcl_notes & "', adminid = '" & lcl_adminid & "' WHERE poolpassid = " & iPassID				
	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSQL
	oCmd.Execute
	Set oCmd = Nothing
end if

 for each Item in request("passIncl")
	   oMembership.AddMember iPassID, Item, sPoolPassID, lcl_isPunchcard, lcl_punchcardlimit
 next

 response.write "</head>" & vbcrlf

'PaymentType = "CreditCard" ---------------------------------------------------
 if LCASE(request("paymenttype")) = "creditcard" then
   'Build the ITEM_NAME value
    lcl_itemname = oMembership.GetMembershipPeriodName(request("iperiodid"))
    lcl_itemname = lcl_itemname & "&nbsp;" & oMembership.GetRateName(request("rateid"))
    lcl_itemname = lcl_itemname & "&nbsp;" & oMembership.GetMembershipName()
    lcl_itemname = lcl_itemname & "&nbsp;Membership"

    response.write "<body onload=""javascript:SubmitPayment();"">" & vbcrlf
    response.write "<form name=""frmpayment"" action=""../recreation_payments/VERISIGN_FORM.ASP"" method=""post"">" & vbcrlf
    response.write "		<input type=""hidden"" name=""iPAYMENT_MODULE"" value=""2"" />" & vbcrlf
    response.write "		<input type=""hidden"" name=""ITEM_NUMBER"" value="""           & iPassId                  & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""ITEM_NAME"" value="""             & lcl_itemname             & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""amount"" value="""                & nAmount                  & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""custom_Resident Type"" value="""  & sType                    & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""custom_Pool Pass Type"" value=""" & sDescription             & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""imembershipid"" value="""         & oMembership.MembershipId & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""iperiodid"" value="""             & request("iperiodid")     & """ />" & vbcrlf
    response.write "		<input type=""hidden"" name=""rateid"" value="""                & request("rateid")        & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""isPunchcard"" value="""           & lcl_isPunchcard          & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""punchcard_limit"" value="""       & lcl_punchcardlimit       & """ />" & vbcrlf
    response.write "</form>" & vbcrlf
    response.write "</body>" & vbcrlf
 else
	   response.redirect "poolpass_receipt.asp?iPoolPassId=" & iPassId & "&isRenewedPass=" & lcl_isRenewedPass
 end if

 response.write "</html>" & vbcrlf

 set oMembership = nothing

'------------------------------------------------------------------------------
Sub UpdatePoolPass( iPoolPassId, sAuthCode, sPNRef, sResult, sRespMsg)

	Set oCmd = Server.CreateObject("ADODB.Command")
    With oCmd
	    .ActiveConnection = Application("DSN")
	    .CommandText = "UpdatePoolPassPurchase"
	    .CommandType = 4
   		.Parameters.Append oCmd.CreateParameter("@PoolPassId", 3, 1, 4, iPoolPassId)
	    .Parameters.Append oCmd.CreateParameter("@PaymentAuthCode", 200, 1, 50, sAuthCode)
   		.Parameters.Append oCmd.CreateParameter("@PaymentPNRef", 200, 1, 50, sPNRef)
   		.Parameters.Append oCmd.CreateParameter("@PaymentResult", 200, 1, 50, sResult)
   		.Parameters.Append oCmd.CreateParameter("@PaymentRespMsg", 200, 1, 255, sRespMsg)
	    .Execute, , adExecuteNoRecords
	  End With

	  Set oCmd = Nothing
End Sub  
Function DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString
End Function
%>
