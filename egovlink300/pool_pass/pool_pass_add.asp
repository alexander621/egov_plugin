<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" -->
<%
 Dim iPoolPassId, Item, nAmount, sDescription, sType

 iorgid = request("iorgid")

 set oMembership = New classMembership

'Determine if this is a purchase or a renewal.
'  - If the "poolpassid" is NULL then it is a "purchase".
'  - Otherwise, if there is a value then it is a "renewal".
 if trim(request("poolpassid")) <> "" then
    sPoolPassID       = trim(request("poolpassid"))
    lcl_isRenewedPass = "Y"

    getPoolPassInfo sPoolPassID, iUserId, iRateId, iMembershipId, iPeriodId, lcl_isSeasonal

 else
    sPoolPassID       = ""
    lcl_isRenewedPass = "N"
   	iUserId           = request("iuserid")
   	sUserType         = request("usertype")
   	iRateId           = request("rateid")
   	iMembershipId     = CLng(request("MembershipId"))
   	iPeriodId         = request("periodid")
    nAmount           = request("amount")
    lcl_isSeasonal    = request("isSeasonal")
 end if

 lcl_startdate      = request("startdate")
 lcl_isPunchcard    = request("isPunchcard")
 lcl_punchcardlimit = request("punchcard_limit")

'Using this global function gives us the ResidentType which we need for our "sUserType" variable.
'Also, get the cost of the poolpass
 'getRateInfo iRateId, lcl_rate_description, lcl_rate_residenttype
 getRateInfo iRateID, nAmount, sMessage, sDescription, iMaxsignups, iAttendanceTypeID, lcl_rate_residenttype, lcl_rate_residenttypedesc, _
             lcl_isPunchcard, lcl_punchcard_limit

storedAmt = nAmount
if iorgid = "228" then
	storedAmt = storedAmt * 1.035
end if


'Create the poolpass
 oMembership.MembershipPurchase iUserID, iorgid, iRateID, storedAmt, iMembershipID, iPeriodID, "creditcard", "online", sPoolPassID, lcl_startdate, iPassID

'Add the members to the poolpass
 for each Item IN request("passIncl")
	   oMembership.AddMember iPassID, Item, sPoolPassID, lcl_isPunchcard, lcl_punchcard_limit
 next

'Get the Membership Name and Period Description
 lcl_membershipname = oMembership.GetMembershipNameById(iMembershipID)
 lcl_period_desc    = oMembership.GetMembershipPeriodName(iPeriodID)
%>
<html>
<head>
<script language="Javascript">
<!--
function SubmitPayment() {
		document.frmpayment.submit();
}
//-->
</script>
</head>
<body onload="javascript:SubmitPayment();">
<form name="frmpayment" action="<%=application("PAYMENTURL")%>/<%=request("sVirtualSite")%>/recreation_payments/VERISIGN_FORM.ASP" method="post">
<!--<form name="frmpayment" action="http://dev4.egovlink.com/<%=request("sVirtualSite")%>/recreation_payments/VERISIGN_FORM.ASP" method="post"> -->
		<input type="hidden" name="iPAYMENT_MODULE" value="2" />
		<input type="hidden" name="ITEM_NUMBER" value="<%=iPassID%>" />
		<input type="hidden" name="iPoolPassId" value="<%=iPassID%>" />
<%
  if PeriodIsSeasonal(request("periodid")) then
     lcl_itemname_value = Year(lcl_startdate) & " Season " & lcl_membershipname & " Membership"
  else
     lcl_itemname_value = lcl_period_desc & "&nbsp;" & lcl_membershipname & " Membership"
  end if
%>
		<input type="hidden" name="ITEM_NAME" value="<%=lcl_itemname_value%>" />
		<input type="hidden" name="iuserid" value="<%=request("iuserid")%>" />
		<input type="hidden" name="amount" value="<%=nAmount%>" />
		<input type="hidden" name="custom_Membership Type" value="<%=lcl_membershipname%> Membership" />
		<input type="hidden" name="custom_Membership Period" value="<%=lcl_period_desc%>" />
		<input type="hidden" name="custom_Resident Type" value="<%=lcl_rate_residenttypedesc%>" />
		<input type="hidden" name="custom_Pool Pass Type" value="<%=sDescription%>" />
  <input type="hidden" name="display_membershipname" value="<%=lcl_membershipname%>" />
  <input type="hidden" name="custom_Punchcard" value="<%=lcl_isPunchcard%>" />
  <input type="hidden" name="custom_Punchcard Limit" value="<%=lcl_punchcard_limit%>" />
</form>
</body>
</html>
<%
'------------------------------------------------------------------------------
function PeriodIsSeasonal(iPeriodId)
	Dim sSql, oPeriod

	sSQL = "SELECT is_seasonal FROM egov_membership_periods WHERE periodid = " & CLng(iPeriodId)

	set oPeriod = Server.CreateObject("ADODB.Recordset")
	oPeriod.Open sSQL, Application("DSN"), 0, 1

	if not oPeriod.eof then
  		PeriodIsSeasonal = oPeriod("is_seasonal")
 else
  		PeriodIsSeasonal = False
	end if

	oPeriod.close
	set oPeriod = nothing

end function

'------------------------------------------------------------------------------
function dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
 	set rsi = Server.CreateObject("ADODB.Recordset")
	 rsi.Open sSQLi, Application("DSN"), 3, 1

  set rsi = nothing

end function
%>
