<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_CASHCHECK_RECEIPT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/14/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   02/14/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityScheduleId 

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "citizen rec purchases" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iFacilityScheduleId = CLng(request("iFacilityScheduleId"))
%>

<html>
<head>
	<title>E-Gov Facility Rental Details</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="javascript">
	<!--

		window.onload = function()
		{
			factory.printing.header = "Facility Rental - Printed on &d"
			factory.printing.footer = "&bFacility Rental - Printed on &d - Page:&p/&P"
			factory.printing.portrait     = true;
			factory.printing.leftMargin   = 0.5;
			factory.printing.topMargin    = 0.5;
			factory.printing.rightMargin  = 0.5;
			factory.printing.bottomMargin = 0.5;

			// enable control buttons
			var templateSupported = factory.printing.IsTemplateSupported();
			var controls = idControls.all.tags("input");
			for ( i = 0; i < controls.length; i++ ) 
			{
				controls[i].disabled = false;
				if (templateSupported && controls[i].className == "ie55" )
					controls[i].style.display = "inline";
			}
		}

	//-->
	</script> 

</head>
<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
	<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
</div>

<object id="factory" viewastext  style="display:none"
  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
   codebase="../includes/smsx.cab#Version=6,3,434,12">
</object>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN: PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<div id="receiptlinks">
		<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)""><%=langBackToStart%></a><span id="printbutton"><input type="button" class="button" onclick="javascript:window.print();" value="Print" /></span>
	</div>

	<h3><%=Session("sOrgName")%> Facility Rental</h3>

	<% DisplayReciept iFacilityScheduleId %>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>



<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayReciept( iFacilityScheduleId )
'--------------------------------------------------------------------------------------------------
Sub DisplayReciept( ByVal iFacilityScheduleId )
	Dim sSql, oRs 

	' Get payment information 
	sSql = "SELECT amount, checkindate, checkintime, checkoutdate, checkouttime, paymenttype, paymentlocation, "
	sSql = sSql & " P.datecreated, P.facilityid, facilityname, lesseeid, ISNULL(processingfee,0.00) AS processingfee, "
	sSql = sSql & " ISNULL(sva,'') AS sva, ISNULL(P.ordernumber,'') AS ordernumber "
	sSql = sSql & " FROM egov_facilityschedule P, egov_facility F"
	sSql = sSql & " WHERE P.facilityid = F.facilityid AND P.facilityscheduleid = " & iFacilityScheduleId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 

		' Display the user information
		ShowUserInfo oRs("lesseeid")

		' TRANSACTION RESULT DETAILS
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write "<tr><th colspan=""2"" align=""left""><b>Transaction Details</b></th></tr>"
		response.write "<tr><td width=""20%"">Purchase Date:</td><td>" & DateValue(oRs("datecreated")) & "</td></tr>"
		response.write "<tr><td>Payment Method:</td><td> " & GetPaymentTypeName( oRs("paymenttype") ) & " </td></tr>"
		response.write "<tr><td>Payment Location:</td><td> " & GetPaymentLocationName( oRs("paymentlocation") ) & " </td></tr>"
		response.write "<tr><td>Amount:</td><td> " & FormatCurrency(oRs("amount"),2) & "</td></tr>"
		If oRs("sva") <> "" Then
			response.write "<tr><td>Processing Fee:</td><td> " & FormatCurrency(oRs("processingfee"),2) & "</td></tr>"
			response.write "<tr><td>Total Charges:</td><td> " & FormatCurrency( (CDbl(oRs("amount")) + CDbl(oRs("processingfee"))),2) & "</td></tr>"
			response.write "<tr><td>Order Number:</td><td> " & oRs("ordernumber") & "</td></tr>"
			response.write "<tr><td>SVA:</td><td> " & oRs("sva") & "</td></tr>"
		End If 
		response.write "</table></div>"
		
		' PRODUCT INFORMATION
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write "<tr><th colspan=""2"" align=""left""><b>Facility Rental Information</b></th></tr>"
		response.write "<tr><td width=""20%"">Order Number:</td><td> " & iFacilityScheduleId & "F3000 </td></tr>"
		response.write "<tr><td>Facility:</td><td>" & oRs("facilityname") & "</td></tr>"
		response.write "<tr><td valign=""top"">Rental Period: </td><td>" & oRs("checkindate") & " " & oRs("checkintime") & " &ndash; " & oRs("checkoutdate") & " " & oRs("checkouttime")&  "</td></tr>"
		response.write "</table></div>"
	Else
		response.write "<p>No details are available.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' string GetPaymentTypeName( iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentTypeName( ByVal iPaymentTypeId )
	Dim sSql, oRs

	' SELECT PAYMENT TYPE NAME
	sSql = "SELECT paymenttypename FROM egov_paymenttypes WHERE paymenttypeid = " & iPaymentTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	' IF PAYMENT TYPE ID FOUND
	If Not oRs.EOF Then
		'SET RETURN VALUE TO PAYMENT NAME
		GetPaymentTypeName = oRs("paymenttypename")
	Else
		GetPaymentTypeName = "UNKNOWN"
	End If

	' CLEAN UP OBJECT
	oRs.Close
	Set oRs = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' string GetPaymentLocationName( iPaymentLocationId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentLocationName( ByVal iPaymentLocationId )
	Dim sSql, oRs

	' SELECT PAYMENT TYPE NAME
	sSql = "SELECT paymentlocationname FROM egov_paymentlocations WHERE paymentlocationid = " & iPaymentLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	' IF PAYMENT TYPE ID FOUND
	If Not oRs.EOF Then
		'SET RETURN VALUE TO PAYMENT NAME
		GetPaymentLocationName = oRs("paymentlocationname")
	Else
		GetPaymentLocationName = "UNKNOWN"
	End If

	' CLEAN UP OBJECT
	oRs.Close 
	Set oRs = Nothing
			
End Function




'--------------------------------------------------------------------------------------------------
' Sub ShowUserInfo( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserInfo( ByVal iUserId )
	Dim oCmd, sResidentDesc, sUserType, oRs

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 1, 4, iUserId)
	    Set oRs = .Execute
	End With

	response.write vbcrlf & "<div class=""purchasereportshadow"">"
	response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
	response.write vbcrlf & "<tr><th colspan=""2"" align=""left"">Purchaser Contact Information</th></tr>"
	response.write vbcrlf & "<tr><td width=""20%"" valign=""top"">Name:</td><td>" & oRs("userfname") & " " & oRs("userlname")
	response.write "</td></tr>"
	response.write vbcrlf & "<tr><td>Email:</td><td>" & oRs("useremail") & "</td></tr>"
	response.write vbcrlf & "<tr><td>Phone:</td><td>" & FormatPhone(oRs("userhomephone")) & "</td></tr>"
	response.write vbcrlf & "<tr><td valign=""top"">Address:</td><td>" & oRs("useraddress") & "<br />" 
	If oRs("useraddress2") = "" Then 
		response.write oRs("useraddress2") & "<br />" 
	End If 
	response.write oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "</td></tr>"
	response.write vbcrlf & "</table></div>"

	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' string FormatPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( ByVal sNumber )

	If Len(sNumber) = 10 Then
		FormatPhone = "(" & Left(sNumber,3) & ") " & Mid(sNumber, 4, 3) & "-" & Right(sNumber,4)
	Else
		FormatPhone = sNumber
	End If

End Function


%>


