<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: poolpass_receipt.asp
' AUTHOR: Steve Loar
' CREATED: 02/7/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   02/7/06   Steve Loar - Code added
' 2.0   08/01/06  Steve Loar - Changed to be part of the purchases report
' 2.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim iPoolPassId, iUserId

	sLevel = "../" ' Override of value from common.asp

	If Not UserHasPermission( Session("UserId"), "citizen rec purchases" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If 

	iPoolPassId = request("iPoolPassId")

	iUserId = GetPurchaserId( iPoolPassId )

%>

<html>
<head>
	<title>E-Gov Membership Purchase Details</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link type="text/css" media="print" rel="stylesheet" href="receiptprint.css" />

	<script language="javascript">
	<!--

		window.onload = function()
		{
			factory.printing.header = "Membership Purchase - Printed on &d"
			factory.printing.footer = "&bMembership Purchase - Printed on &d - Page:&p/&P"
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

<!--BEGIN PAGE CONTENT-->
<!--<div id="receiptcontent">-->
<div id="content">
	<div id="centercontent">

	<div id="receiptlinks">
		<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
		<!--<span id="printbutton"><input type="button" onclick="javascript:window.print();" value="Print" /></span>-->
	</div>

	<h3><%=GetCityName()%> Pool Membership Purchase</h3> 

	<% ShowUserInfo iUserId  %>

	<div class="purchasereportshadow">
		<table border="0" cellpadding="5" cellspacing="0" class="purchasereport">
		<tr><th colspan="2" align="left">Transaction Details</th></tr>
		<% ShowPassInfo iPoolPassId %>
		</table>
	</div>
	<div class="purchasereportshadow">
		<table border="0" cellpadding="5" cellspacing="0" class="purchasereport">
		<tr><th align="left" colspan="2">Pool Membership Details</th></tr>
		<% ShowPassMembers iPoolPassId %>
		</table>
	</div>

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

'------------------------------------------------------------------------------------------------------------
' void ShowPassInfo iPoolPassId 
'------------------------------------------------------------------------------------------------------------
Sub ShowPassInfo( ByVal iPoolPassId )
	Dim sSql, oRs 

	' Ths is the transaction details 
	sSql = "SELECT U.userfname, U.userlname, U.useraddress, U.useraddress2, U.usercity, U.userstate, U.userzip, "
	sSql = sSql & "P.paymentamount, P.paymenttype, P.paymentdate, P.paymentlocation, R.description, T.description AS residenttype, "
	sSql = sSql & "P.paymentauthcode, P.paymentpnref, P.paymentrespmsg, P.paymentresult, ISNULL(p.processingfee,0.00) AS processingfee, "
	sSql = sSql & "ISNULL(sva,'') AS sva, ISNULL(P.ordernumber,'') AS ordernumber "
	sSql = sSql & "FROM egov_poolpasspurchases P, egov_users U, egov_poolpassrates R, egov_poolpassresidenttypes T "
	sSql = sSql & "WHERE P.orgid = " & Session("OrgID") & " AND P.poolpassid = " & iPoolPassId 
	sSql = sSql & " AND U.userid = P.userid AND P.rateid = R.rateid AND R.residenttype = T.resident_type AND T.orgid = P.orgid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write  "<tr><td width=""20%"">Purchase Date: </td><td>" & DateValue(oRs("paymentdate")) & "</td></tr>"
		response.write  "<tr><td>Payment Method: </td><td>" & MakeProper(oRs("paymenttype")) & "</td></tr>"
		response.write  "<tr><td>Payment Location: </td><td>" & MakeProper(oRs("paymentlocation")) & "</td></tr>"
		response.write  "<tr><td>Amount: </td><td>" & FormatCurrency(oRs("paymentamount"),2) & "</td></tr>"
		If oRs("sva") <> "" Then
			response.write  "<tr><td>Processing Fee: </td><td>" & FormatCurrency(oRs("processingfee"),2) & "</td></tr>"
			response.write  "<tr><td>Amount Charged: </td><td>" & FormatCurrency((CDbl(oRs("processingfee")) + CDbl(oRs("paymentamount"))),2) & "</td></tr>"
			response.write "<tr><td>Order Number:</td><td> " & oRs("ordernumber") & "</td></tr>"
			response.write "<tr><td>SVA:</td><td> " & oRs("sva") & "</td></tr>"
		End If 
	End If
		
	oRs.close
	Set oRs = Nothing

End Sub  


'------------------------------------------------------------------------------------------------------------
' void ShowPassMembers iPoolPassId
'------------------------------------------------------------------------------------------------------------
Sub ShowPassMembers( ByVal iPoolPassId)
	Dim sSql, oRs 

	' Get the pass info
	sSql = "SELECT U.userfname, U.userlname, U.useraddress, U.useraddress2, U.usercity, U.userstate, U.userzip, "
	sSql = sSql & "P.paymentamount, P.paymenttype, P.paymentdate, P.paymentlocation, R.description, T.description as residenttype, "
	sSql = sSql & "P.paymentauthcode, P.paymentpnref, P.paymentrespmsg, P.paymentresult "
	sSql = sSql & "FROM egov_poolpasspurchases P, egov_users U, egov_poolpassrates R, egov_poolpassresidenttypes T "
	sSql = sSql & "WHERE P.orgid = " & Session("OrgID") & " AND P.poolpassid = " & iPoolPassId 
	sSql = sSql & " AND U.userid = P.userid AND P.rateid = R.rateid AND R.residenttype = T.resident_type AND T.orgid = P.orgid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write  "<tr><td width=""20%"">Member No: </td><td>" & iPoolPassId & "</td></tr>"
		response.write  "<tr><td width=""20%"">Season: </td><td>" & Year(oRs("paymentdate")) & "</td></tr>"
		response.write  "<tr><td>Membership Type: </td><td>" & oRs("residenttype") & " &mdash; " & oRs("description") & "</td></tr>"
	End If 
	oRs.close
	Set oRs = Nothing

	' Get the pass members
	sSql = "SELECT F.firstname, F.lastname, F.relationship FROM egov_familymembers F, egov_poolpassmembers P "
	sSql = sSql & "WHERE P.poolpassid = " & iPoolPassId & " AND F.familymemberid = P.familymemberid "
	sSql = sSql & "ORDER BY birthdate, lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<tr><td colspan=""2""><strong>This Membership includes</strong></td></tr>"
	Do While Not oRs.eof 
		response.write vbcrlf & "<tr><td width=""20%"">" & oRs("firstname") & " " & oRs("lastname") & "</td><td>" & TranslateMember(oRs("relationship")) &"</td></tr>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' string GetCityName()
'------------------------------------------------------------------------------------------------------------
Function GetCityName()
	Dim sSql, oRs

	GetCityName = ""

	sSql = "SELECT orgname FROM organizations WHERE orgid = " & Session("OrgID")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetCityName = oRs("orgname")
	End If
		
	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------------------------------------
' string MakeProper( sString )
'------------------------------------------------------------------------------------------------------------
Function MakeProper( ByVal sString )

	If sString = "" Then
		MakeProper = ""
	Else
		MakeProper = UCase(Left(sString,1)) & LCase(Mid(sString,2))
	End If 

End Function 


'------------------------------------------------------------------------------------------------------------
' string TranslateMember( sRelationship )
'------------------------------------------------------------------------------------------------------------
Function TranslateMember( ByVal sRelationship )

	If UCase(sRelationship) = "YOURSELF" Then
		TranslateMember = "Purchaser"
	Else 
		TranslateMember = sRelationship
	End If 
	
End Function 


'--------------------------------------------------------------------------------------------------
' void ShowUserInfo iUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowUserInfo( ByVal iUserId )
	Dim oCmd, sResidentDesc, sUserType

	sUserType = GetUserResidentType(iUserid)
	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iUserid, Session("OrgID"))
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 1, 4, iUserId)
	    Set oUser = .Execute
	End With

	response.write vbcrlf & "<div class=""purchasereportshadow"">"
	response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
	response.write vbcrlf & "<tr><th colspan=""2"" align=""left"">Purchaser Contact Information</th></tr>"
	response.write vbcrlf & "<tr><td width=""20%"" valign=""top"">Name:</td><td>" & oUser("userfname") & " " & oUser("userlname")
	response.write "<br /><strong>" & sResidentDesc & "</strong>"
	response.write "</td></tr>"
	response.write vbcrlf & "<tr><td>Email:</td><td>" & oUser("useremail") & "</td></tr>"
	response.write vbcrlf & "<tr><td>Phone:</td><td>" & FormatPhone(oUser("userhomephone")) & "</td></tr>"
	response.write vbcrlf & "<tr><td valign=""top"">Address:</td><td>" & oUser("useraddress") & "<br />" 
	If oUser("useraddress2") = "" Then 
		response.write oUser("useraddress2") & "<br />" 
	End If 
	response.write oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & "</td></tr>"
	response.write vbcrlf & "</table></div>"

	oUser.Close
	Set oUser = Nothing
	Set oCmd = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetPurchaserId( iPoolPassId )
'--------------------------------------------------------------------------------------------------
Function GetPurchaserId( ByVal iPoolPassId )
	Dim sSql, oRs

	GetPurchaserId = 0

	sSql = "SELECT P.userid FROM egov_poolpasspurchases P "
	sSql = sSql & " WHERE P.orgid = " & Session("OrgID") & " AND P.poolpassid = " & iPoolPassId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.eof Then
		GetPurchaserId = oRs("userid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>


