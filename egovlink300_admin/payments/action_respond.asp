<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_respond.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the payments list
'
' MODIFICATION HISTORY
' 1.0   ???			???? - INITIAL VERSION
' 1.1	10/12/2006	Steve Loar - Security, Header and nav changed
' 2.0	07/27/2011	Steve Loar - Major changes to get this to work in non-IE browsers
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oRs, sSql, iTrackID, blnUpdate, lngTrackingNumber, sPaymentInfo, blnFound, sTitle, sStatus
Dim datSubmitDate, sDetails, iUserid, sTheUserid, iemployeeid

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "receipts", sLevel	' In common.asp

' GET INFORMATION FOR THIS Payment
iTrackID = CLng(request("control"))

sSql = "SELECT paymentservicename, paymentstatus, paymentdate, paymentsummary, payment_information, userid, assigned_userid "
sSql = sSql & "FROM egov_payment_list WHERE paymentid = " & iTrackID 

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

' CHECK FOR INFORMATION
If Not oRs.EOF Then
	' REQUEST FOUND GET INFORMATION	
	blnFound = True
	sTitle = oRs("paymentservicename")
	sStatus = oRs("paymentstatus")
	datSubmitDate = oRs("paymentdate")
	sDetails = oRs("paymentsummary")
	sPaymentInfo = oRs("payment_information")
	iUserid = oRs("userid")
	sTheUserid = oRs("userid")
	iemployeeid =  oRs("assigned_userid")

	If datSubmitDate <> "" Then
		lngTrackingNumber = iTrackID  & Replace(FormatDateTime(CDate(datSubmitDate),4),":","")
	Else
		lngTrackingNumber = "000000000"
	End If
Else
	' REQUEST NOT FOUND
	blnFound = False
End If

oRs.Close 
Set oRs = Nothing 

' old body tag that does not work in most browsers
' <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="toggleDisplay('contact_user');update_user_display();toggleDisplay('info');toggleDisplay('log');toggleDisplay('update_form');toggleDisplay('details')">
%>

<html>
<head>
	<title><%=langBSPayments%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="payment_styles.css" />

	<script>
	<!--

	//-->
	</script>
</head>

<body>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<div id="content">
		<div id="centercontent">

			<div id="paymentdetailstitle">
				<font size="+1"><strong>Review/Respond to Online Payment (<%=lngTrackingNumber%>)</strong></font>
				<br /><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="action_line_list.asp?useSessions=1""><%=langBackToStart%></a>
			</div>

				<div id="paymentservicename">
					Payment Service Name: &nbsp;
<%
					If sTitle <> "" Then
						response.write sTitle 
					Else
						response.write "<font color=""red"">!No online payment category name provided</font>"
					End If

					response.write "<br /><br />Payment Date: " & DateValue(CDate(datSubmitDate))
%>
				</div>
						
				<div id="user_expand">Contact Information:</div>
<%
				' GET User INFORMATION
				fnDisplayUserInfo iUserid 
%>

				<div id="detail_expand">Payment Detail Information:
				<% if session("orgid") = "153" then response.write "<a href=""edit_payment.asp?control=" & iTrackID & """>Edit</a>" %>
				</div>
<%
				'DISPLAY PAYMENT Form DETAILS
				If sPaymentInfo <> "" Then
					aPaymentinfo = Split( sPaymentInfo, "</br>")
					response.write "<table id=""detail_info"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
					For Each sRow In aPaymentinfo
						If Trim(sRow) <> "" Then 
							response.write "<tr>"
							'response.write "<td>" & sRow & "</td>"
							aCells = Split( sRow, ":" )
							response.write "<td align=""right"" class=""label"" nowrap=""nowrap"">"
							response.write Trim(Replace( aCells(0), "custom_", "" ))
							response.write ":</td>"
							response.write "<td>"
							response.write aCells(1)
							response.write "</td>"
							response.write "</tr>"
						End If 
					Next 
					response.write "</table>"
				Else
					response.write "<p><font color=red>!No payment information available!</font></p>"
				End If
%>

				<div id="transaction_expand">Payment Transaction Details:</div>
<%
				'DISPLAY PAYMENT TRANSACTION DETAILS
				If sDetails <> "" Then
					aDetailinfo = Split( sDetails, "<br>")
					response.write "<table id=""transaction_info"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
					For Each sRow In aDetailinfo
						If Trim(sRow) <> "" Then 
							response.write "<tr>"
							'response.write "<td>" & sRow & "</td>"
							aCells = Split( sRow, ":" )
							response.write "<td align=""right"" class=""label"" nowrap=""nowrap"">"
							response.write Trim(Replace( aCells(0), "custom_", "" ))
							response.write ":</td>"
							response.write "<td>"
							response.write aCells(1)
							response.write "</td>"
							response.write "</tr>"
						End If 
					Next 
					response.write "</table>"
					'response.write sDetails 
				Else
					response.write "<p><font color=""red"">!No transaction details available!</font></p>"
				End If
%>

 
		</div>
	</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void fnDisplayUserInfo iID 
'------------------------------------------------------------------------------
Sub fnDisplayUserInfo( ByVal iID )
	Dim sSql, oRs

	' GET INFORMATION FOR SPECIFIED USER
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(userbusinessname,'') AS userbusinessname, "
	sSql = sSql & "ISNULL(useremail,'') AS useremail, ISNULL(userhomephone,'') AS userhomephone, ISNULL(userfax,'') AS userfax, ISNULL(useraddress,'') AS useraddress, "
	sSql = sSql & "ISNULL(userzip,'') AS userzip, ISNULL(userstate,'') AS userstate "
	sSql = sSql & "FROM egov_users WHERE userid = " & iID

	If iID = "" Or IsNull(iID) Then 
		response.write "<div style=""display:none;"" id=""contact_user""><font color=""red""><i>No information available for specified user.</i></font></div>"
	Else 
		' OPEN RECORDSET
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1
		
		If Not oRs.EOF Then
			'response.write "<div id=""contact_user"" style=""margin-top:5px;border:solid 1px #000000;background-color:#eeeeee;"">"
			response.write "<table id=""contact_user"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
			response.write "<tr><td class=""label"" align=""right"">First Name:</td><td>" & oRs("userfname") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Last Name:</td><td>" & oRs("userlname") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Business Name:</td><td>" & oRs("userbusinessname") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Email:</td><td>" & oRs("useremail") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Daytime Phone:</td><td>" & oRs("userhomephone") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Fax:</td><td>" & oRs("userfax") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Street:</td><td>" & oRs("useraddress") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">State:</td><td>" & oRs("userstate") & "</td></tr>"
			response.write "<tr><td class=""label"" align=""right"">Zip:</td><td>" & oRs("userzip") & "</td></tr>"
			'response.write "<tr><td class=""label"" align=""right"">Country:</td><td>" & oRs("usercountry") & "</td></tr>"
			response.write "</table>"
			'response.write "</div>"
		Else
			response.write "<div id=""contact_user""><font color=""red""><i>No information available for specified user.</i></font></div>"
		End If
		 
		 oRs.Close
		 Set oRs = Nothing 

	End If 

End Sub 



%>
