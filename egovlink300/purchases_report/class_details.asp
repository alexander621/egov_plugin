<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_viewpurchase.asp
' AUTHOR: Steve Loar
' CREATED: 08/01/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the details for a recreation activity purchase.
'
' MODIFICATION HISTORY
' 1.0   08/01/06   Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentId, oOrganization

Set oOrganization = New classOrganization

iPaymentId = request("ipaymentid")
iClassListId = request("iclasslistid")

'Session("RedirectPage") = "purchases_report/class_details.asp"
'Session("RedirectLang") = "Return to View Purchases"
'session("ManageURL") = ""

%>

<!-- #include file="../classes/class_global_functions.asp" //-->


<html>
<head>
	<title><%=oOrganization.GetOrgName()%> E-Gov Recreation Activity Purchase Details</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

<script language="Javascript">
<!--
	// Put any JavaScript here

//-->
</script>

</head>

<!--#Include file="../include_top.asp"-->

	<!--BEGIN:  USER REGISTRATION - USER MENU-->
<%	RegisteredUserDisplay( "../" ) %>
	<!--END:  USER REGISTRATION - USER MENU-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		<div id="receiptlinks">
			<img src="../images/arrow_2back.gif" align="absmiddle"><a href="javascript:history.go(-1)">&nbsp;Back</a><span id="printbutton"><input type="button" onclick="javascript:window.print();" value="Print" /></span>
		</div>
		

		<h3><%=oOrganization.GetOrgName()%> Recreation Activity Purchase Details</h3> 

<%
	Dim iUserId, nTotal, dPaymentDate, nRowTotal, bMultiWeeks

	If GetPaymentDetails( iPaymentId, iUserId, nTotal, dPaymentDate ) Then 

'		response.write vbcrlf & "<div id=""topright"">"
'		response.write vbcrlf & "<p>Purchase Date: " & DateValue(CDate(dPaymentDate)) & "</p>"
		'response.write vbcrlf & "<p>Purchase Total: " & FormatCurrency(nTotal) & "</p>"
'		response.write vbcrlf & "</div>"

		ShowUserInfo iUserId 

		' Get the classes for this purchase order by start date
		sSql = "Select L.classid, C.classname, C.startdate, C.enddate, C.isparent, C.classtypeid, isnull(C.parentclassid,0) as parentclassid, "
		sSql = sSql & " C.locationid, P.name as location, isnull(P.address1,' ') as address1, isnull(P.address2,' ') as address2, "
		sSql = sSql & " T.starttime, T.endtime, T.waitlistsize, L.status, L.quantity, L.amount, isnull(L.familymemberid,0) as familymemberid, "
		sSql = sSql & "L.classtimeid, CP.paymentdate, CP.paymenttotal, PL.paymentlocationname, PT.paymenttypename "
		sSql = sSql & " From egov_class_list L, egov_class C, egov_class_time T, egov_class_location P, egov_class_payment CP, egov_paymentlocations PL, egov_paymenttypes PT "
		sSql = sSql & " where C.classid = L.classid and L.classtimeid = T.timeid and P.locationid = C.locationid "
		sSql = sSql & " and L.paymentid = " & iPaymentId & " and L.classlistid = " & iClassListId & " and L.paymentid = CP.paymentid "
		sSql = sSql & " and CP.paymentlocationid = PL.paymentlocationid and CP.paymenttypeid = PT. paymenttypeid "
		sSql = sSql & " Order By C.startdate, L.classid, L.familymemberid"

		Set oDetails = Server.CreateObject("ADODB.Recordset")
		oDetails.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

		If Not oDetails.EOF Then 
			' Display the transaction details
			response.write vbcrlf & "<div class=""purchasereportshadow"">"
			response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
			response.write vbcrlf & "<tr>"
			response.write "<th align=""left"" colspan=""2"">Transaction Details</th>"
			response.write "</tr>"
			response.write vbcrlf & "<tr><td width=""20%"">Purchase Date:</td><td>" & DateValue(oDetails("paymentdate")) & "</td></tr>"
			response.write vbcrlf & "<tr><td>Payment Method:</td><td>" & oDetails("paymenttypename") & "</td></tr>"
			response.write vbcrlf & "<tr><td>Payment Location:</td><td>" & oDetails("paymentlocationname") & "</td></tr>"
			response.write vbcrlf & "<tr><td>Purchase Total:</td><td>" & FormatCurrency(oDetails("paymenttotal"),2) & "</td></tr>"
			response.write vbcrlf & "</table>"
			response.write vbcrlf & "</div>"

			Do While Not oDetails.EOF 
				' Display the purchase details
				response.write vbcrlf & "<div class=""purchasereportshadow"">"
				response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
				response.write vbcrlf & "<tr>"
				response.write "<th align=""left"">Purchased Recreation Activity</th><th>Qty</th><th align=""right"">Price</th><th align=""right"">Total</th>"
				response.write "</tr>"
				response.write vbcrlf & "<tr>"
				'show details here
				response.write "<td><h5>" & oDetails("classname") & "</h5>"
				If oDetails("isparent") And oDetails("classtypeid") = 1 Then 
					' This is the series level
					response.write " &ndash; Entire Series"
				End If 
				If clng(oDetails("familymemberid")) <> 0 Then 
					response.write "<br /> &nbsp;  &nbsp; <strong>Attendee:</strong> " & GetFamilyMemberInfo( oDetails("familymemberid") )
				End If 

				' Show the class/event dates
				response.write "<br /><br /> &nbsp;  &nbsp; <strong>Occurs:</strong><br />"
				If oDetails("startdate") <> "" Then 
					response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & MonthName(Month(oDetails("startdate"))) & " " & Day(oDetails("startdate")) & ", " & Year(oDetails("startdate"))
				End If 
				' handle enddate
				bMultiWeeks = false
				If oDetails("enddate") <> "" And oDetails("enddate") <> oDetails("startdate") Then 
					response.write " &ndash; " & MonthName(Month(oDetails("enddate"))) & " " & Day(oDetails("enddate")) & ", " & Year(oDetails("enddate"))
					If DateDiff("d", oDetails("startdate"), oDetails("enddate")) > 7 Then 
						bMultiWeeks = True
					Else 
						bMultiWeeks = false
					End If 
				End If 

				' Days of the week
				ShowDaysOfWeek oDetails("classid"), bMultiWeeks
				' Time of the class/event
				response.write "<br /> &nbsp;  &nbsp; &nbsp; &nbsp;" & oDetails("starttime") 
				If oDetails("endtime") <> oDetails("starttime") Then
					response.write " &ndash; " & oDetails("endtime")
				End If


				response.write "<br /><br /> &nbsp;  &nbsp; <strong>Location:</strong><br />" 
				response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & oDetails("location") & "<br />"
				If Trim(oDetails("address1")) <> "" Then 
					response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & oDetails("address1") & "<br />"
				End If 
				If Trim(oDetails("address2")) <> "" Then 
					response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & oDetails("address2") & "<br />"
				End If 

				response.write "</td>"
				response.write "<td align=""center"" valign=""top"">" & oDetails("quantity") & "</td>"
				If oDetails("status") = "ACTIVE" Then 
					If Not IsNull(oDetails("amount")) Then 
						response.write "<td align=""right"" valign=""top"">"
						response.write FormatCurrency(oDetails("amount")) & "</td>"
						'nRowTotal = clng(oDetails("quantity")) * CDbl(oDetails("amount"))
						nRowTotal = CDbl(oDetails("amount"))
						response.write "<td align=""right"" valign=""top"">" & FormatCurrency(nRowTotal) & "</td>"
					Else
						' No price, see if they are part of a purchased series
						If CheckIfSeries( oDetails("parentclassid"), iPaymentId ) Then
							response.write "<td colspan=""2"" align=""center"" valign=""top"">Part of Series<br />"
						Else 
							response.write "<td align=""right"" valign=""top"">&nbsp;</td><td align=""right"" valign=""top"">&nbsp;</td>"
						End If 
					End If 
				Else
					response.write "<td colspan=""2"" align=""center"" valign=""top"">"
					If CheckIfSeries( oDetails("parentclassid"), iPaymentId ) Then
						response.write "Part of Series<br />"
					End If 
					response.write "On Wait List<br />"
					response.write GetWaitPosition(oDetails("classid"), iUserId, oDetails("familymemberid")) & " of " & oDetails("waitlistsize")
					response.write "</td>"
				End If 

				' DISPLAY WAIVER LINK
				ListWaivers iOrgId, oDetails("classid"),0,0,1

				oDetails.movenext
				response.write "</tr>"
			Loop 

		End If 

		oDetails.close
		Set oDetails = Nothing
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"
	Else 
		response.write "<P>No Details could be found for the requested Order.</p>"
	End If 

	Set oOrganization = Nothing 
%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../include_bottom.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowUserInfo( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserInfo( iUserId )
	Dim oCmd, sResidentDesc, sUserType

	sUserType = GetUserResidentType(iUserid)
	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iUserid, iOrgId)
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iUserId", adInteger, 1, 4, iUserId)
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

	oUser.close
	Set oUser = Nothing
	Set oCmd = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetPaymentDetails( iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate )
'--------------------------------------------------------------------------------------------------
Function GetPaymentDetails( iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate )
	Dim sSql, oDetails

	sSql = "Select userid, paymenttotal, paymentdate from egov_class_payment where paymentid = " & iPaymentId & " and orgid = " & iOrgId

	Set oDetails = Server.CreateObject("ADODB.Recordset")
	oDetails.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oDetails.EOF Then 
		iUserId = oDetails("userid")
		nTotal = oDetails("paymenttotal")
		dPaymentDate = oDetails("paymentdate")
		GetPaymentDetails = True 
	Else 
		GetPaymentDetails = False 
	End If 

	oDetails.close
	Set oDetails = Nothing

End Function  


'--------------------------------------------------------------------------------------------------
' Function GetFamilyMemberInfo( iFamilyMemberId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMemberInfo( iFamilyMemberId )
	Dim sSql, oName
	
	If iFamilyMemberId <> 0 Then 
		sSql = "Select firstname, lastname, birthdate, relationship From egov_familymembers Where familymemberid = " & iFamilyMemberId

		Set oName = Server.CreateObject("ADODB.Recordset")
		oName.Open sSQL, Application("DSN"), adOpenKeyset, adLockOptimistic

		GetFamilyMemberInfo = oName("firstname") & " " & oName("lastname") 
		If Not IsNull(oName("birthdate")) Then 
			If UCase(oName("relationship")) = "CHILD" Then 
				GetFamilyMemberInfo = GetFamilyMemberInfo & " (" & DateDiff("yyyy", oName("birthdate"), Now()) & " Years Old)"
			End If 
		End If 

		oName.close
		Set oName = Nothing
	End If 
	
End Function  


'--------------------------------------------------------------------------------------------------
' Function GetFamilyMemberInfo( iFamilyMemberId )
'--------------------------------------------------------------------------------------------------
Function CheckIfSeries( iClassId, iPaymentId )
	Dim sSql, oClass

	sSql = "Select count(classlistid) as hits From egov_class_list Where paymentid = " & iPaymentId & " and classid = " & iClassId

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If clng(oClass("hits")) > 0 Then 
		CheckIfSeries = True
	Else 
		CheckIfSeries = False 
	End If 

	oClass.close
	Set oClass = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' SUB LISTWAIVERS(IORGID,CLASSID,BLNSHOWNAME,BLNSHOWDESCRIPTION,BLNSHOWLINK)
'--------------------------------------------------------------------------------------------------
Sub ListWaivers(iorgid,classid,blnshowname,blnshowdescription,blnshowlink)

	sSQL = "Select * from egov_class_waivers INNER JOIN egov_class_to_waivers ON egov_class_waivers.waiverid=egov_class_to_waivers.waiverid where orgid = '" & iorgid & "' AND classid='" & classid & "' order by waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly
	
	' LIST ALL WAIVER FOR ORGANIZATION
	If Not oWaiver.EOF Then
	
		response.write "<tr><td colspan=4>"

		' WAIVER TITLE
		response.write "<div class=waivertitle>This class requires the following waivers:"
		'response.write "<br><br><A href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border=0 src=""../images/adreader.gif"" hspace=10>Get Adobe Reader.</a>"
		response.write "</div>"

		Do While Not oWaiver.EOF
		
			response.write "<div class=waiver>" 

			' WAIVER NAME
			If blnshowname Then
				response.write "<div class=waivername>" & oWaiver("waivername") & "</div>"
			End If

			' WAIVER DESCRIPTION
			If blnshowdescription Then
				response.write "<div class=waiverdescription>" & oWaiver("waiverdescription") & "</div>"
			End If

			' WAIVER LINK
			If blnshowlink Then
				response.write "<div class=waiverlink>&bull; <a  href=""" & oWaiver("waiverurl") & """ target=""_NEW"" class=waiverlink>Click here to download " & Ucase(oWaiver("waivername")) & " waiver.</a></div>"
			End If
	
			response.write "</div>"

			oWaiver.MoveNext
		
		Loop 

		response.write "</td></tr>"

	End If 
	
	' CLOSE AND CLEAR OBJECTS
	oWaiver.close
	Set oWaiver = nothing

End Sub



%>
