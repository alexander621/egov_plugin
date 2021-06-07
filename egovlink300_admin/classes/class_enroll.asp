<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_enroll.asp
' AUTHOR: Steve Loar
' CREATED: 03/22/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This handles the signup process for classes and events and makes a receipt.
'
' MODIFICATION HISTORY
' 1.0   03/22/06   Steve Loar - Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>

<!-- #include file="../includes/common.asp" //-->

<html>
<head>


<title>E-Gov Class/Event Receipt</title>

<%
	Dim iUserId, iClassId, iTimeId, iPaymentId, iPaymentLocationId, iPaymentTypeId, iFamilymemberId
	Dim iQuantity, fAmount, sStatus, iClassListId, iPriceTypeId, iChildClassListId

	iClassId = request("classid")
	iUserId = request("userid")
	iTimeId = request("timeid")
	iPriceTypeId = request("pricetypeid")

	If clng(request("optionid")) = 2 Then 
		' Ticket Event
		iFamilymemberId = "NULL"
		iQuantity = clng(request("quantity"))
	Else
		' Registration Required
		iFamilymemberId = request("familymemberid")
		iQuantity = 1
	End If 
	
	If request("buyorwait") = "W" Then 
		' This is for the wait list 
		response.write "This is for the wait list <br />"
		iPaymentLocationId = "NULL"
		iPaymentTypeId = "NULL" 
		fAmount = "NULL"
		sStatus = "WAITLIST"
	Else 
		' this is for purchases
		response.write "This is a purchase <br />"
		iPaymentLocationId = request("PaymentLocationId") 
		iPaymentTypeId = request("PaymentTypeId") 
		' get the amount
		If clng(iPriceTypeId) = 0 Then 
			' Other Price
			fAmount = request("amount") 
		Else
			fAmount = GetAmount( iPriceTypeId, iClassId ) 
		End If
		fAmount = CCur(fAmount) * iQuantity
		sStatus = "ACTIVE"
	End If 

	' Insert the class_payment row
	iPaymentId = MakeClassPayment( iPaymentLocationId, iPaymentTypeId )
	response.write "iPaymentId = " & iPaymentId & "<br />"

	
	' Add to the Single Event Class List
	iClassListId = AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, fAmount, iPaymentId)

	' Increment egov_class_time counts
	UpdateClassTime iTimeId, iQuantity, request("buyorwait")


	If request("isparent") = "True" And clng(request("classtypeid")) = 1 Then 
		' Get the Series children and add to their Class Lists and quantities
		sSql = "Select C.classid, T.timeid From egov_class C, egov_class_time T "
		sSql = sSql & " Where C.classid = T.classid and C.parentclassid = " & iClassId

		Set oChild = Server.CreateObject("ADODB.Recordset")
		oChild.Open sSQL, Application("DSN"), 3, 1

		Do While Not oChild.EOF
			' Update the children's quantities
			iChildClassListId = AddToClassList(iUserId, oChild("classid"), sStatus, iQuantity, oChild("timeid"), iFamilymemberId, "NULL", iPaymentId)
			' Increment egov_class_time counts for the children
			UpdateClassTime oChild("timeid"), iQuantity, request("buyorwait")
			oChild.movenext
		Loop 

		oChild.close
		Set oChild = Nothing
	End If 
%>

<link rel="stylesheet" type="text/css" href="../global.css">
<link rel="stylesheet" type="text/css" href="classes.css">

</head>
<body>

<%DrawTabs tabRecreation,1%>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		<a href="<%=Session("RedirectPage")%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Purchase Another <%=request("classname")%></a><br /><br />
		<a href="class_list.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Purchase a different Class/Event</a><br /><br />
		
		<h3><%=request("classname")%></h3><br /><br />

		Class/Event Dates:<br /><br />

		Class/Event Time:<br /><br />

		Class Location:<br /><br />

		Purchaser: &nbsp; <% = getClassPurchaserName( iUserId ) %><br /><br />

		Purchase Date: &nbsp; <%=DateValue(Now())%><br /><br />

		<% If clng(request("optionid")) = 2 Then %>
			Ticket Qty: &nbsp; <%=iQuantity%> <br /><br />
		<% Else %>
			Person Enrolled: &nbsp; <% =getFamilyMemberName( iFamilymemberId ) %><br /><br />
		<% End If %>

		<% If request("buyorwait") = "B" Then %>
				Total Cost: &nbsp; <%= FormatCurrency(fAmount) %><br /><br />
		<% Else %>
				Addition to the waitlist was successful<br /><br />
		<% End If %>

	</div>
</div>


<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>


</html>

<!--#Include file="class_global_functions.asp"-->  

	
<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' Function getClassPurchaserName( iUserId )
'--------------------------------------------------------------------------------------------------
Function getClassPurchaserName( iUserId )
	Dim sSql, oName

	sSql = "Select userfname, userlname From egov_users Where userid = " & iUserId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 1, 3

	getClassPurchaserName = oName("userfname") & " " & oName("userlname")

	oName.close
	Set oName = Nothing


End Function 


'--------------------------------------------------------------------------------------------------
' Sub  UpdateClassTime( iTimeId, iQuantity, sBuyorwait )
'--------------------------------------------------------------------------------------------------
Sub UpdateClassTime( iTimeId, iQuantity, sBuyorwait )
	Dim sSql, sField, oTime, iQty

	If sBuyorwait = "B" Then
		sSql = "Select timeid, enrollmentsize From egov_class_time Where timeid = " & iTimeId
	Else
		sSql = "Select timeid, waitlistsize From egov_class_time Where timeid = " & iTimeId
	End If 
	'response.write sSQL & "<br /><br />"

	' Open a recordset and update the quantity
	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.CursorLocation = 3
	oTime.Open sSQL, Application("DSN"), 1, 3

	If sBuyorwait = "B" Then
		iQty = oTime("enrollmentsize")
		oTime("enrollmentsize") = (iQty + clng(iQuantity))
	Else 
		iQty = oTime("waitlistsize")
		oTime("waitlistsize") = (iQty + clng(iQuantity))
	End If 
	
	oTime.Update
	oTime.close
	Set oTime = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function MakeClassPayment( iPaymentLocationId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function MakeClassPayment( iPaymentLocationId, iPaymentTypeId )
	Dim sSql, oInsert

	MakeClassPayment = 0

	sSql = "Insert into egov_class_payment (paymentlocationid, paymenttypeid) Values (" & iPaymentLocationId & ", " & iPaymentTypeId & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	'response.write sSQL & "<br /><br />"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSql, Application("DSN") , 3, 3
	MakeClassPayment = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, fAmount, iPaymentId)
'--------------------------------------------------------------------------------------------------
Function AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, fAmount, iPaymentId)
	Dim sSql, oInsert

	AddToClassList = 0

	sSql = "Insert into egov_class_list (userid, classid, status, quantity, classtimeid, familymemberid, amount, paymentid) Values (" 
	sSql = sSql & iUserId & ", " & iClassId & ", '" & sStatus & "', " & iQuantity & ", " & iTimeId & ", " 
	sSql = sSql & iFamilymemberId & ", " & fAmount & ", " & iPaymentId & " )"

	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	'response.write sSQL & "<br /><br />"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	
	oInsert.Open sSql, Application("DSN") , 3, 3
	AddToClassList = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAmount( iPriceTypeId, iClassId )
'--------------------------------------------------------------------------------------------------
Function GetAmount( iPriceTypeId, iClassId )
	Dim sSql, oAmount

	sSql = "Select amount from egov_class_pricetype_price where classid = " & iClassId & " and pricetypeid = " & iPriceTypeId
	'response.write sSQL & "<br /><br />"

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 3, 1

	GetAmount = oAmount("amount")

	oAmount.close
	Set oAmount = Nothing

End Function 


Sub GetLocation( iClassid, bIsParent )
	Dim sSql, oAvail

	If bIsParent Then 
		' Get the availability of the children events
		sSql = "Select T.timeid, T.starttime, T.endtime, isnull(T.min,0) as min, isnull(T.max,0) as max, "
		sSql = sSql & " T.enrollmentsize, T.waitlistsize, C.startdate"
		sSql = sSql & " from egov_class_time T, egov_class C"
		sSql = sSql & " where C.parentclassid = " & iClassid
		sSql = sSql & " and C.classid = T.classid"
		sSql = sSql & " order by C.startdate, T.starttime"
	Else 
		' Get the single event
		sSql = "Select T.timeid, T.starttime, T.endtime, isnull(T.min,0) as min, isnull(T.max,0) as max,"
		sSql = sSql & " T.enrollmentsize, T.waitlistsize, C.startdate"
		sSql = sSql & " from egov_class_time T, egov_class C where C.classid = " & iClassid & " and C.classid = T.classid order by T.starttime"
	End If 

	Set oAvail = Server.CreateObject("ADODB.Recordset")
	oAvail.Open sSQL, Application("DSN"), 3, 1
	
	If Not oAvail.EOF Then 
		response.write vbcrlf & "<table id=""tableavail"" border=""0"" cellpadding=""2"" cellspacing""0"">"
		response.write vbcrlf & "<caption>Availability</caption>"
		response.write vbcrlf & "<tr><th>Date</th><th>Time</th><th>Min</th><th>Max</th><th>Enrolled</th><th>Waiting</th></tr>"
		Do While Not oAvail.EOF
			response.write vbcrlf & "<tr><td>" 
			response.write DatePart("m",oAvail("startdate")) & "/" & DatePart("d",oAvail("startdate")) & "</td><td>"
			response.write oAvail("starttime") 
			If oAvail("endtime") <> oAvail("starttime") Then
				response.write "&ndash;" & oAvail("endtime")
			End If 
			response.write "</td>"
			response.write "<td align=""center"">" 
			If clng(oAvail("min")) = 0 Then
				response.write "none"
			Else
				response.write oAvail("min")
			End If 
			response.write "</td><td align=""center"">" 
			If clng(oAvail("max")) = 0 Then
				response.write "none"
			Else
				response.write oAvail("max")
			End If
			response.write "</td><td align=""center"">" & oAvail("enrollmentsize") & "</td>"
			' get waitlist here
			response.write "<td align=""center"">" & oAvail("waitlistsize") & "</td>"
			response.write "</tr>"
			oAvail.movenext 
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oAvail.close
	Set oAvail = Nothing

End Sub 
%>
