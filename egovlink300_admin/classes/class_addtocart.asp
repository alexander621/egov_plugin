<!--#Include file="../includes/common.asp"-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_addtocart.asp
' AUTHOR: Steve Loar
' CREATED: 03/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds items to the shopping cart and increments the class/event size.
'
' MODIFICATION HISTORY
' 1.0 03/27/06	 Steve Loar  - Initial Version
' 1.1 08/25/06	 Steve Loar  - Re-do of discounts
' 2.0	03/28/07	 Steve Loar  - Total overhaul for Menlo Park Project
' 2.1 05/28/08  David Boyer - Added Override Discount
' 2.2 01/07/09  David Boyer - Added "DisplayRosterPublic" fields for Craig,CO custom team registration
' 2.3 11/30/09  David Boyer - Added "pants size" to team registration section
' 2.4 11/30/09  David Boyer - Now pull team/pants sizes from database
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check for org features
 lcl_orghasfeature_discounts = orghasfeature("discounts")

	Dim iUserId, iClassId, iTimeId, iPaymentId, iPaymentLocationId, iPaymentTypeId, iFamilymemberId
	Dim iQuantity, fAmount, sStatus, iClassListId, iPriceTypeId, iChildClassListId, oChild, bIsParent
	Dim iClassTypeId, sOptionType, iItemTypeId, iCartId

	iClassId     = request("classid")
	iUserId      = request("userid")
	iTimeId      = request("timeid")
	iPriceTypeId = Null   ' This is now many choices
	iOptionId    = request("optionid")
	sOptionType  = request("optiontype")
	sBuyOrWait   = request("buyorwait")
	bIsParent    = request("isparent")
	iClassTypeId = request("classtypeid")
	iItemTypeId  = request("itemtypeid")  ' Class, facility, pool, gift
	iUseOverride = request("useoverridediscount")

	' Set session variables for returning to the signup page
	session("eGovUserId") = iUserId
	session("searchname") = request("searchname")

	' If the userid does not match what is in the cart, then clear the cart
	iCartUserId = getCartUserId()
	If (CLng(iCartUserId) <> CLng(0) And CLng(iUserId) <> CLng(iCartUserId)) Then 
		RemoveAllItemsFromCart(Session.SessionID)  ' Use this to remove and reset counts
	End If 

	If sOptionType = "tickets" Then 
		'Ticket Event
		'iFamilymemberId = "NULL"
		iAttendeeUserId = CLng(iUserId) ' Purchaser is the attendee
		iFamilymemberId = GetCitizenFamilyId( iAttendeeUserId )
		iQuantity = clng(request("quantity"))
	Else
		'Registration Required
		iFamilymemberId = request("familymemberid")
		iAttendeeUserId = GetAttendeeUserId( iFamilymemberId )
		iQuantity = clng(1)
	End If 

	If sBuyOrWait = "W" Then 
		'This is for the wait list 
		'response.write "This is for the wait list <br />"
		'iPaymentLocationId = "NULL"
		'iPaymentTypeId = "NULL" 
		fAmount      = 0.00  'The waitlist is free
		response.write "fAmount = " & fAmount & "<br />"
		sStatus      = "WAITLIST"
		iPriceTypeId = "NULL"
		dDropInDate  = "NULL"
		isDropIn     = 0
	Else   ' sBuyOrWait = "B"
		fAmount      = 0.00
		iPriceTypeId = "NULL"
		'Total the price choicesFor Each Id In request("pricetypeid")
		For Each Id In request("pricetypeid")
			iPriceTypeId = Id
			fAmount      = fAmount + (CDbl(request("amount" & Id)) * iQuantity)
			response.write "fAmount = " & fAmount & "<br />"
			If request("dropindate" & Id) <> "" Then
				dDropInDate = "'" & request("dropindate" & Id) & "'"
				isDropIn    = 1
				sStatus     = "DROPIN"
			Else
				dDropInDate = "NULL"
				isDropIn    = 0
				sStatus     = "ACTIVE"
			End If
		Next
	End If

	iDisplayRosterPublic           = request("displayrosterpublic")
	iRosterGrade                   = "NULL"
	iRosterShirtSize               = "NULL"
	iRosterPantsSize               = "NULL"
	iRosterCoachType               = "NULL"
	iRosterVolunteerCoachName      = "NULL"
	iRosterVolunteerCoachDayPhone  = ""
	iRosterVolunteerCoachCellPhone = ""
	iRosterVolunteerCoachEmail     = "NULL"

	if iDisplayRosterPublic then
		'Validate Team Registration Fields.
		if request("rostergrade") <> "" then
			iRosterGrade = "'" & dbready_string(request("rosterGrade"),2) & "'"
		end if

		if request("rostershirtsize") <> "" then
			iRosterShirtSize = "'" & dbready_string(request("rostershirtsize"),50) & "'"
		end if

		if request("rosterpantssize") <> "" then
			iRosterPantsSize = "'" & dbready_string(request("rosterpantssize"),50) & "'"
		end if

		if request("rostercoachtype") <> "" then
			iRosterCoachType = "'" & dbready_string(request("rosterCoachType"),50) & "'"
		end if

		if request("rostervolunteercoachname") <> "" then
			iRosterVolunteerCoachName = "'" & dbready_string(request("rosterVolunteerCoachName"),100) & "'"
		end if

		'Volunteer Coach Day Phone
		if request("skip_volcoachday_areacode") <> "" then
			iRosterVolunteerCoachDayPhone = iRosterVolunteerCoachDayPhone & dbready_string(request("skip_volcoachday_areacode"),3)
		end if

		if request("skip_volcoachday_exchange") <> "" then
			iRosterVolunteerCoachDayPhone = iRosterVolunteerCoachDayPhone & dbready_string(request("skip_volcoachday_exchange"),3)
		end if

		if request("skip_volcoachday_line") <> "" then
			iRosterVolunteerCoachDayPhone = iRosterVolunteerCoachDayPhone & dbready_string(request("skip_volcoachday_line"),4)
		end if

		if iRosterVolunteerCoachDayPhone <> "" then
			iRosterVolunteerCoachDayPhone = "'" & dbready_string(iRosterVolunteerCoachDayPhone,10) & "'"
		else
			iRosterVolunteerCoachDayPhone = "NULL"
		end if


		'Volunteer Coach CellPhone
		if request("skip_volcoachcell_areacode") <> "" then
			iRosterVolunteerCoachCellPhone = iRosterVolunteerCoachCellPhone & dbready_string(request("skip_volcoachcell_areacode"),3)
		end if

		if request("skip_volcoachcell_exchange") <> "" then
			iRosterVolunteerCoachCellPhone = iRosterVolunteerCoachCellPhone & dbready_string(request("skip_volcoachcell_exchange"),3)
		end if

		if request("skip_volcoachcell_line") <> "" then
			iRosterVolunteerCoachCellPhone = iRosterVolunteerCoachCellPhone & dbready_string(request("skip_volcoachcell_line"),4)
		end if

		if iRosterVolunteerCoachCellPhone <> "" then
			iRosterVolunteerCoachCellPhone = "'" & dbready_string(iRosterVolunteerCoachCellPhone,10) & "'"
		else
			iRosterVolunteerCoachCellPhone = "NULL"
		end if

		'Volunteer Coach Email
		if request("rostervolunteercoachemail") <> "" then
			iRosterVolunteerCoachEmail = "'" & dbready_string(request("rosterVolunteerCoachEmail"),100) & "'"
		end if

	end if

'	fAmount = "NULL"  ' Pricing is now in egov_class_cart_price
	response.write "final fAmount = " & fAmount & "<br />"

	iCartId = AddToCart(iClassId, iUserId, iTimeId, iFamilymemberId, iQuantity, fAmount, iPriceTypeId, iOptionId , sBuyOrWait, _
                     bIsParent, iClassTypeId, iItemTypeId, isDropIn, dDropInDate, iDisplayRosterPublic, iRosterGrade, _
                     iRosterShirtSize, iRosterPantsSize, iRosterCoachType, iRosterVolunteerCoachName, iRosterVolunteerCoachDayPhone, _
                     iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail)

	If sBuyOrWait = "B" Then 
 		'Add rows to the egov_class_cart_price table for purchases
  		For Each Id In request("pricetypeid")
			if request("useOverrideDiscount" & Id) <> "" then
				lcl_override = request("useOverrideDiscount" & Id)
			else
				lcl_override = 0
			end if

			AddToCartPrice iCartId, Id, CDbl(request("amount" & Id)), (CDbl(request("amount" & Id)) * iQuantity), lcl_override
  		Next 
	Else
 		'Waitlist
  		AddToCartPrice iCartId, 0, 0.00, 0.00, 0
	End If 

	'Insert the class_payment row
	 'iPaymentId = MakeClassPayment( iPaymentLocationId, iPaymentTypeId )
	 'response.write "iPaymentId = " & iPaymentId & "<br />"

	
	'Add to the Single Event Class List
	 'iClassListId = AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, fAmount, iPaymentId)

	If isDropIn = 0 Then 
		' Increment egov_class_time counts
		UpdateClassTime iTimeId, iQuantity, sBuyOrWait

		If bIsParent And clng(request("classtypeid")) = 1 Then 
			'Get the Series children and add to their Class Lists and quantities
			sSql = "SELECT C.classid, T.timeid "
			sSql = sSql & " FROM egov_class C, egov_class_time T "
			sSql = sSql & " WHERE C.classid = T.classid "
			sSql = sSql & " AND C.parentclassid = " & iClassId

  			Set oChild = Server.CreateObject("ADODB.Recordset")
		  	oChild.Open sSql, Application("DSN"), 3, 1

  			Do While Not oChild.EOF
				'Update the children's quantities
				'iChildClassListId = AddToClassList(iUserId, oChild("classid"), sStatus, iQuantity, oChild("timeid"), iFamilymemberId, "NULL", iPaymentId)
				'Increment egov_class_time counts for the children
				UpdateClassTime oChild("timeid"), iQuantity, sBuyOrWait

				oChild.movenext
			Loop 

			oChild.Close
			Set oChild = Nothing
		End If 
	End If 

	' response.write "Successfully added to cart"

	' Recalculate the prices
	ResetCartPrices

	If lcl_orghasfeature_discounts Then 
		'Recalculate any discounts
		DetermineDiscounts
	End If 

	' Redirect to the cart viewer here
	response.redirect "class_cart.asp?iUserId=" & iUserId & "&iClassId=" & iClassId & "&iTimeId=" & iTimeId
'	response.end
%>

<!--#Include file="class_global_functions.asp"--> 

<%
'------------------------------------------------------------------------------
Function AddToCart(iClassId, iUserId, iTimeId, iFamilymemberId, iQuantity, fAmount, iPriceTypeId, iOptionId , sBuyOrWait, _
                   bIsParent, iClassTypeId, iItemTypeId, isDropIn, dDropInDate, iDisplayRosterPublic, iRosterGrade, iRosterShirtSize, _
                   iRosterPantsSize, iRosterCoachType, iRosterVolunteerCoachName, iRosterVolunteerCoachDayPhone, _
                   iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail)
	Dim sSql, oInsert, iCartid, iIsParent

	If bIsParent Then 
 	 	iIsParent = 1
	Else 
	 	 iIsParent = 0
	End If 

	response.write "insert fAmount = " & fAmount & "<br />"

	sSql = "SET NOCOUNT ON;"
	sSql = sSql & " INSERT INTO egov_class_cart ("
	sSql = sSql & "classid, "
	sSql = sSql & "userid, "
	sSql = sSql & "classtimeid, "
	sSql = sSql & "familymemberid, "
	sSql = sSql & "quantity, "
	sSql = sSql & "amount, "
	sSql = sSql & "pricetypeid, "
	sSql = sSql & "optionid, "
	sSql = sSql & "buyorwait, "
	sSql = sSql & "orgid, "
	sSql = sSql & "sessionid, "
	sSql = sSql & "isparent, "
	sSql = sSql & "classtypeid, "
	sSql = sSql & "itemtypeid, "
	sSql = sSql & "dateadded, "
	sSql = sSql & "isdropin, "
	sSql = sSql & "dropindate"

	if iDisplayRosterPublic then
		sSql = sSql & ","
		sSql = sSql & "rostergrade, "
		sSql = sSql & "rostershirtsize, "
		sSql = sSql & "rosterpantssize, "
		sSql = sSql & "rostercoachtype, "
		sSql = sSql & "rostervolunteercoachname, "
		sSql = sSql & "rostervolunteercoachdayphone, "
		sSql = sSql & "rostervolunteercoachcellphone, "
		sSql = sSql & "rostervolunteercoachemail "
	end if

	sSql = sSql & ") "
	sSql = sSql & " VALUES (" 
	sSql = sSql & iClassId          & ", "
	sSql = sSql & iUserId           & ", "
	sSql = sSql & iTimeId           & ", "
	sSql = sSql & iFamilymemberId   & ", "
	sSql = sSql & iQuantity         & ", "
	sSql = sSql & fAmount           & ", "
	sSql = sSql & iPriceTypeId      & ", "
	sSql = sSql & iOptionId         & ", "
	sSql = sSql & "'" & sBuyOrWait  & "', "
	sSql = sSql & Session("OrgID")  & ", "
	sSql = sSql & Session.SessionID & ", "
	sSql = sSql & iIsParent         & ", "
	sSql = sSql & iClassTypeId      & ", "
	sSql = sSql & iItemTypeId       & ", "
	sSql = sSql & "dbo.GetLocalDate(" & Session("OrgID") & ", getdate()), "
	sSql = sSql & isDropIn          & ", "
	sSql = sSql & dDropInDate

	If iDisplayRosterPublic Then 
		sSql = sSql & ", "
		sSql = sSql & iRosterGrade                   & ", "
		sSql = sSql & iRosterShirtSize               & ", "
		sSql = sSql & iRosterPantsSize               & ", "
		sSql = sSql & iRosterCoachType               & ", "
		sSql = sSql & iRosterVolunteerCoachName      & ", "
		sSql = sSql & iRosterVolunteerCoachDayPhone  & ", "
		sSql = sSql & iRosterVolunteerCoachCellPhone & ", "
		sSql = sSql & iRosterVolunteerCoachEmail
	End If 

	sSql = sSql & ");"
	sSql = sSql & "SELECT @@IDENTITY AS ROWID;"
	
	'response.write sSql & "<br /><br />"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	iCartid = oInsert("ROWID")

	oInsert.Close
	Set oInsert = Nothing

	AddToCart = iCartid

End Function 


'------------------------------------------------------------------------------
' AddToCartPrice iCartId, iPriceTypeId, dUnitAmount, dPriceAmount, iuseOverrideDiscount
'------------------------------------------------------------------------------
Sub AddToCartPrice( ByVal iCartId, ByVal iPriceTypeId, ByVal dUnitAmount, ByVal dPriceAmount, ByVal iuseOverrideDiscount )
	Dim sSql

	sSql = "INSERT INTO egov_class_cart_price ( cartid, pricetypeid, unitprice, amount, useOverrideDiscount ) "
	sSql = sSql & " VALUES ( "
	sSql = sSql & iCartId      & ", "
	sSql = sSql & iPriceTypeId & ", "
	sSql = sSql & dUnitAmount  & ", "
	sSql = sSql & dPriceAmount & ", "
	sSql = sSql & iuseOverrideDiscount
	sSql = sSql & " )"

	'response.write "<br />" & sSql & "<br />"

	RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' string sName = getClassPurchaserName( iUserId )
'------------------------------------------------------------------------------
Function getClassPurchaserName( iUserId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname "
	sSql = sSql & "FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		getClassPurchaserName = oRs("userfname") & " " & oRs("userlname")
	Else
		getClassPurchaserName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


%>
