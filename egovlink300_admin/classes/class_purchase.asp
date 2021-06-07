<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_purchase.asp
' AUTHOR: Steve Loar
' CREATED: 03/28/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module moves items in the cart to the class list and payment tables.
'
' MODIFICATION HISTORY
' 1.0	03/28/06 Steve Loar - Initial Version
' 2.0	04/27/07 Steve Loar -  Overhauled for Menlo Park Project
' 2.1	01/08/09 David Boyer - Added "DisplayRosterPublic" fields for Craig,CO custom team registration
' 1.5 11/19/09 David Boyer - Added "pants size" to team registration section
' 1.6 11/19/09 David Boyer - Now pull team/pants sizes from database
' 1.7	04/07/2010	Steve Loar - No more regatta team members, added team group size
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim sSql, iPaymentId, nAmount, sStatus, nTotal, bIsParent, iClassTypeId, iAdminLocationId, iCartId
	Dim iCitizenId, iAdminUserId, iJournalEntryTypeID, sItemType, x, iMaxPaymentTypes, iLedgerId, sNotes
	Dim sBuyOrWait, iRegattaTeamId, sShipToCity, sShipToState, sShipToZip
	Dim dShipping, dSalesTax, dSalesTaxRate, bInStateOnly, iMerchandiseOrderId, sShipToName, sShipToAddress

	nTotal = 0

	If Not CartHasItems() Then 
		'This is a final check that there are items in the cart before processing. This seems to have happened once.
		response.redirect "class_cart.asp"
	End If 

	iPaymentLocationId = request("PaymentLocationId") 

	'this is where the admin person is working today
	If session("LocationId") <> "" Then 
  		iAdminLocationId = session("LocationId")
	Else 
		iAdminLocationId = 0 
	End If 

	iCitizenId   = request("iUserId") ' Purchasing citizen (Head of Household)
	iAdminUserId = Session("UserID")

	'response.write "amount = [" & request("amount") & "]<br />" & vbcrlf

	sAmount = CDbl(request("amount")) ' Payment total

	iJournalEntryTypeID = GetJournalEntryTypeID( "purchase" )

	sNotes = dbsafe(request("notes"))

	'Insert the egov_class_payment row (Journal)
	iPaymentId = MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes )

	'This is for the Payment Ledger Data
	If sAmount > 0 Then 
	 	'Loop through each payment and make a ledger entry and a payment Info row
		 'Get the max payment id then loop thru and do inserts for those that have payment amounts
  		iMaxPaymentTypes = GetmaxPaymentTypeId( Session("Orgid") )
  		x = 1

  		Do While x <= iMaxPaymentTypes
			If request("amount" & x) <> "" Then 
				If HasChecks( x ) Then 
					'Check
					sCheck            = "'" & request("checkno") & "'"
					iCitizenAccountId = "NULL"
					sPlusMinus        = "+"
					cPriorBalance     = "NULL"
					iAccountId        = GetPaymentAccountId( Session("Orgid"), x )
				Else 
					sCheck = "NULL"

					If HasCitizensAccounts( x ) Then ' this is amount4
						iCitizenAccountId = request("accountid")	' this is the pick of family member accounts to get the payment from
						iAccountId        = iCitizenAccountId
						sPlusMinus        = "-"
						cPriorBalance     = GetCitizenCurrentBalance( iCitizenAccountId )

						'Debit the account that was the source of the funds
						AdjustCitizenAccountBalance iCitizenAccountId, "debit", request("amount" & x)
					Else 
						iCitizenAccountId = "NULL"
						sPlusMinus        = "+"
						cPriorBalance     = "NULL"
						iAccountId        = GetPaymentAccountId( Session("Orgid"), x )
						'Charge and Cash
					End If 
				End If 

				'Make the ledger entry for the payment - In class_global_functions.asp 
				iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iPaymentId, CDbl(request("amount" & x)), "NULL", "debit", sPlusMinus, "NULL", 1, x, cPriorBalance, "NULL" )

				'Make the entry in the egov_verisign_payment_information table - This is in ../includes/common.asp
				InsertPaymentInformation iPaymentId, iLedgerId, x, CDbl(request("amount" & x)), "APPROVED", sCheck, iCitizenAccountId
			End If 

			x = x + 1
		Loop   
	End If 

	response.write "<br />Purchase Success"
	'response.end

	'Get the Items in the cart
	sSql = "SELECT cartid, classid, userid, isnull(familymemberid,0) as familymemberid, quantity, amount, optionid, ISNULL(classtimeid,0) AS classtimeid, "
	sSql = sSql & " pricetypeid, buyorwait, sessionid, orgid, isparent, classtypeid, itemtypeid, isdropin, dropindate, "
	sSql = sSql & " rostergrade, rostershirtsize, rosterpantssize, rostercoachtype, rostervolunteercoachname, rostervolunteercoachdayphone, "
	sSql = sSql & " rostervolunteercoachcellphone, rostervolunteercoachemail, isregatta, ISNULL(regattateamid,0) AS regattateamid, "
	sSql = sSql & " shiptoname, shiptoaddress, shiptocity, shiptostate, shiptozip "
	sSql = sSql & " FROM egov_class_cart "
	sSql = sSql & " WHERE sessionid = " & session.sessionid

	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSql, Application("DSN"), 0, 1

	If Not oCart.EOF Then 

 		'Loop through the items in the cart
  		Do While Not oCart.EOF
			iCartId      = oCart("cartid")
			iUserId      = oCart("userid")
			iClassId     = oCart("classid")
			iTimeId      = oCart("classtimeid")
			iQuantity    = oCart("quantity")
			bIsParent    = oCart("isparent")
			iClassTypeId = oCart("classtypeid")
			iItemTypeId  = oCart("itemtypeid")
			sItemType    = GetItemType( iItemTypeId )  'Get the type of thing in the cart

			If oCart("isdropin") Then 
				iIsDropIn = 1
			Else 
				iIsDropIn = 0
			End If 

			If oCart("dropindate") <> "" Then 
				sDropInDate = "'" & oCart("dropindate") & "'"
			Else 
				sDropInDate = "NULL"
			End If 

			Select Case sItemType		'For each type of item, process it differently
				Case "recreation activity"

					iFamilymemberId = oCart("familymemberid")

					If iFamilymemberId = 0 Then 
						iFamilymemberId = "NULL"
						iAttendeeUserId = iUserId  'Attendee is person who bought it
					Else 
						iAttendeeUserId = GetAttendeeUserId( iFamilymemberId )
					End If 

					If oCart("buyorwait") = "W" Then 
						'This is for the wait list 
						nAmount    = 0.00  'The waitlist is free
						sStatus    = "WAITLIST"
						sBuyOrWait = "W"
					Else 
						nAmount = GetCartItemPrice( iCartId )
						If oCart("isdropin") Then 
							sStatus = "DROPIN"
						Else 
							sStatus = "ACTIVE"
						End If 
						sBuyOrWait = "B"
					End If 

					'Validate Team Registration Fields.
					iRosterGrade                   = "NULL"
					iRosterShirtSize               = "NULL"
					iRosterPantsSize               = "NULL"
					iRosterCoachType               = "NULL"
					iRosterVolunteerCoachName      = "NULL"
					iRosterVolunteerCoachDayPhone  = "NULL"
					iRosterVolunteerCoachCellPhone = "NULL"
					iRosterVolunteerCoachEmail     = "NULL"

					If oCart("rostergrade") <> "" Then 
						iRosterGrade = "'" & dbready_string(oCart("rosterGrade"),2) & "'"
					End If 

					If oCart("rostershirtsize") <> "" Then 
						iRosterShirtSize = "'" & dbready_string(oCart("rostershirtsize"),50) & "'"
					End If

					If oCart("rosterpantssize") <> "" Then 
						iRosterPantsSize = "'" & dbready_string(oCart("rosterpantssize"),50) & "'"
					End If 

					If oCart("rostercoachtype") <> "" Then 
						iRosterCoachType = "'" & dbready_string(oCart("rosterCoachType"),50) & "'"
					End If 

					If oCart("rostervolunteercoachname") <> "" Then 
						iRosterVolunteerCoachName = "'" & dbready_string(oCart("rosterVolunteerCoachName"),100) & "'"
					End If 

					If oCart("rostervolunteercoachdayphone") <> "" Then 
						iRosterVolunteerCoachDayPhone = "'" & dbready_string(oCart("rosterVolunteerCoachDayPhone"),10) & "'"
					End If 

					If oCart("rostervolunteercoachcellphone") <> "" Then 
						iRosterVolunteerCoachCellPhone = "'" & dbready_string(oCart("rosterVolunteerCoachCellPhone"),10) & "'"
					End If 

					If oCart("rostervolunteercoachemail") <> "" Then 
						iRosterVolunteerCoachEmail = "'" & dbready_string(oCart("rosterVolunteerCoachEmail"),100) & "'"
					End If 

					'Add to the Class List
					iClassListId = AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, nAmount, iPaymentId, _
					iAttendeeUserId, iIsDropIn, sDropInDate, iRosterGrade, iRosterShirtSize, iRosterPantsSize, iRosterCoachType, _
					iRosterVolunteerCoachName, iRosterVolunteerCoachDayPhone, iRosterVolunteerCoachCellPhone, _
					iRosterVolunteerCoachEmail)

					'Add to egov_journal_item_status
					CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, sStatus, sBuyOrWait

					'c
					'If oCart("buyorwait") <> "W" Then
					AddClassLedgerRows Session("Orgid"), iPaymentId, iCartId, iItemTypeId, iClassListId, "credit", "+"
					'End If 

					If bIsParent And clng(iClassTypeId) = 1 Then 
						AddToChildren iUserId, iClassId, sStatus, iQuantity, iFamilymemberId, iPaymentId, iAttendeeUserId, iIsDropIn, _
						sDropInDate, iRosterGrade, iRosterShirtSize, iRosterPantsSize, iRosterCoachType, iRosterVolunteerCoachName, _
						iRosterVolunteerCoachDayPhone, iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail
					End If 

				Case "regatta team"
					' Adding Regatta Teams
					iFamilymemberId = "NULL"
					iAttendeeUserId = iUserId  'Attendee is person who bought it
					sStatus = "ACTIVE"
					sBuyOrWait = "B"
					nAmount = GetCartItemPrice( iCartId )

					'Add to the Class List
					iClassListId = AddToClassList(iUserId, iClassId, sStatus, iQuantity, "NULL", iFamilymemberId, nAmount, iPaymentId, _
						iAttendeeUserId, iIsDropIn, sDropInDate, "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL")

					'Add to egov_journal_item_status
					CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, sStatus, sBuyOrWait

					'Create ledger row
					AddClassLedgerRows Session("Orgid"), iPaymentId, iCartId, iItemTypeId, iClassListId, "credit", "+"

					' Create the team row and captain row in the team members
					iRegattaTeamId = CreateRegattaTeam( iCartId, iClassListId ) 

					' Put the team id on the classlist
					sSql = "UPDATE egov_class_list SET regattateamid = " & iRegattaTeamId & " WHERE classlistid = " & iClassListId
					RunSQLStatement sSql

					' Add the team member rows
					'iTeamMemberCount = AddRegattaTeamMembers( iCartId, iClassListId, iRegattaTeamId )
					'iTeamMemberCount = iTeamMemberCount + 1 ' Add the captain to the member count

					' Update the member count on the team row
					'sSql = "UPDATE egov_regattateams SET membercount = " & iTeamMemberCount & " WHERE regattateamid = " & iRegattaTeamId
					'RunSQLStatement sSql

				Case "regatta member"
					' Adding Regatta Team members
					iFamilymemberId = "NULL"
					iAttendeeUserId = iUserId  'Attendee is person who bought it
					sStatus = "ACTIVE"
					sBuyOrWait = "B"
					nAmount = GetCartItemPrice( iCartId )
					iTeamId = CLng(oCart("regattateamid"))

					'Add to the Class List
					iClassListId = AddToClassList(iUserId, iClassId, sStatus, iQuantity, "NULL", iFamilymemberId, nAmount, iPaymentId, _
						iAttendeeUserId, iIsDropIn, sDropInDate, "NULL", "NULL", "NULL", "NULL", "NULL", "NULL", "NULL")

					'Add to egov_journal_item_status
					CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, sStatus, sBuyOrWait

					'Create ledger row
					AddClassLedgerRows Session("Orgid"), iPaymentId, iCartId, iItemTypeId, iClassListId, "credit", "+"

					' Put the team id on the classlist
					sSql = "UPDATE egov_class_list SET regattateamid = " & iTeamId & " WHERE classlistid = " & iClassListId
					'response.write sSql & "<br /><br />"
					RunSQLStatement sSql

					' Add the team member rows
					iTeamMemberCount = AddRegattaTeamMembers( iCartId, iClassListId, iTeamId )

					' Get the new total count of team members
					iTotalTeamMemberCount = GetRegattaTeamMemberCount( iTeamId )

					' Update the member count on the team row
					sSql = "UPDATE egov_regattateams SET membercount = " & iTotalTeamMemberCount & " WHERE regattateamid = " & iTeamId
					'response.write sSql & "<br /><br />"
					RunSQLStatement sSql

				Case "merchandise"
					dShipping = CalcShippingAndHandlingForItem( iCartId )
					If CDbl(dShipping) = CDbl(0.00)Then 
						dShipping = "NULL"
					End If 
					dSalesTax = CalcSalesTaxForItem( iCartId )
					'response.write dSalesTax & "<br /><br />"
					If CDbl(dSalesTax) = CDbl(0.00)Then 
						dSalesTax = "NULL"
					End If 
					'response.write dSalesTax & "<br /><br />"
					dSalesTaxRate = GetSalesTaxRate( bInStateOnly )
					'iFamilymemberId = "NULL"
					'iAttendeeUserId = iUserId  'Attendee is person who bought it
					'sStatus = "ACTIVE"
					'sBuyOrWait = "B"
					nAmount = oCart("amount")
					iQuantity = oCart("quantity")
					sShipToName = "'" & dbsafe(oCart( "shiptoname" )) & "'"
					sShipToAddress = "'" & dbsafe(oCart( "shiptoaddress" )) & "'"
					sShipToCity = "'" & dbsafe(oCart( "shiptocity" )) & "'"
					sShipToState = "'" & dbsafe(oCart( "shiptostate" )) & "'"
					sShipToZip = "'" & dbsafe(oCart( "shiptozip" )) & "'"

					' Add to egov_merchandiseorders - incl shipping for this merchandise and sales tax
					sSql = "INSERT INTO egov_merchandiseorders ( orgid, userid, adminuserid, shippingfee, "
					sSql = sSql & " taxrate, taxamount, orderamount, shiptoname, shiptoaddress, shiptocity, shiptostate, "
					sSql = sSql & " shiptozip, itemcount, paymentid, orderdate ) VALUES ( " & session("orgid") & ", " & iUserId & ", "
					sSql = sSql & iAdminUserId & ", " & dShipping & ", " & dSalesTaxRate & ", " & dSalesTax & ", "
					sSql = sSql & nAmount & ", " & sShipToName & ", " & sShipToAddress & ", " & sShipToCity & ", "
					sSql = sSql & sShipToState & ", " & sShipToZip & ", " & iQuantity & ", " & iPaymentId & ", "
					sSql = sSql & "dbo.GetLocalDate(" & Session("orgid") & ",GetDate()) )"

					response.write "<br /><br />"
					response.write sSql & "<br /><br />"
					iMerchandiseOrderId = RunInsertStatement( sSql )

					' Add to egov_merchandiseorderitems
					AddMerchandiseOrderItems iMerchandiseOrderId, iCartId 

					' Add to Accounts Ledger Row
					sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
					sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, merchandiseorderid ) VALUES ( "
					sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', NULL, " & nAmount & ", " & iItemTypeId & ", '+', " 
					sSql = sSql & iMerchandiseOrderId & ", 0, NULL, NULL, " & iMerchandiseOrderId & " )"
					response.write sSql & "<br /><br />"
					RunSQLStatement sSql

				Case "shipping and handling fees"
					nAmount = oCart("amount")
					' Add to Accounts Ledger Row
					sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
					sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, merchandiseorderid ) VALUES ( "
					sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', NULL, " & nAmount & ", " & iItemTypeId & ", '+', " 
					sSql = sSql & "NULL, 0, NULL, NULL, NULL )"
					response.write sSql & "<br /><br />"
					RunSQLStatement sSql

				Case "sales tax"
					nAmount = oCart("amount")
					' Add to Accounts Ledger Row
					sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
					sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, merchandiseorderid ) VALUES ( "
					sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', NULL, " & nAmount & ", " & iItemTypeId & ", '+', " 
					sSql = sSql & "NULL, 0, NULL, NULL, NULL )"
					response.write sSql & "<br /><br />"
					RunSQLStatement sSql

		End Select 

		oCart.MoveNext
    	Loop 

  	oCart.Close
	Set oCart = Nothing 

  	'clear the cart but keep the counts updated
  	ClearCart

End If 
 
' see if the org has the undo feature and set the session variable'
If OrgHasFeature("undo on receipt") Then
	' In ../includes/common.asp'
	SetUnDoBtnDisplay iPaymentId, True	
End If 

'take them to the receipt viewing page
response.redirect "view_receipt.asp?iPaymentId=" & iPaymentId


'------------------------------------------------------------------------------
' integer CreateRegattaTeam( iCartId, iClassListId ) 
'------------------------------------------------------------------------------
Function CreateRegattaTeam( ByVal iCartId, ByVal iClassListId ) 
	Dim sSql, oRs, sRegattaTeam, iClassId, sCaptainFirstname, sCaptainLastname, sCaptainaddress, sCaptaincity
	Dim sCaptainstate, sCaptainzip, sCaptainphone, iRegattaTeamId, iRegattaTeamGroupId
		
	sSql = "SELECT regattateam, classid, captainfirstname, captainlastname, captainaddress, "
	sSql = sSql & " captaincity, captainstate, captainzip, captainphone, regattateamgroupid "
	sSql = sSql & " FROM egov_class_cart_regattateams WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sRegattaTeam = "'" & dbsafe(oRs("regattateam")) & "'" 
		iRegattaTeamGroupId = CLng(oRs("regattateamgroupid"))
		iClassId = oRs("classid")
		sCaptainFirstname = "'" & dbsafe(oRs("captainfirstname")) & "'" 
		sCaptainLastname = "'" & dbsafe(oRs("captainlastname")) & "'" 
		sCaptainaddress = "'" & dbsafe(oRs("captainaddress")) & "'" 
		sCaptaincity = "'" & dbsafe(oRs("captaincity")) & "'" 
		sCaptainstate = "'" & dbsafe(oRs("captainstate")) & "'" 
		sCaptainzip = "'" & dbsafe(oRs("captainzip")) & "'" 
		sCaptainphone = "'" & dbsafe(oRs("captainphone")) & "'" 

		sSql = "INSERT INTO egov_regattateams ( regattateam, orgid, classid, classlistid, captainfirstname, captainlastname, captainaddress, "
		sSql = sSql & " captaincity, captainstate, captainzip, captainphone, regattateamgroupid ) VALUES ( " & sRegattaTeam & ", "
		sSql = sSql & session("orgid") & ", " & iClassId & ", " & iClassListId & ", " & sCaptainFirstname & ", " & sCaptainLastname & ", "
		sSql = sSql & sCaptainaddress & ", " & sCaptaincity & ", " & sCaptainstate & ", " 
		sSql = sSql & sCaptainzip & ", " & sCaptainphone & ", " & iRegattaTeamGroupId & " )"
		response.write sSql & "<br /><br />"

		iRegattaTeamId = RunInsertStatement( sSql )
		'response.end

		' Put the captain on the roster of members
'		sSql = "INSERT INTO egov_regattateammembers ( isteamcaptain, orgid, classlistid, regattateamid, regattateammember) VALUES ( 1, "
'		sSql = sSql & session("orgid") & ", " & iClassListId & ", " & iRegattaTeamId & ", " & sCaptainname & " )"
'		RunSQLStatement sSql
	Else
		iRegattaTeamId = 0 
	End If 

	oRs.Close
	Set oRs = Nothing 

	CreateRegattaTeam = iRegattaTeamId

End Function 


'------------------------------------------------------------------------------
' void AddClassLedgerRows( iOrgId, iPaymentId, iCartId, iItemTypeId, iClassListId, sEntryType, sPlusMinus )
'------------------------------------------------------------------------------
Sub AddClassLedgerRows( ByVal iOrgId, ByVal iPaymentId, ByVal iCartId, ByVal iItemTypeId, ByVal iClassListId, ByVal sEntryType, ByVal sPlusMinus )
	Dim sSql, oCart, iLedgerId

	'Get the cart price rows
	sSql = "SELECT pricetypeid, amount FROM egov_class_cart_price WHERE cartid = " & iCartId

	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSql, Application("DSN"), 0, 1

	Do While Not oCart.EOF
		'Pull any accountids - These may not exist, so pull seperately
		iAccountId = GetAccountId( oCart("pricetypeid"), iCartId )

		'create the ledger rows for the class accounts
		'MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPricetypeid )
		iLedgerId = MakeLedgerEntry( iOrgId, iAccountId, iPaymentId, CDbl(oCart("amount")), iItemTypeId, sEntryType, sPlusMinus, iClassListId, 0, "NULL", "NULL", oCart("pricetypeid") )
		oCart.MoveNext
	Loop 

	oCart.Close
	Set oCart = Nothing 
End Sub 


'------------------------------------------------------------------------------
' integer GetAccountId( iPriceTypeId, iCartId )
'------------------------------------------------------------------------------
Function GetAccountId( ByVal iPriceTypeId, ByVal iCartId )
	Dim sSql, oAccount

	'Get the cart price rows
	sSql = "SELECT P.accountid "
	sSql = sSql & " FROM egov_class_cart C, egov_class_pricetype_price P "
	sSql = sSql & " WHERE C.cartid = " & iCartId
	sSql = sSql & " AND C.classid = P.classid "
	sSql = sSql & " AND P.pricetypeid = " & iPriceTypeId

	Set oAccount = Server.CreateObject("ADODB.Recordset")
	oAccount.Open sSql, Application("DSN"), 0, 1

	If Not oAccount.eof Then 
  		If IsNull(oAccount("accountid")) Then 
    		GetAccountId = "NULL"
		Else 
    		GetAccountId = clng(oAccount("accountid"))
	    End If 
	 Else 
   		GetAccountId = "NULL"
	 End If 

	oAccount.Close
	Set oAccount = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer AddToClassList(iUserId, iClassId, sStatus, iQuantity, iTimeId, iFamilymemberId, fAmount, iPaymentId, iAttendeeUserId, _
'                        iIsDropIn, sDropInDate, iRosterGrade, iRosterShirtSize, iRosterCoachType, iRosterVolunteerCoachName, _
'                        iRosterVolunteerCoachDayPhone, iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail)
'------------------------------------------------------------------------------
Function AddToClassList( ByVal iUserId, ByVal iClassId, ByVal sStatus, ByVal iQuantity, ByVal iTimeId, ByVal iFamilymemberId, ByVal fAmount, ByVal iPaymentId, ByVal iAttendeeUserId, ByVal iIsDropIn, ByVal sDropInDate, ByVal iRosterGrade, ByVal iRosterShirtSize, ByVal iRosterPantsSize, ByVal iRosterCoachType, ByVal iRosterVolunteerCoachName, ByVal iRosterVolunteerCoachDayPhone, ByVal iRosterVolunteerCoachCellPhone, ByVal iRosterVolunteerCoachEmail )
	Dim sSql, oInsert

	AddToClassList = 0

	sSql = "INSERT INTO egov_class_list ("
	sSql = sSql & "userid, "
	sSql = sSql & "classid, "
	sSql = sSql & "status, "
	sSql = sSql & "quantity, "
	sSql = sSql & "classtimeid, "
	sSql = sSql & "familymemberid, "
	sSql = sSql & "amount, "
	sSql = sSql & "paymentid, "
	sSql = sSql & "attendeeuserid, "
	sSql = sSql & "isdropin, "
	sSql = sSql & "dropindate, "
	sSql = sSql & "rostergrade, "
	sSql = sSql & "rostershirtsize, "
	sSql = sSql & "rosterpantssize, "
	sSql = sSql & "rostercoachtype, "
	sSql = sSql & "rostervolunteercoachname, "
	sSql = sSql & "rostervolunteercoachdayphone, "
	sSql = sSql & "rostervolunteercoachcellphone, "
	sSql = sSql & "rostervolunteercoachemail "
	sSql = sSql & ") VALUES (" 
	sSql = sSql & iUserId                        & ", "
	sSql = sSql & iClassId                       & ", "
	sSql = sSql & "'" & sStatus                  & "', "
	sSql = sSql & iQuantity                      & ", "
	sSql = sSql & iTimeId                        & ", "
	sSql = sSql & iFamilymemberId                & ", "
	sSql = sSql & fAmount                        & ", "
	sSql = sSql & iPaymentId                     & ", "
	sSql = sSql & iAttendeeUserId                & ", "
	sSql = sSql & iIsDropIn                      & ", "
	sSql = sSql & sDropInDate                    & ", "
	sSql = sSql & iRosterGrade                   & ", "
	sSql = sSql & iRosterShirtSize               & ", "
	sSql = sSql & iRosterPantsSize               & ", "
	sSql = sSql & iRosterCoachType               & ", "
	sSql = sSql & iRosterVolunteerCoachName      & ", "
	sSql = sSql & iRosterVolunteerCoachDayPhone  & ", "
	sSql = sSql & iRosterVolunteerCoachCellPhone & ", "
	sSql = sSql & iRosterVolunteerCoachEmail
	sSql = sSql & ")"

	'response.write sSql & "<br />" & vbcrlf

	AddToClassList = RunInsertStatement( sSql )

End Function 


'------------------------------------------------------------------------------
' void  AddToChildren(iUserId, iClassId, sStatus, iQuantity, iFamilymemberId, iPaymentId, iAttendeeUserId, iIsDropIn, sDropInDate, _
'                  iRosterGrade, iRosterShirtSize, iRosterCoachType, iRosterVolunteerCoachName, iRosterVolunteerCoachDayPhone, _
'                  iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail)
'------------------------------------------------------------------------------
Sub AddToChildren( ByVal iUserId, ByVal iClassId, ByVal sStatus, ByVal iQuantity, ByVal iFamilymemberId, ByVal iPaymentId, ByVal iAttendeeUserId, ByVal iIsDropIn, ByVal sDropInDate, _
                  ByVal iRosterGrade, ByVal iRosterShirtSize, ByVal iRosterPantsSize, ByVal iRosterCoachType, ByVal iRosterVolunteerCoachName, ByVal iRosterVolunteerCoachDayPhone, _
                  ByVal iRosterVolunteerCoachCellPhone, ByVal iRosterVolunteerCoachEmail)

	Dim sSql, oRs, iChildClassListId

	sSql = "SELECT C.classid, T.timeid FROM egov_class C, egov_class_time T "
	sSql = sSql & " WHERE C.classid = T.classid  AND C.parentclassid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
  		iChildClassListId = AddToClassList(iUserId, oRs("classid"), sStatus, iQuantity, oRs("timeid"), iFamilymemberId, "NULL", _
                                       iPaymentId, iAttendeeUserId, iIsDropIn, sDropInDate, iRosterGrade, iRosterShirtSize, _
                                       iRosterPantsSize, iRosterCoachType, iRosterVolunteerCoachName, iRosterVolunteerCoachDayPhone, _
                                       iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail)
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void UpdatePaymentInfo( iPaymentId, iUserId, nTotal )
'------------------------------------------------------------------------------
Sub UpdatePaymentInfo( ByVal iPaymentId, ByVal iUserId, ByVal nTotal )
	Dim sSql, oCmd

	sSql = "UPDATE egov_class_payment SET "
	sSql = sSql & " userid = "       & iUserId & ", "
	sSql = sSql & " paymenttotal = " & nTotal
	sSql = sSql & " WHERE paymentid = " & iPaymentId

	RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

 	If Not VarType( strDB ) = vbString Then 
	   	DBsafe = strDB
 	Else 
	   	DBsafe = Replace( strDB, "'", "''" )
 	End If 

End Function 


'------------------------------------------------------------------------------
' void AddMerchandiseOrderItems iMerchandiseOrderId, iCartId 
'------------------------------------------------------------------------------
Sub AddMerchandiseOrderItems( ByVal iMerchandiseOrderId, ByVal iCartId )
	Dim sSql, oRs, iIsNoColor, iIsNoSize

	sSql = "SELECT I.merchandisecatalogid, M.merchandise, MC.merchandisecolor, MC.isnocolor, MS.merchandisesize, MS.isnosize, MS.displayorder, I.quantity, I.price "
	sSql = sSql & " FROM egov_class_cart_merchandiseitems I, egov_merchandisecatalog C, egov_merchandise M, "
	sSql = sSql & " egov_merchandisecolors MC, egov_merchandisesizes MS "
	sSql = sSql & " WHERE I.merchandisecatalogid = C.merchandisecatalogid "
	sSql = sSql & " AND C.merchandiseid = M.merchandiseid AND C.merchandisecolorid = MC.merchandisecolorid "
	sSql = sSql & " AND C.merchandisesizeid = MS.merchandisesizeid AND I.cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If oRs("isnocolor") Then
			iIsNoColor = 1
		Else
			iIsNoColor = 0
		End If 
		If oRs("isnosize") Then
			iIsNoSize = 1
		Else
			iIsNoSize = 0
		End If

		sSql = "INSERT INTO egov_merchandiseorderitems ( merchandiseorderid, orgid, merchandisecatalogid, merchandise, quantity, "
		sSql = sSql & " itemprice, merchandisecolor, merchandisesize, displayorder, isnocolor, isnosize ) VALUES ( " & iMerchandiseOrderId & ", "
		sSql = sSql & session("orgid") & ", " & oRs("merchandisecatalogid") & ", '" & dbsafe(oRs("merchandise")) & "', "
		sSql = sSql & oRs("quantity") & ", " & oRs("price") & ", '" & dbsafe(oRs("merchandisecolor")) & "', '" 
		sSql = sSql & dbsafe(oRs("merchandisesize")) & "', " & oRs("displayorder") & ", " & iIsNoColor & ", " & iIsNoSize & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub   




%>
