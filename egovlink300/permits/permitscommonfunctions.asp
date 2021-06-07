<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitscommonfunctions.asp
' AUTHOR: Steve Loar
' CREATED: 04/22/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared functions for permits. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   04/22/2010   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' string GetAdminName( iUserId )
'------------------------------------------------------------------------------
Function GetAdminName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT firstname + ' ' + lastname AS username FROM users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetAdminName = oRs("username")
	Else
		GetAdminName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetCheckNo( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetCheckNo( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT checkno FROM egov_verisign_payment_information "
	sSql = sSql & " WHERE checkno IS NOT NULL AND paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCheckNo = oRs("checkno")
	Else
		GetCheckNo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetConstructionType( iConstructionTypeId )
'--------------------------------------------------------------------------------------------------
Function GetConstructionType( ByVal iConstructionTypeId )
	Dim sSql, oRs

	sSql = "SELECT constructiontype FROM egov_constructiontypes WHERE constructiontypeid = " & iConstructionTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetConstructionType = oRs("constructiontype")
	Else
		GetConstructionType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetInvoiceContact( iPermitContactId )
'--------------------------------------------------------------------------------------------------
Function GetInvoiceContact( ByVal iPermitContactId )
	Dim sSql, oRs, sReturn

	sSql = "SELECT ISNULL(company,'') AS company, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname "
	sSql = sSql & " FROM egov_permitcontacts WHERE permitcontactid = " & iPermitContactId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If oRs("firstname") <> "" Then 
			sReturn = oRs("firstname") & " " & oRs("lastname")
		Else
			sReturn = oRs("company")
		End If 
	Else
		sReturn = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

	GetInvoiceContact = sReturn

End Function 


'-------------------------------------------------------------------------------------------------
' double GetInvoicedTotal( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetInvoicedTotal( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(totalamount),0.00) AS totalamount FROM egov_permitinvoices "
	sSql = sSql & "WHERE isvoided = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoicedTotal =  FormatNumber(oRS("totalamount"),2,,,0)
	Else
		GetInvoicedTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'-------------------------------------------------------------------------------------------------
' double GetInvoicePaymentTotal( iInvoiceId ) 
'-------------------------------------------------------------------------------------------------
Function GetInvoicePaymentTotal( ByVal iInvoiceId ) 
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(amount),0.00) AS totalamount FROM egov_accounts_ledger "
	sSql = sSql & " WHERE invoiceid = " & iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoicePaymentTotal =  FormatNumber(oRS("totalamount"),2,,,0)
	Else
		GetInvoicePaymentTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'--------------------------------------------------------------------------------------------------
' string GetLastLogDate( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetLastLogDate( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT MAX(entrydate) AS entrydate FROM egov_permitlog WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetLastLogDate = DateValue(oRs("entrydate"))
	Else 
		GetLastLogDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetOccupancyType( iOccupancyTypeId )
'--------------------------------------------------------------------------------------------------
Function GetOccupancyType( ByVal iOccupancyTypeId )
	Dim sSql, oRs, sReturn

	sReturn = ""

	sSql = "SELECT ISNULL(usegroupcode,'') AS usegroupcode, occupancytype FROM egov_occupancytypes "
	sSql = sSql & " WHERE occupancytypeid = " & iOccupancyTypeId
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("usegroupcode") <> "" Then
			sReturn = oRs("usegroupcode") & " "
		End If 
		sReturn = sReturn & oRs("occupancytype")
	Else
		sReturn = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetOccupancyType = sReturn

End Function 


'-------------------------------------------------------------------------------------------------
' money GetPaidTotal( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPaidTotal( ByVal iPermitId )
	Dim sSql, oRs

	' Get the total paid from the Journal Table
	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal FROM egov_accounts_ledger L, egov_permitinvoices I "
	sSql = sSql & " WHERE I.isvoided = 0 AND L.invoiceid = I.invoiceid AND L.ispaymentaccount = 0 "
	sSql = sSql & " AND L.permitid = " & iPermitId
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaidTotal =  FormatNumber(oRS("paymenttotal"),2,,,0)
	Else
		GetPaidTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'--------------------------------------------------------------------------------------------------
' string = GetPermitApplicantName( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitApplicantName( ByVal iPermitId )
	Dim sSql, oRs, sApplicant

	sApplicant = ""
	sSql = "SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSQl & " ISNULL(company,'') AS company " 
	sSql = sSQl & " FROM egov_permitcontacts WHERE isapplicant = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sApplicant = oRs("firstname") & " " & oRs("lastname") & "<br />"
		End If 
		If oRs("company") <> "" And sApplicant = "" Then 
			sApplicant = oRs("company") & "<br />" 
		End If 
	Else
		sApplicant = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitApplicantName = sApplicant

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitContactDetails( iPermitId, sContactType )
'--------------------------------------------------------------------------------------------------
Function GetPermitContactDetails( ByVal iPermitId, ByVal sContactType )
	Dim sSql, oRs, sDetails

	sDetails = ""
	sSql = " SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE " & sContactType & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY company, lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		If sDetails <> "" Then 
			sDetails = sDetails & "<br /><br />"
		End If 
		If oRs("firstname") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("company") & "</strong><br />" 
		End If 
		If Trim(oRs("address")) <> "" Then 
			sDetails = sDetails & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sDetails = sDetails & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sDetails = sDetails & FormatPhoneNumber( oRs("phone") ) 
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetPermitContactDetails = sDetails

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitIsExpired( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsExpired( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT isexpired FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isexpired") Then 
			GetPermitIsExpired = True 
		Else
			GetPermitIsExpired = False 
		End If 
	Else
		GetPermitIsExpired = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitIsOnHold( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsOnHold( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT isonhold FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isonhold") Then 
			GetPermitIsOnHold = True 
		Else
			GetPermitIsOnHold = False 
		End If 
	Else
		GetPermitIsOnHold = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitIsVoided( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsVoided( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT isvoided FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isvoided") Then 
			GetPermitIsVoided = True 
		Else
			GetPermitIsVoided = False 
		End If 
	Else
		GetPermitIsVoided = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, bBreakOnCity )
'--------------------------------------------------------------------------------------------------
Function GetPermitLocation( ByVal iPermitId, ByRef sLegalDescription, ByRef sListedOwner, ByRef iPermitAddressId, ByRef sCounty, ByRef sParcelid, ByVal bBreakOnCity )
	Dim sSql, oRs

	sSql = "SELECT permitaddressid, residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & " residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, "
	sSql = sSql & " ISNULL(residentunit,'') AS residentunit, ISNULL(residentcity,'') AS residentcity, "
	sSql = sSql & " ISNULL(residentstate,'') AS residentstate, ISNULL(legaldescription,'') AS legaldescription, "
	sSql = sSQl & " ISNULL(listedowner,'') AS listedowner, ISNULL(residentzip,'') AS residentzip, ISNULL(county,'') AS county, "
	sSql = sSQl & " ISNULL(parcelidnumber,'') AS parcelid "
	sSql = sSQl & " FROM egov_permitaddress WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPermitLocation = oRs("residentstreetnumber")
		If oRs("residentstreetprefix") <> "" Then
			GetPermitLocation = GetPermitLocation & " " & oRs("residentstreetprefix")
		End If
		GetPermitLocation = GetPermitLocation & " " & oRs("residentstreetname")
		If oRs("streetsuffix") <> "" Then
			GetPermitLocation = GetPermitLocation & " " & oRs("streetsuffix")
		End If
		If oRs("streetdirection") <> "" Then
			GetPermitLocation = GetPermitLocation & " " & oRs("streetdirection")
		End If
		If oRs("residentunit") <> "" Then
			GetPermitLocation = GetPermitLocation & ", " & oRs("residentunit")
		End If
		If bBreakOnCity Then
			GetPermitLocation = GetPermitLocation & "<br />"
		End If 
		
		If oRs("residentcity") <> "" Then
			If bBreakOnCity Then
				GetPermitLocation = GetPermitLocation & oRs("residentcity")
			Else 
				GetPermitLocation = GetPermitLocation & ", " & oRs("residentcity")
			End If 
		End If 
		If oRs("residentstate") <> "" Then
			GetPermitLocation = GetPermitLocation & ", " & oRs("residentstate")
		End If 
		If oRs("residentzip") <> "" Then 
			GetPermitLocation = GetPermitLocation & " " & oRs("residentzip")
		End If 
		sLegalDescription = Trim(oRs("legaldescription"))
		sListedOwner = Trim(oRs("listedowner"))
		iPermitAddressId = oRs("permitaddressid")
		sCounty =  oRs("county")
		sParcelid =  oRs("parcelid")
	Else 
		GetPermitLocation = ""
		sLegalDescription = ""
		sListedOwner = ""
		sCounty = ""
		sParcelid = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitNumber( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitNumber( ByVal iPermitId )
	Dim sSql, oRs, sPermitNumberYear, sPermitNumberPrefix, sPermitNumber, sFormatedNumber

	sPermitNumberYear = ""
	sPermitNumberPrefix = ""
	sPermitNumber = "0"
	sFormatedNumber = ""

	sSql = "SELECT ISNULL(permitnumber,0) AS permitnumber, permitnumberyear, permitnumberprefix "
	sSql = sSql & "FROM egov_permits WHERE permitid = " & iPermitId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		'GetPermitNumber = oRs("permitnumberyear") & oRs("permitnumberprefix") & oRs("permitnumber")
		sPermitNumberYear = oRs("permitnumberyear")
		sPermitNumberPrefix = oRs("permitnumberprefix")
		sPermitNumber = Trim(oRs("permitnumber"))
	End If 

	oRs.CLose
	Set oRs = Nothing 

	If CLng(sPermitNumber) > CLng(0) Then 
		' Now get the permit number format
		sSql = "SELECT element, characters FROM egov_permitnumberformat "
		sSql = sSql & "WHERE isforbuildingpermits = 1 AND orgid = " & iOrgid & " ORDER BY position"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			Do While Not oRs.EOF 
				Select Case oRs("element")
					Case "year"
						sFormatedNumber = sFormatedNumber & Right(sPermitNumberYear, clng(oRs("characters")))
					Case "dash"
						sFormatedNumber = sFormatedNumber & "-"
					Case "prefix"
						If PermitNumberPrefixIsNotNone( sPermitNumberPrefix ) Then 
							sFormatedNumber = sFormatedNumber & sPermitNumberPrefix
						End If 
					Case "space"
						sFormatedNumber = sFormatedNumber & Space(clng(oRs("characters")))
					Case "sequence"
						If clng(Len(sPermitNumber)) < clng(oRs("characters")) Then 
							sFormatedNumber = sFormatedNumber & Replace(Space(clng(oRs("characters")) - Len(sPermitNumber))," ","0") & sPermitNumber
						Else
							sFormatedNumber = sFormatedNumber & sPermitNumber
						End If 
				End Select 
				oRs.MoveNext 
			Loop 
		End If 

		oRs.CLose
		Set oRs = Nothing 
	End If 

	GetPermitNumber = sFormatedNumber

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitPermitTypeFlag( iPermitid, sFlag )
'--------------------------------------------------------------------------------------------------
Function GetPermitPermitTypeFlag( ByVal iPermitid, ByVal sFlag )
	Dim sSql, oRs

	sSql = "SELECT " & sFlag & " AS isflagged FROM egov_permitpermittypes WHERE permitid = " & iPermitid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If oRs("isflagged") Then 
			GetPermitPermitTypeFlag = True 
		Else
			GetPermitPermitTypeFlag = False  
		End If 
	Else
		GetPermitPermitTypeFlag = False  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitPlansByContact( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitPlansByContact( ByVal iPermitId )
	Dim sSql, oRs, sDetails

	sDetails = ""
	sSql = "SELECT ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(C.company,'') AS company, ISNULL(C.address,'') AS address, ISNULL(C.city,'') AS city, "
	sSql = sSql & " ISNULL(C.state,'') AS state, ISNULL(C.zip,'') AS zip, ISNULL(C.phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts C, egov_permits P "
	sSql = sSql & " WHERE C.permitcontactid = P.plansbycontactid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("firstname") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("company") & "</strong><br />" 
		End If 
		If Trim(oRs("address")) <> "" Then 
			sDetails = sDetails & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sDetails = sDetails & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sDetails = sDetails & FormatPhoneNumber( oRs("phone") ) 
		End If 
	End If  

	oRs.Close
	Set oRs = Nothing 

	GetPermitPlansByContact = sDetails

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitReviewerName( iReviewerUserId )
'--------------------------------------------------------------------------------------------------
Function GetPermitReviewerName( ByVal iReviewerUserId )
	Dim sSql, oRs

	sSql = "SELECT firstname, lastname FROM users WHERE userid = " & iReviewerUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitReviewerName = oRs("firstname") & " " & oRs("lastname")
	Else
		GetPermitReviewerName = ""
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitStatusByStatusId( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusByStatusId( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT permitstatus FROM egov_permitstatuses WHERE permitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPermitStatusByStatusId = oRs("permitstatus")
	Else
		GetPermitStatusByStatusId = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitStatusId( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusId( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPermitStatusId = oRs("permitstatusid")
	Else
		GetPermitStatusId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitStatusIdByStatusType( sType )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusIdByStatusType( ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses WHERE orgid = " & iOrgid
	sSql = sSql & " AND isforbuildingpermits = 1 AND " & sType & " = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPermitStatusIdByStatusType = CLng(oRs("permitstatusid"))
	Else
		GetPermitStatusIdByStatusType = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitTypeDesc( iPermitId, bIncludePrefix )
'--------------------------------------------------------------------------------------------------
Function GetPermitTypeDesc( ByVal iPermitId, ByVal bIncludePrefix )
	Dim sSql, oRs, sType

	sType = ""
	sSql = "SELECT permittypedesc, permittype FROM egov_permitpermittypes "
	sSql = sSQl & " WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If bIncludePrefix Then 
			sType = oRs("permittype") & " &ndash; "
		End If 
		sType = sType & oRs("permittypedesc")
	Else 
		sType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitTypeDesc = sType

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitUseClass( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitUseClass( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT useclass FROM egov_permituseclasses U, egov_permits P "
	sSql = sSql & "WHERE P.useclassid = U.useclassid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPermitUseClass = oRs("useclass")
	Else 
		GetPermitUseClass = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitUseType( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitUseType( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT usetype FROM egov_permitusetypes U, egov_permits P "
	sSql = sSql & "WHERE P.usetypeid = U.usetypeid AND P.permitid = " & iPermitId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPermitUseType = oRs("usetype")
	Else 
		GetPermitUseType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitWorkClass( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitWorkClass( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT workclass FROM egov_permitworkclasses W, egov_permits P "
	sSql = sSql & "WHERE P.workclassid = W.workclassid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPermitWorkClass = oRs("workclass")
	Else 
		GetPermitWorkClass = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitWorkScope( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitWorkScope( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT workscope FROM egov_permitworkscope W, egov_permits P "
	sSql = sSql & "WHERE P.workscopeid = W.workscopeid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPermitWorkScope = oRs("workscope")
	Else 
		GetPermitWorkScope = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetWaivedTotal( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetWaivedTotal( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(totalamount),0.00) AS totalamount FROM egov_permitinvoices "
	sSql = sSql & " WHERE isvoided = 0 AND allfeeswaived = 1 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetWaivedTotal =  FormatNumber(oRS("totalamount"),2,,,0)
	Else
		GetWaivedTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitHasDetail( iPermitid, sDetailField )
'--------------------------------------------------------------------------------------------------
Function PermitHasDetail( ByVal iPermitid, ByVal sDetailField )
	Dim sSql, oRs

	sSql = "SELECT COUNT(F.detailfieldid) AS hits "
	sSql = sSql & " FROM egov_permits P, egov_permittypes_to_permitdetailfields F, egov_permitdetailfields D "
	sSql = sSql & " WHERE P.permitid = " & iPermitId & " AND D.detailfield = '" & sDetailField & "' AND "
	sSql = sSql & " P.permittypeid = F.permittypeid AND F.detailfieldid = D.detailfieldid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasDetail = True 
		Else 
			PermitHasDetail = False  
		End If 
	Else
		PermitHasDetail = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitNumberPrefixIsNotNone( sPermitNumberPrefix )
'--------------------------------------------------------------------------------------------------
Function PermitNumberPrefixIsNotNone( ByVal sPermitNumberPrefix )
	Dim sSql, oRs

	sSql = "SELECT isnone FROM egov_permitnumberprefixes WHERE orgid = " & session("orgid")
	sSql = sSql & " AND permitnumberprefix = '" & sPermitNumberPrefix & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isnone") Then 
			PermitNumberPrefixIsNotNone = False
		Else 
			PermitNumberPrefixIsNotNone = True
		End If 
	Else
		PermitNumberPrefixIsNotNone = True
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowAttachmentList iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowAttachmentList( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitattachmentid, attachmentname, ISNULL(description,'') AS description, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, dateadded "
	sSql = sSql & " FROM egov_permitattachments WHERE permitid = " & iPermitId
	sSql = sSql & " ORDER BY 1 DESC"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<tr><th class=""firstcell"">File Name</th><th>Description</th><th>Date Added</th></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center"" class=""firstcell"">" & oRs("attachmentname") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & oRs("description") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & DateValue(oRs("dateadded")) & "</td>"
			'response.write "<td align=""center"" class=""bordercell"">" & GetAdminName( oRs("adminuserid") ) & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
	Else 
		response.write vbcrlf & "<tr><td colspan=""3"">No Attachments</td></tr>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowInspectionList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowInspectionList( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT I.permitinspectionid, I.permitinspectiontype, I.inspectiondescription, I.isrequired, S.inspectionstatus, "
	sSql = sSql & " I.inspecteddate, I.isreinspection, ISNULL(I.inspectoruserid,0) AS inspectoruserid, isfinal "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY I.inspectionorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
		
			' Inspection
			response.write "<td class=""firstcell"">" & oRs("permitinspectiontype") & " &mdash; " & oRs("inspectiondescription") & "</td>"

			' Reinspection
			response.write "<td align=""center"" class=""bordercell"">"
			If oRs("isreinspection") Then
				response.write "Reinspection"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			' Status
			response.write "<td align=""center"" class=""bordercell"">" & oRs("inspectionstatus") & "</td>"

			' Date
			response.write "<td align=""center"" class=""bordercell"">" 
			If IsNull(oRs("inspecteddate")) Then
				response.write "&nbsp;"
			Else 
				response.write DateValue(oRs("inspecteddate"))
			End If 
			response.write "</td>"

			' Inspector
			response.write "<td align=""center"" class=""bordercell"">"
			If CLng(oRs("inspectoruserid")) > CLng(0) Then 
				response.write GetAdminName( CLng(oRs("inspectoruserid")) )
			Else
				response.write "Unassigned"
			End If 
			response.write "</td>"

			response.write "</tr>"

			ShowInspectionNotes oRs("permitinspectionid")

			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowInspectionList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowInspectionNotes iPermitInspectionId 
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionNotes( ByVal iPermitInspectionId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSql & " S.inspectionstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_inspectionstatuses S, users U "
	sSql = sSql & " WHERE S.inspectionstatusid = L.inspectionstatusid AND U.userid = L.adminuserid "
	sSql = sSql & " AND permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " AND L.isinspectionentry = 1 ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr><td class=""firstcell"" valign=""top"" align=""right"">Notes:</td><td class=""bordercell"" colspan=""4"">"
		Do While Not oRs.EOF 
			If oRs("externalcomment") <> "" Then 
				iRowCount = iRowCount + 1
				If CLng(iRowCount) > CLng(1) Then
					response.write "<hr />"
				End If 
				response.write oRs("firstname") & " " & oRs("lastname") & " &ndash; " & oRs("inspectionstatus") & " &ndash; " & DateValue(oRs("entrydate")) & "<br />"
			
				response.write oRs("externalcomment")
			End If 
			oRs.MoveNext
		Loop 

		If iRowCount = CLng(0) Then
			response.write "&nbsp;"
		End If 

		response.write "</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInvoicePayments iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoicePayments( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(L.amount,0.00) AS amount, P.paymenttypename, P.requirescheckno "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_paymenttypes P "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND J.paymentid = " & iPaymentId
	sSql = sSql & " AND L.entrytype = 'debit' AND L.paymenttypeid = P.paymenttypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		'response.write vbcrlf & "<br />"
		response.write "&nbsp;" & oRs("paymenttypename") 
		If oRs("requirescheckno") Then 
			response.write " #: " & GetCheckNo( iPaymentId )
		End If 
		response.write " for " & FormatCurrency(oRs("amount"),2)
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInvoices iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoices( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	sSql = "SELECT I.invoiceid, I.invoicedate, I.totalamount, ISNULL(I.paymentid,0) AS paymentid, I.permitcontactid, "
	sSql = sSql & " S.invoicestatus, I.allfeeswaived, S.isvoid FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.isvoided = 0 AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center"" class=""firstcell"">" & oRs("invoiceid") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & FormatDateTime(oRs("invoicedate"),2) & "</td>"
			
			response.write "<td align=""center"" class=""bordercell"">"
			response.write GetInvoiceContact( oRs("permitcontactid") )
			response.write "</td>"

			response.write "<td align=""center"" class=""bordercell"">" & oRs("invoicestatus") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & FormatNumber(oRs("totalamount"),2) & "</td>"
			response.write "<td align=""center"" class=""bordercell"">"
			If oRs("allfeeswaived") Then 
				response.write FormatNumber(oRs("totalamount"),2)
			Else
				response.write FormatNumber(GetInvoicePaymentTotal( CLng(oRs("invoiceid")) ),2)   ' in permitcommonfunctions.asp
			End If 
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPayments iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPayments( ByVal iPermitId )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT L.paymentid, J.paymentdate, ISNULL(SUM(L.amount),0.00) AS paymenttotal "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_permitinvoices I, egov_class_payment J "
	sSql = sSql & " WHERE I.isvoided = 0 AND L.invoiceid = I.invoiceid AND L.permitid = " & iPermitId
	sSql = sSql & " AND J.paymentid = L.paymentid AND L.ispaymentaccount = 0 "
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td align=""center"" valign=""top"" class=""firstcell"">" & oRs("paymentid") & "</td>"
		response.write "<td align=""center"" valign=""top"" class=""bordercell"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
		
		response.write "<td valign=""top"" class=""bordercell"">"
		' Show payment types and amount
		ShowInvoicePayments oRs("paymentid")

		response.write "</td>"
		response.write "<td align=""right"" valign=""top"" class=""bordercell"">" & FormatNumber(oRs("paymenttotal"),2) & " &nbsp;</td>"
		response.write "</tr>"
		
		dTotal = dTotal + CDbl(oRs("paymenttotal"))

		oRs.MoveNext
	Loop 
	response.write vbcrlf & "<tr><td colspan=""3""align=""right"" class=""firstcell""><strong>Total Payments</strong>&nbsp;</td>"
	response.write "<td align=""right"" class=""bordercell"">" & FormatNumber(dTotal,2) & " &nbsp;</td>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitFees iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFees( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT F.permitfeeid, F.isrequired, F.includefee, f.isfixturetypefee, "
	sSql = sSql & " ISNULL(F.permitfeeprefix,'') AS permitfeeprefix, F.permitfee, "
	sSql = sSql & " F.isvaluationtypefee, F.isconstructiontypefee, F.feeamount, "
	sSql = sSql & " ISNULL(F.paymentid,0) AS paymentid, M.permitfeemethod, M.isflatfee, "
	sSql = sSql & " M.ismanual, M.isfixture, F.isupfrontfee, M.ishourly "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitid =" & iPermitId
	sSql = sSql & " ORDER BY F.displayorder, F.permitfeeid"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"   
		response.write "<td nowrap=""nowrap"" align=""center"" class=""firstcell"">"  ' Category cell
		If oRs("permitfeeprefix") = "" Then
			response.write "&nbsp;"
		Else 
			response.write oRs("permitfeeprefix")
		End If 
		response.write "</td>"
		response.write "<td class=""bordercell"">"  ' Description cell
		response.write oRs("permitfee")
		response.write "</td>"
		response.write "<td align=""center"" class=""bordercell"">"  ' Method cell
		response.write oRs("permitfeemethod")
		response.write "</td>"
		response.write "<td align=""center"" class=""bordercell"">"  ' Fee amount cell
		response.write FormatNumber(oRs("feeamount"),2)
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitNotes iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitNotes( ByVal iPermitId )
	Dim sSql, oRs, iRowCount

	iRowCount = CLng(0)

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, "
	sSql = sSql & " ISNULL(externalcomment,'') AS externalcomment, S.permitstatus, ISNULL(L.adminuserid,0) AS adminuserid, "
	sSql = sSql & " ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_permitstatuses S "
	sSql = sSql & " WHERE isactivityentry = 1 AND S.permitstatusid = L.permitstatusid AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If oRs("externalcomment") <> "" Then
			iRowCount = iRowCount + CLng(1)
			response.write vbcrlf & "<tr>"
			response.write "<td class=""onlycell"">"
'			If CLng(oRs("adminuserid")) > CLng(0) then
'				response.write GetAdminName( CLng(oRs("adminuserid")) ) ' In common.asp
'			Else
'				response.write "System Generated"
'			End If 
'			response.write " &ndash; " & oRs("permitstatus") & " &ndash; " & oRs("entrydate") & "<br />"
			
			response.write DateValue(oRs("entrydate")) & " &nbsp "
			response.write oRs("externalcomment")
			response.write "</td></tr>"
		End If 
		
		oRs.MoveNext
	Loop 

	If iRowCount = CLng(0) Then 
		response.write vbcrlf & "<tr><td>No Notes</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowReviewList iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewList( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT R.permitreviewid, R.permitreviewtype, R.isrequired, R.isincluded, S.reviewstatus, "
	sSql = sSql & " ISNULL(R.revieweruserid,0) AS revieweruserid, R.reviewed "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE R.reviewstatusid = S.reviewstatusid AND R.permitid = " & iPermitId
	sSql = sSql & " ORDER BY R.revieworder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			' Review type
			response.write "<td class=""firstcell"">" & oRs("permitreviewtype") & "</td>"
			' Status
			response.write "<td align=""center"" class=""bordercell"">" & oRs("reviewstatus") & "</td>"
			' Reviewed
			response.write "<td align=""center"" class=""bordercell"">"
			If IsNull(oRs("reviewed")) Then
				response.write "&nbsp;"
			Else
				response.write DateValue(oRs("reviewed"))
			End If 
			response.write "</td>"

			' Reviewer
			response.write "<td align=""center"" class=""bordercell"">"
			If CLng(oRs("revieweruserid")) > CLng(0) Then 
				response.write GetPermitReviewerName( CLng(oRs("revieweruserid")) )
			Else
				response.write "Unassigned"
			End If 
			response.write "</td>"

			response.write "</tr>"

			ShowReviewNotes oRs("permitreviewid") 

			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReviewNotes iPermitReviewId 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewNotes( ByVal iPermitReviewId )
	Dim sSql, oRs, iRowCount

	iRowCount = CLng(0)

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSql & " S.reviewstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_reviewstatuses S, users U "
	sSql = sSql & " WHERE S.reviewstatusid = L.reviewstatusid AND U.userid = L.adminuserid AND permitreviewid = " & iPermitReviewId
	sSql = sSql & " AND L.isreviewentry = 1 ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr><td class=""firstcell"" valign=""top"" align=""right"">Notes:</td><td class=""bordercell"" colspan=""3"">"
		Do While Not oRs.EOF 
			If oRs("externalcomment") <> "" Then
				iRowCount = iRowCount + CLng(1)
				If CLng(iRowCount) > CLng(1) Then
					response.write "<hr />"
				End If 
				response.write oRs("firstname") & " " & oRs("lastname") & " &ndash; " & oRs("reviewstatus") & " &ndash; " & DateValue(oRs("entrydate")) & "<br />"
			 
				response.write oRs("externalcomment") 
			End If 
			oRs.MoveNext
		Loop 

		If iRowCount = CLng(0) Then
			response.write "&nbsp;"
		End If 

		response.write "</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>

