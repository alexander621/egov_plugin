<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: newpermitcreate.asp
' AUTHOR: Steve Loar
' CREATED: 02/27/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates permits
'
' MODIFICATION HISTORY
' 1.0   02/27/2008	Steve Loar - INITIAL VERSION
' 1.1	03/27/2008	Steve Loar - Added county
' 1.2	04/02/2008	Steve Loar - removed landvalue, totalvalue, taxdistrict, added streetdirection
' 1.3	07/25/2008	Steve Loar - Inspectors of unassigned added
' 1.4	07/21/2009	Steve Loar - New fields for Lansing IL
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitStatusId, iApplicantUserId, iAdminUserId, sStreetName, sAddress, sStreetNumber, iPermitTypeId
Dim iPermitAddressTypeId, iPermitId, sSql, sPermitPrefix, sPermitNoYear, sExpirationDate, sDisplayAddress
Dim iFeeTotal, sContactType, iUseTypeId, sHasTempCo, sHasCo, sPermitLocation, iPermitCategoryId
Dim iPermitLocationRequirementId, sLocationRequired

' For applied date so it is in local time
' ", dbo.GetLocalDate(" & Session("OrgID") & ",getdate())"

iFeeTotal = CDbl(0.00)

iPermitStatusId = GetInitialPermitStatusId()
'response.write "iPermitStatusId=" & iPermitStatusId & "<br />"

sContactType = UCase(Left(request("userid"),1))	' either U=user or C=contactor
iApplicantUserId = CLng(Mid(request("userid"),2))
'response.write "iApplicantUserId=" & iApplicantUserId & "<br />"
iAdminUserId = CLng(session("userid"))
'response.write "iAdminUserId=" & iAdminUserId & "<br />"

iPermitAddressTypeId = request("residentaddressid")

iPermitTypeId = CLng(request("permittypeid"))
'response.write "iPermitTypeId=" & iPermitTypeId & "<br />"
sPermitPrefix = GetPermitPrefix( iPermitTypeId ) 
sPermitNoYear = "'" & CStr(Year(Date())) & "'"
sExpirationDate = GetExpirationDate( iPermitTypeId )
iUseTypeId = GetPermitTypeUseType( iPermitTypeId )

' Get the categoryid
iPermitCategoryId = GetPermitTypeCategoryId( iPermitTypeId )

' Get the permitlocationrequirementid
iPermitLocationRequirementId = GetPermitTypeLocationRequirementId( iPermitTypeId )

sLocationRequired = request("locationrequired")
response.write "sLocationRequired = " & sLocationRequired & "<br /><br />"

If request("location") <> "" Then 
	If sLocationRequired = "location" Then 
		sPermitLocation = "'" & dbsafe(request("location")) & "'"
	Else
		sPermitLocation = "NULL"
	End If 
Else
	sPermitLocation = "NULL"
End If 

' Create the Permit row
sSql = "INSERT INTO egov_permits ( orgid, permitnumberprefix, permitnumberyear, applicantuserid, "
sSql = sSql & "adminuserid, permitstatusid, applieddate, expirationdate, permittypeid, lastactivitydate, "
sSql = sSql & "isbuildingpermit, usetypeid, permitlocation, permitcategoryid, permitlocationrequirementid "
sSql = sSql & " ) VALUES ( "
sSql = sSql & session("orgid") & ", " & sPermitPrefix & ", " & sPermitNoYear & ", " & iApplicantUserId & ", "
sSql = sSql & iAdminUserId & ", " & iPermitStatusId & ", dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), "
sSql = sSql & sExpirationDate & ", " & iPermitTypeId & ", dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), 1, "
sSql = sSql & iUseTypeId & ", " & sPermitLocation & ", " & iPermitCategoryId & ", " & iPermitLocationRequirementId 
sSql = sSql & " )"
response.write sSql & "<br /><br />"
'response.End 

iPermitId = RunIdentityInsert( sSql )
'response.write "iPermitId=" & iPermitId & "<br />"

' Get the permit Type info and create an entry for that
CreatePermitPermitType iPermitId, iPermitTypeId

' Get the applicant info and create a row for that
If sContactType = "U" Then 
	CreateApplicantFromUser iPermitId, iApplicantUserId
Else
	CreateApplicantFromContractor iPermitId, iApplicantUserId
End If 

' Bring in the Fees and Fixtures
CreatePermitFees iPermitId, iPermitTypeId, iFeeTotal

If CDbl(iFeeTotal) > CDbl(0.00) Then
	UpdateFeeTotal iPermitId, iFeeTotal
End If 

' Bring in any Reviews
CreatePermitReviews iPermitId, iPermitTypeId

' Bring in any Inspections
CreatePermitInspections iPermitId, iPermitTypeId

' Bring in any Required Licenses
CreatePermitRequiredLicenses iPermitId, iPermitTypeId

' Bring in the Address Info
CreatePermitAddress iPermitId, iPermitAddressTypeId, sLocationRequired

' Copy any custom fields that this permit type needs
CreatePermitCustomFields iPermitId, iPermitTypeId

' Log the permit creation
MakeAPermitLogEntry iPermitId, "'New Permit Request Created'", "'New Permit Request Created'", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

' Take them to the edit screen
response.redirect "permitedit.asp?permitid=" & iPermitId


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' integer GetInitialPermitStatusId()
'-------------------------------------------------------------------------------------------------
Function GetInitialPermitStatusId()
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses WHERE isinitialstatus = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInitialPermitStatusId = oRs("permitstatusid")
	Else
		GetInitialPermitStatusId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetAddressId( sStreetNumber, sStreetName )
'-------------------------------------------------------------------------------------------------
Function GetAddressId( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs

	sSql = "SELECT residentaddressid FROM egov_residentaddresses WHERE residentstreetnumber = " & sStreetNumber
	sSql = sSql & " and residentstreetname = '" & dbsafe(sStreetName) & "' and orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetAddressId = oRs("residentaddressid")
	Else
		GetAddressId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitPrefix( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitPrefix( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitnumberprefix,'B') AS permitnumberprefix FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitPrefix = "'" & oRs("permitnumberprefix") & "'"
	Else
		GetPermitPrefix = "''"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' date GetExpirationDate( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetExpirationDate( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(expirationdays,365) AS expirationdays FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetExpirationDate = "'" & DateAdd("d",CLng(oRs("expirationdays")), Date()) & "'"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' void CreatePermitPermitType iPermitId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CreatePermitPermitType( ByVal iPermitId, ByVal iPermitTypeId )
	Dim sSql, oRs, iIsBuildingPermitType, iIsAutoApproved, sExpirationDays, sPublicDescription
	Dim sPermitSubTitle, sPermitRightTitle, sPermitTitleBottom, sPermitFooter, sListFixtures
	Dim sShowConstructionType, sShowFeeTotal, sShowOccupancyType, sShowJobValue, sShowWorkDesc
	Dim sShowFootages, sShowProposedUse, sAdditionalFooterInfo, sPermitTitle, sApprovingOfficial
	Dim sPermitLogo, sGroupByInvoiceCategories, sInvoiceLogo, sInvoiceHeader, sShowElectricalContractor
	Dim sShowMechanicalContractor, sShowPlumbingContractor, sShowApplicantLicense, iPermitCategoryId
	Dim sShowCounty, sShowParcelid, sShowPlansBy, sShowPrimaryContact, sHasTempCo, sHasCo, sShowApprovedAsOnTCO, sShowApprovedAsOnCO
	Dim sShowConstTypeOnTCO, sShowConstTypeOnCO, sShowOccTypeonTCO, sShowOccTypeonCO, sShowOccupantsOnTCO, sShowOccupantsOnCO
	Dim sTempCOLogo, sCoLogo, sTempCOTitle, sTempCOSubTitle, sCOTitle, sCOSubTitle, sTempCOAddress, sCOAddress
	Dim sTempCOTopText, sCOTopText, sTempCOBottomText, sCOBottomText, sTempCOCodeRef, sCOCodeRef
	Dim sTempCOApproval, sCOApproval, sTempCOFooter, sCOFooter, sTempCOSubFooter, sCOSubFooter
	Dim sShowTotalSqFt, sShowApprovedAs, sShowFeeTypeTotals, sShowOccupancyUse, sShowPayments, iDocumentId
	Dim iPermitLocationRequirementId, sAttachmentReviewerAlert

	sSql = "SELECT ISNULL(permittype, '') AS permittype, ISNULL(permittypedesc, '') AS permittypedesc, permitcategoryid, "
	sSql = sSql & " expirationdays, isautoapproved, ISNULL(permitnumberprefix, 'B') AS permitnumberprefix, publicdescription, "
	sSql = sSql & " permittitle, additionalfooterinfo, approvingofficial, permitsubtitle, permitrighttitle, "
	sSql = sSql & " permittitlebottom, permitfooter, permitsubfooter, listfixtures, showconstructiontype, showfeetotal, "
	sSql = sSql & " showoccupancytype, showjobvalue, showworkdesc, showfootages, showproposeduse, ISNULL(permitlogo,'') AS permitlogo, "
	sSql = sSql & " groupbyinvoicecategories, ISNULL(invoicelogo,'') AS invoicelogo, ISNULL(invoiceheader,'') AS invoiceheader, "
	sSql = sSql & " showelectricalcontractor, showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, "
	sSql = sSql & " showcounty, showparcelid, showplansby, showprimarycontact, ISNULL(usetypeid,0) AS usetypeid, hastempco, hasco, showapprovedasontco, "
	sSql = sSql & " showconsttypeontco, showocctypeontco, showoccupantsontco, showapprovedasonco, showconsttypeonco, showocctypeonco, "
	sSql = sSql & " showoccupantsonco, ISNULL(tempcologo,'') AS tempcologo, ISNULL(cologo,'') AS cologo, ISNULL(tempcotitle,'') AS tempcotitle, "
	sSql = sSql & " ISNULL(tempcosubtitle,'') AS tempcosubtitle, ISNULL(cotitle,'') AS cotitle, ISNULL(cosubtitle,'') AS cosubtitle, "
	sSql = sSql & " ISNULL(tempcoaddress,'') AS tempcoaddress, ISNULL(coaddress,'') AS coaddress, ISNULL(tempcotoptext,'') AS tempcotoptext, "
	sSql = sSql & " ISNULL(cotoptext,'') AS cotoptext, ISNULL(tempcobottomtext,'') AS tempcobottomtext, ISNULL(cobottomtext,'') AS cobottomtext, "
	sSql = sSql & " ISNULL(tempcocoderef,'') AS tempcocoderef, ISNULL(cocoderef,'') AS cocoderef, ISNULL(tempcoapproval,'') AS tempcoapproval, "
	sSql = sSql & " ISNULL(coapproval,'') AS coapproval, ISNULL(tempcofooter,'') AS tempcofooter, ISNULL(tempcosubfooter,'') AS tempcosubfooter, "
	sSql = sSql & " ISNULL(cofooter,'') AS cofooter, ISNULL(cosubfooter,'') AS cosubfooter, showtotalsqft, showapprovedas, showfeetypetotals, "
	sSql = sSql & " showoccupancyuse, showpayments, documentid, permitlocationrequirementid, attachmentrevieweralert "
	sSql = sSql & " FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then

		iPermitCategoryId = oRs("permitcategoryid")

		iPermitLocationRequirementId = oRs("permitlocationrequirementid")
		
		If oRs("showoccupancyuse") Then
			sShowOccupancyUse = 1
		Else
			sShowOccupancyUse = 0
		End If
'		If oRs("isbuildingpermittype") Then
'			iIsBuildingPermitType = 1
'		Else
'			iIsBuildingPermitType = 0
'		End If 
		If oRs("isautoapproved") Then
			iIsAutoApproved = 1
		Else
			iIsAutoApproved = 0
		End If 
		If IsNull(oRs("expirationdays")) Then
			sExpirationDays = "NULL"
		Else
			sExpirationDays = oRs("expirationdays")
		End If 
		If IsNull(oRs("publicdescription")) Then
			sPublicDescription = "NULL"
		Else
			sPublicDescription = "'" & dbsafe(oRs("publicdescription")) & "'"
		End If
		If IsNull(oRs("permittitle")) Then
			sPermitTitle = "NULL"
		Else
			sPermitTitle = "'" & DBsafeWithHTML(oRs("permittitle")) & "'"
		End If
		If IsNull(oRs("additionalfooterinfo")) Then
			sAdditionalFooterInfo = "NULL"
		Else
			sAdditionalFooterInfo = "'" & DBsafeWithHTML(oRs("additionalfooterinfo")) & "'"
		End If
		If IsNull(oRs("approvingofficial")) Then
			sApprovingOfficial = "NULL"
		Else
			sApprovingOfficial = "'" & DBsafeWithHTML(oRs("approvingofficial")) & "'"
		End If
		If IsNull(oRs("permitsubtitle")) Then
			sPermitSubTitle = "NULL"
		Else
			sPermitSubTitle = "'" & DBsafeWithHTML(oRs("permitsubtitle")) & "'"
		End If
		If IsNull(oRs("permitrighttitle")) Then
			sPermitRightTitle = "NULL"
		Else
			sPermitRightTitle = "'" & DBsafeWithHTML(oRs("permitrighttitle")) & "'"
		End If
		If IsNull(oRs("permittitlebottom")) Then
			sPermitTitleBottom = "NULL"
		Else
			sPermitTitleBottom = "'" & dbsafe(oRs("permittitlebottom")) & "'"
		End If
		If IsNull(oRs("permitfooter")) Then
			sPermitFooter = "NULL"
		Else
			sPermitFooter = "'" & DBsafeWithHTML(oRs("permitfooter")) & "'"
		End If
		If IsNull(oRs("permitsubfooter")) Then
			sPermitSubFooter = "NULL"
		Else
			sPermitSubFooter = "'" & DBsafeWithHTML(oRs("permitsubfooter")) & "'"
		End If

		If oRs("listfixtures") Then
			sListFixtures = 1
		Else
			sListFixtures = 0
		End If 
		If oRs("showconstructiontype") Then
			sShowConstructionType = 1
		Else
			sShowConstructionType = 0
		End If 
		If oRs("showfeetotal") Then
			sShowFeeTotal = 1
		Else
			sShowFeeTotal = 0
		End If 
		If oRs("showoccupancytype") Then
			sShowOccupancyType = 1
		Else
			sShowOccupancyType = 0
		End If 
		If oRs("showjobvalue") Then
			sShowJobValue = 1
		Else
			sShowJobValue = 0
		End If 
		If oRs("showworkdesc") Then
			sShowWorkDesc = 1
		Else
			sShowWorkDesc = 0
		End If 
		If oRs("showfootages") Then
			sShowFootages = 1
		Else
			sShowFootages = 0
		End If 
		If oRs("showproposeduse") Then
			sShowProposedUse = 1
		Else
			sShowProposedUse = 0
		End If 
		If oRs("permitlogo") = "" Then
			sPermitLogo = "NULL"
		Else
			sPermitLogo =  "'" & DBsafeWithHTML(oRs("permitlogo")) & "'"
		End If
		If oRs("groupbyinvoicecategories") = "on" Then 
			sGroupByInvoiceCategories = 1
		Else 
			sGroupByInvoiceCategories = 0
		End If 
		If oRs("invoicelogo") = "" Then
			sInvoiceLogo = "NULL"
		Else
			sInvoiceLogo =  "'" & DBsafeWithHTML(oRs("invoicelogo")) & "'"
		End If
		If oRs("invoiceheader") = "" Then
			sInvoiceHeader = "NULL"
		Else
			sInvoiceHeader =  "'" & DBsafeWithHTML(oRs("invoiceheader")) & "'"
		End If
		If oRs("showelectricalcontractor") Then 
			sShowElectricalContractor = 1
		Else 
			sShowElectricalContractor = 0
		End If 
		If oRs("showmechanicalcontractor") Then 
			sShowMechanicalContractor = 1
		Else 
			sShowMechanicalContractor = 0
		End If 
		If oRs("showplumbingcontractor") Then 
			sShowPlumbingContractor = 1
		Else 
			sShowPlumbingContractor = 0
		End If 
		If oRs("showapplicantlicense") Then 
			sShowApplicantLicense = 1
		Else 
			sShowApplicantLicense = 0
		End If 
		If oRs("showcounty") Then 
			sShowCounty = 1
		Else 
			sShowCounty = 0
		End If
		If oRs("showparcelid") Then 
			sShowParcelid = 1
		Else 
			sShowParcelid = 0
		End If
		If oRs("showplansby") Then 
			sShowPlansBy = 1
		Else 
			sShowPlansBy = 0
		End If
		
		If oRs("showprimarycontact") Then 
			sShowPrimaryContact = 1
		Else 
			sShowPrimaryContact = 0
		End If

		iUseTypeId = CLng(oRs("usetypeid"))

		If oRs("hastempco") Then 
			sHasTempCo = 1
		Else 
			sHasTempCo = 0
		End If 

		If oRs("hasco") Then 
			sHasCo = 1
		Else 
			sHasCo = 0
		End If 

		If oRs("showapprovedasontco") Then 
			sShowApprovedAsOnTCO = 1
		Else 
			sShowApprovedAsOnTCO = 0
		End If 

		If oRs("showconsttypeontco") Then 
			sShowConstTypeOnTCO = 1
		Else 
			sShowConstTypeOnTCO = 0
		End If 

		If oRs("showocctypeontco") Then 
			sShowOccTypeonTCO = 1
		Else 
			sShowOccTypeonTCO = 0
		End If 

		If oRs("showoccupantsontco") Then 
			sShowOccupantsOnTCO = 1
		Else 
			sShowOccupantsOnTCO = 0
		End If 

		If oRs("showapprovedasonco") Then 
			sShowApprovedAsOnCO = 1
		Else 
			sShowApprovedAsOnCO = 0
		End If 

		If oRs("showconsttypeonco") Then 
			sShowConstTypeOnCO = 1
		Else 
			sShowConstTypeOnCO = 0
		End If 

		If oRs("showocctypeonco") Then 
			sShowOccTypeonCO = 1
		Else 
			sShowOccTypeonCO = 0
		End If 

		If oRs("showoccupantsonco") Then 
			sShowOccupantsOnCO = 1
		Else 
			sShowOccupantsOnCO = 0
		End If 

		If oRs("tempcologo") = "" Then
			sTempCOLogo = "NULL"
		Else
			sTempCOLogo =  "'" & DBsafeWithHTML(oRs("tempcologo")) & "'"
		End If

		If oRs("cologo") = "" Then
			sCOLogo = "NULL"
		Else
			sCOLogo =  "'" & DBsafeWithHTML(oRs("cologo")) & "'"
		End If
		If oRs("tempcotitle") = "" Then
			sTempCOTitle = "NULL"
		Else
			sTempCOTitle =  "'" & DBsafeWithHTML(oRs("tempcotitle")) & "'"
		End If
		If oRs("tempcosubtitle") = "" Then
			sTempCOSubTitle = "NULL"
		Else
			sTempCOSubTitle =  "'" & DBsafeWithHTML(oRs("tempcosubtitle")) & "'"
		End If
		If oRs("cotitle") = "" Then
			sCOTitle = "NULL"
		Else
			sCOTitle =  "'" & DBsafeWithHTML(oRs("cotitle")) & "'"
		End If
		If oRs("cosubtitle") = "" Then
			sCOSubTitle = "NULL"
		Else
			sCOSubTitle =  "'" & DBsafeWithHTML(oRs("cosubtitle")) & "'"
		End If
		If oRs("tempcoaddress") = "" Then
			sTempCOAddress = "NULL"
		Else
			sTempCOAddress =  "'" & DBsafeWithHTML(oRs("tempcoaddress")) & "'"
		End If
		If oRs("coaddress") = "" Then
			sCOAddress = "NULL"
		Else
			sCOAddress =  "'" & DBsafeWithHTML(oRs("coaddress")) & "'"
		End If
		If oRs("tempcotoptext") = "" Then
			sTempCOTopText = "NULL"
		Else
			sTempCOTopText =  "'" & DBsafeWithHTML(oRs("tempcotoptext")) & "'"
		End If
		If oRs("cotoptext") = "" Then
			sCOTopText = "NULL"
		Else
			sCOTopText =  "'" & DBsafeWithHTML(oRs("cotoptext")) & "'"
		End If
		If oRs("tempcobottomtext") = "" Then
			sTempCOBottomText = "NULL"
		Else
			sTempCOBottomText =  "'" & DBsafeWithHTML(oRs("tempcobottomtext")) & "'"
		End If
		If oRs("cobottomtext") = "" Then
			sCOBottomText = "NULL"
		Else
			sCOBottomText =  "'" & DBsafeWithHTML(oRs("cobottomtext")) & "'"
		End If
		If oRs("tempcocoderef") = "" Then
			sTempCOCodeRef = "NULL"
		Else
			sTempCOCodeRef =  "'" & DBsafeWithHTML(oRs("tempcocoderef")) & "'"
		End If
		If oRs("cocoderef") = "" Then
			sCOCodeRef = "NULL"
		Else
			sCOCodeRef =  "'" & DBsafeWithHTML(oRs("cocoderef")) & "'"
		End If
		If oRs("tempcoapproval") = "" Then
			sTempCOApproval = "NULL"
		Else
			sTempCOApproval =  "'" & DBsafeWithHTML(oRs("tempcoapproval")) & "'"
		End If
		If oRs("coapproval") = "" Then
			sCOApproval = "NULL"
		Else
			sCOApproval =  "'" & DBsafeWithHTML(oRs("coapproval")) & "'"
		End If
		If oRs("tempcofooter") = "" Then
			sTempCOFooter = "NULL"
		Else
			sTempCOFooter =  "'" & DBsafeWithHTML(oRs("tempcofooter")) & "'"
		End If
		If oRs("tempcosubfooter") = "" Then
			sTempCOSubFooter = "NULL"
		Else
			sTempCOSubFooter =  "'" & DBsafeWithHTML(oRs("tempcosubfooter")) & "'"
		End If
		If oRs("cofooter") = "" Then
			sCOFooter = "NULL"
		Else
			sCOFooter =  "'" & DBsafeWithHTML(oRs("cofooter")) & "'"
		End If
		If oRs("cosubfooter") = "" Then
			sCOSubFooter = "NULL"
		Else
			sCOSubFooter =  "'" & DBsafeWithHTML(oRs("cosubfooter")) & "'"
		End If
		If oRs("showtotalsqft") Then
			sShowTotalSqFt = 1
		Else
			sShowTotalSqFt = 0
		End If 
		If oRs("showapprovedas") Then
			sShowApprovedAs = 1
		Else
			sShowApprovedAs = 0
		End If 
		If oRs("showfeetypetotals") Then
			sShowFeeTypeTotals = 1
		Else
			sShowFeeTypeTotals = 0
		End If
		If oRs("showoccupancyuse") Then
			sShowOccupancyUse = 1
		Else
			sShowOccupancyUse = 0
		End If
		If oRs("showpayments") Then
			sShowPayments = 1
		Else
			sShowPayments = 0
		End If
		iDocumentId = oRs("documentid")
		If oRs("attachmentrevieweralert") Then
			sAttachmentReviewerAlert = 1
		Else
			sAttachmentReviewerAlert = 0
		End If 

		sSql = "INSERT INTO egov_permitpermittypes ( permitid, permittypeid, orgid, permittype, permittypedesc, "
		sSql = sSql & " permitcategoryid, expirationdays, isautoapproved, displayorder, permitnumberprefix, "
		sSql = sSql & " publicdescription, permittitle, additionalfooterinfo, "
		sSql = sSql & " approvingofficial, permitsubtitle, permitrighttitle, permittitlebottom, permitfooter, "
		sSql = sSql & " permitsubfooter, listfixtures, showconstructiontype, showfeetotal, "
		sSql = sSql & " showoccupancytype, showjobvalue, showworkdesc, showfootages, showproposeduse, permitlogo, "
		sSql = sSql & " groupbyinvoicecategories, invoicelogo, invoiceheader, showelectricalcontractor, "
		sSql = sSql & " showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, showcounty, showparcelid, "
		sSql = sSql & " showplansby, showprimarycontact, hastempco, hasco, showapprovedasontco, showconsttypeontco, "
		sSql = sSql & " showocctypeontco, showoccupantsontco, showapprovedasonco, showconsttypeonco, showocctypeonco, "
		sSql = sSql & " showoccupantsonco, tempcologo, cologo, tempcotitle, tempcosubtitle, cotitle, cosubtitle, tempcoaddress, "
		sSql = sSql & " coaddress, tempcotoptext, cotoptext, tempcobottomtext, cobottomtext, tempcocoderef, cocoderef, "
		sSql = sSql & " tempcoapproval, coapproval, tempcofooter, tempcosubfooter, cofooter, cosubfooter, showtotalsqft, "
		sSql = sSql & " showapprovedas, showfeetypetotals, showoccupancyuse, showpayments, documentid, permitlocationrequirementid, "
		sSql = sSql & " attachmentrevieweralert ) VALUES ( "
		sSql = sSql & iPermitId & ", " & iPermitTypeId & ", " & session("orgid") & ", '" & dbsafe(oRs("permittype")) & "', '"
		sSql = sSql & dbsafe(oRs("permittypedesc")) & "', " & iPermitCategoryId & ", " & sExpirationDays & ", " & iIsAutoApproved
		sSql = sSql & ", 1, '" & oRs("permitnumberprefix") & "', " & sPublicDescription & ", " & sPermitTitle 
		sSql = sSql & ", " & sAdditionalFooterInfo & ", " & sApprovingOfficial & ", " & sPermitSubTitle & ", " & sPermitRightTitle
		sSql = sSql & ", " & sPermitTitleBottom & ", " & sPermitFooter & ", " & sPermitSubFooter & ", " & sListFixtures
		sSql = sSql & ", " & sShowConstructionType & ", " & sShowFeeTotal & ", " & sShowOccupancyType & ", " & sShowJobValue
		sSql = sSql & ", " & sShowWorkDesc & ", " & sShowFootages & ", " & sShowProposedUse & ", " & sPermitLogo 
		sSql = sSql & ", " & sGroupByInvoiceCategories & ", " & sInvoiceLogo & ", " & sInvoiceHeader & ", " & sShowElectricalContractor 
		sSql = sSql & ", " & sShowMechanicalContractor & ", " & sShowPlumbingContractor & ", " & sShowApplicantLicense
		sSql = sSql & ", " & sShowCounty & ", " & sShowParcelid & ", " & sShowPlansBy & ", " & sShowPrimaryContact
		sSql = sSql & ", " & sHasTempCo & ", " & sHasCO & ", " & sShowApprovedAsOnTCO & ", " & sShowConstTypeOnTCO & ", " & sShowOccTypeonTCO
		sSql = sSql & ", " & sShowOccupantsOnTCO & ", " & sShowApprovedAsOnCO & ", " & sShowConstTypeOnCO & ", " & sShowOccTypeonCO
		sSql = sSql & ", " & sShowOccupantsOnCO & ", " & sTempCOLogo & ", " & sCOLogo & ", " & sTempCOTitle & ", " & sTempCOSubTitle
		sSql = sSql & ", " & sCOTitle & ", " & sCOSubTitle & ", " & sTempCOAddress & ", " & sCOAddress & ", " & sTempCOTopText
		sSql = sSql & ", " & sCOTopText & ", " & sTempCOBottomText & ", " & sCOBottomText & ", " & sTempCOCodeRef & ", " & sCOCodeRef
		sSql = sSql & ", " & sTempCOApproval & ", " & sCOApproval & ", " & sTempCOFooter & ", " & sTempCOSubFooter & ", " & sCOFooter
		sSql = sSql & ", " & sCOSubFooter & ", " & sShowTotalSqFt & ", " & sShowApprovedAs & ", " & sShowFeeTypeTotals
		sSql = sSql & ", " & sShowOccupancyUse & ", " & sShowPayments & ", " & iDocumentId & ", " & iPermitLocationRequirementId
		sSql = sSql & ", " & sAttachmentReviewerAlert & " )"

		RunSQL sSql 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreateApplicantFromUser iPermitId, iApplicantUserId 
'-------------------------------------------------------------------------------------------------
Sub CreateApplicantFromUser( ByVal iPermitId, ByVal iApplicantUserId )
	Dim sSql, oRs, sUserbusinessname, sUserfname, sUserlname, sUseraddress, sUserCity, sUserstate
	Dim sUserzip, sUseremail, sUserhomephone, sUsercell, sUserfax

	sSql = "SELECT userbusinessname, userfname, userlname, useraddress, usercity, userstate, userzip, "
	sSql = sSql & " useremail, userhomephone, usercell, userfax, userpassword, userworkphone, emergencycontact, "
	sSql = sSql & " emergencyphone, neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable, "
	sSql = sSql & " residencyverified FROM egov_users WHERE userid = " & iApplicantUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("userbusinessname")) Then
			sUserbusinessname = "NULL"
		Else
			sUserbusinessname = "'" & dbsafe(oRs("userbusinessname")) & "'"
		End If 
		If IsNull(oRs("userfname")) Then
			sUserfname = "NULL"
		Else
			sUserfname = "'" & dbsafe(oRs("userfname")) & "'"
		End If 
		If IsNull(oRs("userlname")) Then
			sUserlname = "NULL"
		Else
			sUserlname = "'" & dbsafe(oRs("userlname")) & "'"
		End If 
		If IsNull(oRs("useraddress")) Then
			sUseraddress = "NULL"
		Else
			sUseraddress = "'" & dbsafe(oRs("useraddress")) & "'"
		End If 
		If IsNull(oRs("usercity")) Then
			sUserCity = "NULL"
		Else
			sUserCity = "'" & dbsafe(oRs("usercity")) & "'"
		End If
		If IsNull(oRs("userstate")) Then
			sUserstate = "NULL"
		Else
			sUserstate = "'" & dbsafe(oRs("userstate")) & "'"
		End If
		If IsNull(oRs("userzip")) Then
			sUserzip = "NULL"
		Else
			sUserzip = "'" & dbsafe(oRs("userzip")) & "'"
		End If
		If IsNull(oRs("useremail")) Then
			sUseremail = "NULL"
		Else
			sUseremail = "'" & dbsafe(oRs("useremail")) & "'"
		End If
		If IsNull(oRs("userhomephone")) Then
			sUserhomephone = "NULL"
		Else
			sUserhomephone = "'" & oRs("userhomephone") & "'"
		End If
		If IsNull(oRs("usercell")) Then
			sUsercell = "NULL"
		Else
			sUsercell = "'" & oRs("usercell") & "'"
		End If
		If IsNull(oRs("userfax")) Then
			sUserfax = "NULL"
		Else
			sUserfax = "'" & oRs("userfax") & "'"
		End If
		sPassword = "'" & oRs("userpassword") & "'"
		sWorkPhone = "'" & oRs("userworkphone") & "'"
		sEmergencyContact = "'" & oRs("emergencycontact") & "'"
		sEmergencyPhone = "'" & oRs("emergencyphone") & "'"
		If IsNull(oRs("neighborhoodid")) Then
			iNeighborhoodid = "NULL"
		else
			iNeighborhoodid = oRs("neighborhoodid")
		End If 
		If IsNull(oRs("residenttype")) Or oRs("residenttype") = "" Then
			sResidentType = "'R'"
		Else 
			sResidentType = "'" & oRs("residenttype") & "'"
		End If 
		sBusinessAddress = "'" & dbsafe(oRs("userbusinessaddress")) & "'"
		sUserUnit = "'" & dbsafe(oRs("userunit")) & "'"
		If oRs("emailnotavailable") Then 
			sEmailnotavailable = 1
		Else
			sEmailnotavailable = 0
		End If 
		If oRs("residencyverified") Then 
			sResidencyVerified = 1
		Else
			sResidencyVerified = 0
		End If 

		sSql = "INSERT INTO egov_permitcontacts ( permitid, permitcontacttypeid, orgid, company, firstname, "
		sSql = sSql & " lastname, address, city, state, zip, email, phone, cell, fax, contacttype, "
		sSql = sSql & " isapplicant, userid, userpassword, userworkphone, emergencycontact, emergencyphone, "
		sSql = sSql & " neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable, residencyverified ) "
		sSql = sSql & " VALUES ( " & iPermitId & ", NULL, " & session("orgid") & ", "
		sSql = sSql & sUserbusinessname & ", " & sUserfname & ", " & sUserlname & ", " & sUseraddress & ", "
		sSql = sSql & sUserCity & ", " & sUserstate & ", " & sUserzip & ", " & sUseremail & ", " & sUserhomephone
		sSql = sSql & ", " & sUsercell & ", " & sUserfax & ", 'U', 1, " & iApplicantUserId
		sSql = sSql & ", " & sPassword & ", " & sWorkPhone & ", " & sEmergencyContact & ", "
		sSql = sSql & sEmergencyPhone & ", " & iNeighborhoodId & ", " & sResidentType & ", "
		sSql = sSql & sBusinessAddress & ", " & sUserUnit & ", " & sEmailnotavailable & ", " & sResidencyVerified & " )"

		RunSQL sSql 
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'-------------------------------------------------------------------------------------------------
' void CreateApplicantFromContractor iPermitId, iContactTypeId 
'-------------------------------------------------------------------------------------------------
Sub CreateApplicantFromContractor( ByVal iPermitId, ByVal iContactTypeId )
	Dim sSql, oRs, sUserbusinessname, sUserfname, sUserlname, sUseraddress, sUserCity, sUserstate
	Dim sUserzip, sUseremail, sUserhomephone, sUsercell, sUserfax, iContractorTypeId, iIsOrganization
	Dim iBusinessTypeId, sStateLicense, sEmployeeCount, sReference1, sAutoInsurancePhone, sBondAgent
	Dim sReference2, sReference3, sOtherLicensedCity1, sOtherLicensedCity2, sGeneralLiabilityAgent
	Dim sGeneralLiabilityPhone, sWorkersCompAgent, sWorkersCompPhone, sAutoInsuranceAgent, sBondAgentPhone

	sSql = "SELECT permitcontacttypeid, company, firstname, lastname, address, city, state, zip, isorganization, "
	sSql = sSql & " email, phone, cell, fax, userid, ISNULL(contractortypeid,0) AS contractortypeid, "
	sSql = sSql & " businesstypeid, statelicense, employeecount, reference1, reference2, reference3, "
	sSql = sSql & " otherlicensedcity1, otherlicensedcity2, generalliabilityagent, generalliabilityphone, "
	sSql = sSql & " workerscompagent, workerscompphone, autoinsuranceagent, autoinsurancephone, "
	sSql = sSql & " bondagent, bondagentphone "
	sSql = sSql & " FROM egov_permitcontacttypes WHERE permitcontacttypeid = " & iContactTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("company")) Then
			sCompany = "NULL"
		Else
			sCompany = "'" & dbsafe(oRs("company")) & "'"
		End If 
		If IsNull(oRs("firstname")) Then
			sFirstname = "NULL"
		Else
			sFirstname = "'" & dbsafe(oRs("firstname")) & "'"
		End If 
		If IsNull(oRs("lastname")) Then
			sLastname = "NULL"
		Else
			sLastname = "'" & dbsafe(oRs("lastname")) & "'"
		End If 
		If IsNull(oRs("address")) Then
			sAddress = "NULL"
		Else
			sAddress = "'" & dbsafe(oRs("address")) & "'"
		End If 
		If IsNull(oRs("city")) Then
			sCity = "NULL"
		Else
			sCity = "'" & dbsafe(oRs("city")) & "'"
		End If
		If IsNull(oRs("state")) Then
			sState = "NULL"
		Else
			sState = "'" & dbsafe(oRs("state")) & "'"
		End If
		If IsNull(oRs("zip")) Then
			sZip = "NULL"
		Else
			sZip = "'" & dbsafe(oRs("zip")) & "'"
		End If
		If IsNull(oRs("email")) Then
			sEmail = "NULL"
		Else
			sEmail = "'" & dbsafe(oRs("email")) & "'"
		End If
		If IsNull(oRs("phone")) Then
			sPhone = "NULL"
		Else
			sPhone = "'" & oRs("phone") & "'"
		End If
		If IsNull(oRs("cell")) Then
			sCell = "NULL"
		Else
			sCell = "'" & oRs("cell") & "'"
		End If
		If IsNull(oRs("fax")) Then
			sFax = "NULL"
		Else
			sFax = "'" & oRs("fax") & "'"
		End If
		If IsNull(oRs("userid")) Then
			iContactUserId = "NULL"
		Else
			iContactUserId = oRs("userid") 
		End If
		If CLng(oRs("contractortypeid")) > CLng(0) Then 
			iContractorTypeId = CLng(oRs("contractortypeid"))
		Else
			iContractorTypeId = "NULL"
		End If 
		If oRs("isorganization") Then 
			iIsOrganization = 1
		Else
			iIsOrganization = 0
		End If 
		If IsNull(oRs("businesstypeid")) Then
			iBusinessTypeId = "NULL"
		Else
			If CLng(oRs("businesstypeid")) > CLng(0) Then 
				iBusinessTypeId = CLng(oRs("businesstypeid"))
			Else 
				iBusinessTypeId = "NULL"
			End If 
		End If 
		If IsNull(oRs("statelicense")) Then
			sStateLicense = "NULL"
		Else
			sStateLicense = "'" & dbsafe(oRs("statelicense")) & "'"
		End If 
		If IsNull(oRs("employeecount")) Then
			sEmployeeCount = "NULL"
		Else
			sEmployeeCount = "'" & dbsafe(oRs("employeecount")) & "'"
		End If 
		If IsNull(oRs("reference1")) Then
			sReference1 = "NULL"
		Else
			sReference1 = "'" & dbsafe(oRs("reference1")) & "'"
		End If 
		If IsNull(oRs("reference2")) Then
			sReference2 = "NULL"
		Else
			sReference2 = "'" & dbsafe(oRs("reference2")) & "'"
		End If 
		If IsNull(oRs("reference3")) Then
			sReference3 = "NULL"
		Else
			sReference3 = "'" & dbsafe(oRs("reference3")) & "'"
		End If 
		If IsNull(oRs("otherlicensedcity1")) Then
			sOtherLicensedCity1 = "NULL"
		Else
			sOtherLicensedCity1 = "'" & dbsafe(oRs("otherlicensedcity1")) & "'"
		End If 
		If IsNull(oRs("otherlicensedcity2")) Then
			sOtherLicensedCity2 = "NULL"
		Else
			sOtherLicensedCity2 = "'" & dbsafe(oRs("otherlicensedcity2")) & "'"
		End If 
		If IsNull(oRs("generalliabilityagent")) Then
			sGeneralLiabilityAgent = "NULL"
		Else
			sGeneralLiabilityAgent = "'" & dbsafe(oRs("generalliabilityagent")) & "'"
		End If 
		If IsNull(oRs("generalliabilityphone")) Then
			sGeneralLiabilityPhone = "NULL"
		Else
			sGeneralLiabilityPhone = "'" & dbsafe(oRs("generalliabilityphone")) & "'"
		End If 
		If IsNull(oRs("workerscompagent")) Then
			sWorkersCompAgent = "NULL"
		Else
			sWorkersCompAgent = "'" & dbsafe(oRs("workerscompagent")) & "'"
		End If 
		If IsNull(oRs("workerscompphone")) Then
			sWorkersCompPhone = "NULL"
		Else
			sWorkersCompPhone = "'" & dbsafe(oRs("workerscompphone")) & "'"
		End If 
		If IsNull(oRs("autoinsuranceagent")) Then
			sAutoInsuranceAgent = "NULL"
		Else
			sAutoInsuranceAgent = "'" & dbsafe(oRs("autoinsuranceagent")) & "'"
		End If 
		If IsNull(oRs("autoinsurancephone")) Then
			sAutoInsurancePhone = "NULL"
		Else
			sAutoInsurancePhone = "'" & dbsafe(oRs("autoinsurancephone")) & "'"
		End If 
		If IsNull(oRs("bondagent")) Then
			sBondAgent = "NULL"
		Else
			sBondAgent = "'" & dbsafe(oRs("bondagent")) & "'"
		End If 
		If IsNull(oRs("bondagentphone")) Then
			sBondAgentPhone = "NULL"
		Else
			sBondAgentPhone = "'" & dbsafe(oRs("bondagentphone")) & "'"
		End If 

		sSql = "INSERT INTO egov_permitcontacts ( permitid, permitcontacttypeid, orgid, company, firstname, "
		sSql = sSql & " lastname, address, city, state, zip, email, phone, cell, fax, contacttype, "
		sSql = sSql & " isapplicant, userid, userworkphone, contractortypeid, isorganization, businesstypeid, "
		sSql = sSql & " statelicense, employeecount, reference1, reference2, reference3, otherlicensedcity1, "
		sSql = sSql & " otherlicensedcity2, generalliabilityagent, generalliabilityphone, workerscompagent, "
		sSql = sSql & " workerscompphone, autoinsuranceagent, autoinsurancephone, bondagent, bondagentphone )"
		sSql = sSql & " VALUES ( " & iPermitId & ", " & iContactTypeId & ", " & session("orgid") & ", "
		sSql = sSql & sCompany & ", " & sFirstname & ", " & sLastname & ", " & sAddress & ", "
		sSql = sSql & sCity & ", " & sState & ", " & sZip & ", " & sEmail & ", " & sPhone & ", "
		sSql = sSql & sCell & ", " & sFax & ", 'C', 1, NULL, " & sPhone & ", " & iContractorTypeId & ", "
		sSql = sSql & iIsOrganization & ", " & iBusinessTypeId & ", " & sStateLicense & ", " & sEmployeeCount & ", "
		sSql = sSql & sReference1 & ", " & sReference2 & ", " & sReference3 & ", " & sOtherLicensedCity1 & ", "
		sSql = sSql & sOtherLicensedCity2 & ", " & sGeneralLiabilityAgent & ", " & sGeneralLiabilityPhone & ", "
		sSql = sSql & sWorkersCompAgent & ", " & sWorkersCompPhone & ", " & sAutoInsuranceAgent & ", "
		sSql = sSql & sAutoInsurancePhone & ", " & sBondAgent & ", " & sBondAgentPhone & " )"

		iContactId = RunIdentityInsert( sSql )

		If clng(iIsOrganization) = clng(0) Then 
			' Get their licenses and insert them
			CreateNewLicenseRecords iContactTypeId, iContactId, iPermitId
		End If 

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitFees iPermitId, iPermitTypeId, iFeeTotal
'-------------------------------------------------------------------------------------------------
Sub CreatePermitFees( ByVal iPermitId, ByVal iPermitTypeId, ByRef iFeeTotal )
	Dim sSql, oRs, iIsfixturetypefee, iAtleastqty, iNotmorethanqty, iBaseamount, iUnitqty, iUnitamount
	Dim iMinimumamount, iIsupfrontfee, iIsreinspectionfee, iIsbuildingpermitfee, iIsrequired, iFeeCount
	Dim iPermitFeeId, iAccountid, iFeeAmount, iIsvaluationtypefee, iIsconstructiontypefee, bOnBBSFeeReport
	Dim iOnSewerFeeReport, iIsResidentialUnitTypeFee, iFeeReportingTypeId

	iFeeCount = 0
	iFeeAmount = CDbl(0.00) 

	sSql = "SELECT F.permitfeetypeid, F.isfixturetypefee, F.isvaluationtypefee, F.isconstructiontypefee, ISNULL(F.upfrontamount,0.00) AS upfrontamount, "
	sSql = sSql & " ISNULL(F.permitfeeprefix, '') AS permitfeeprefix, ISNULL(F.permitfee, '') AS permitfee, isresidentialunittypefee, "
	sSql = sSql & " F.permitfeecategorytypeid, F.permitfeemethodid, F.atleastqty, F.notmorethanqty, F.onsewerfeereport, F.onbbsfeereport, "
	sSql = sSql & " ISNULL(F.baseamount,0.00) AS baseamount, F.unitqty, F.unitamount, F.minimumamount, F.ispercentagetypefee, F.percentage, "
	sSql = sSql & " F.isupfrontfee, F.isreinspectionfee, F.isbuildingpermitfee, F.accountid, T.isrequired, T.displayorder, M.isflatfee, "
	sSql = sSql & " ISNULL(feereportingtypeid,0) AS feereportingtypeid "
	sSql = sSql & " FROM egov_permitfeetypes F, egov_permittypes_to_permitfeetypes T, egov_permitfeemethods M "
	sSql = sSql & " WHERE T.permitfeetypeid = F.permitfeetypeid AND F.permitfeemethodid = M.permitfeemethodid AND T.permittypeid = " & iPermitTypeId 
	sSql = sSql & " ORDER BY T.displayorder, F.permitfeetypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iFeeCount = iFeeCount + 1
		If oRs("isfixturetypefee") Then
			iIsfixturetypefee = 1
		Else
			iIsfixturetypefee = 0
		End If 
		If oRs("isvaluationtypefee") Then
			iIsvaluationtypefee = 1
		Else
			iIsvaluationtypefee = 0
		End If 
		If oRs("isconstructiontypefee") Then
			iIsconstructiontypefee = 1
		Else
			iIsconstructiontypefee = 0
		End If 
		If IsNull(oRs("atleastqty")) Then
			iAtleastqty = "NULL"
		Else
			iAtleastqty = oRs("atleastqty")
		End If 
		If IsNull(oRs("notmorethanqty")) Then
			iNotmorethanqty = "NULL"
		Else
			iNotmorethanqty = oRs("notmorethanqty")
		End If 
		If IsNull(oRs("baseamount")) Then
			iBaseamount = "NULL"
		Else
			iBaseamount = oRs("baseamount")
		End If 
		If IsNull(oRs("unitqty")) Then
			iUnitqty = "NULL"
		Else
			iUnitqty = oRs("unitqty")
		End If 
		If IsNull(oRs("unitamount")) Then
			iUnitamount = "NULL"
		Else
			iUnitamount = oRs("unitamount")
		End If 
		If IsNull(oRs("minimumamount")) Then
			iMinimumamount = "NULL"
		Else
			iMinimumamount = oRs("minimumamount")
		End If
		If oRs("isupfrontfee") Then
			iIsupfrontfee = 1
		Else
			iIsupfrontfee = 0
		End If 
		If oRs("isreinspectionfee") Then
			iIsreinspectionfee = 1
		Else
			iIsreinspectionfee = 0
		End If 
		If oRs("isbuildingpermitfee") Then
			iIsbuildingpermitfee = 1
		Else
			iIsbuildingpermitfee = 0
		End If 
		If IsNull(oRs("accountid")) Then
			iAccountid = "NULL"
		Else
			iAccountid = oRs("accountid")
		End If
		If oRs("isrequired") Then
			iIsrequired = 1
		Else
			iIsrequired = 0
		End If 

		If oRs("isflatfee") Then
			If oRs("isrequired") Then
				' sum the flatfees for the permit total amount
				iFeeTotal = iFeeTotal + CDbl(oRs("baseamount"))
			End If 
			iFeeAmount = CDbl(oRs("baseamount"))
		Else
			iFeeAmount = CDbl(0.00)
		End If 

		If oRs("ispercentagetypefee") Then
			iIspercentagetypefee = 1
		Else
			iIspercentagetypefee = 0
		End If 

		If IsNull(oRs("percentage")) Then
			sPercentage =  "NULL"
		Else 
			sPercentage = CDbl(oRs("percentage"))
		End If 

		If oRs("isresidentialunittypefee") Then
			iIsResidentialUnitTypeFee = 1
		Else
			iIsResidentialUnitTypeFee = 0
		End If 

		If CLng(oRs("feereportingtypeid")) <> CLng(0) Then 
			iFeeReportingTypeId = CLng(oRs("feereportingtypeid"))
		Else
			iFeeReportingTypeId = "NULL"
		End If 

		sSql = "INSERT INTO egov_permitfees ( permitid, permitfeetypeid, orgid, isfixturetypefee, isvaluationtypefee, isconstructiontypefee, permitfeeprefix, permitfee, "
		sSql = sSql & " permitfeecategorytypeid, permitfeemethodid, atleastqty, notmorethanqty, baseamount, quantity, unitqty, "
		sSql = sSql & " unitamount, minimumamount, isupfrontfee, isreinspectionfee, isbuildingpermitfee, accountid, isrequired, "
		sSql = sSql & " displayorder, feeamount, amountpaid, includefee, ispercentagetypefee, percentage, "
		sSql = sSql & " upfrontamount, feereportingtypeid, isresidentialunittypefee ) VALUES ( "
		sSql = sSql & iPermitId & ", " & oRs("permitfeetypeid") & ", " & session("orgid") & ", " & iIsfixturetypefee & ", "
		sSql = sSql & iIsvaluationtypefee & ", " & iIsconstructiontypefee & ", '" & oRs("permitfeeprefix") & "', '" & oRs("permitfee") & "', "
		sSql = sSql & oRs("permitfeecategorytypeid") & ", " & oRs("permitfeemethodid") & ", " & iAtleastqty & ", " & iNotmorethanqty & ", "
		sSql = sSql & iBaseamount & ", 0, " & iUnitqty & ", " & iUnitamount & ", " & iMinimumamount & ", " & iIsupfrontfee & ", "
		sSql = sSql & iIsreinspectionfee & ", " & iIsbuildingpermitfee & ", " & iAccountid & ", " & iIsrequired & ", "
		sSql = sSql & iFeeCount & ", " & iFeeAmount & ", 0.00, 1, " & iIspercentagetypefee & ", " & sPercentage & ", "
		sSql = sSql & oRs("upfrontamount") & ", " & iFeeReportingTypeId & ", " & iIsResidentialUnitTypeFee & " )"
		iPermitFeeId = RunIdentityInsert( sSql )

		If oRs("isfixturetypefee") Then
			' Get the fixtures and bring them over
			'response.write "<p>Fixtures Here</p>"
			CreatePermitFixtures iPermitId, oRs("permitfeetypeid"), iPermitFeeId
		End If 

		If oRs("isvaluationtypefee") Then
			' Get the valuations and bring them over
			CreatePermitValuationStepFees iPermitId, oRs("permitfeetypeid"), iPermitFeeId
		End If 

		If oRs("isresidentialunittypefee") Then
			CreatePermitResidentialUnitStepFees iPermitId, oRs("permitfeetypeid"), iPermitFeeId
		End If 

		' Pull in any Fee Multipliers
		CreatePermitFeeMultipliers oRs("permitfeetypeid"), iPermitFeeId, iPermitId

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitReviews iPermitId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CreatePermitReviews( ByVal iPermitId, ByVal iPermitTypeId )
	Dim sSql, oRs, iRowCount, iIsRequired, iIsIncluded, iInitialStatusid, iNotifyOnRelease

	iInitialStatusid = GetReviewStatusId( "isinitialstatus" )	' in permitcommonfunctions.asp
	iRowCount = 0

	sSql = "SELECT F.permitreviewtypeid, F.permitreviewtype, F.reviewdescription, T.revieweruserid, T.isrequired, T.notifyonrelease "
	sSql = sSql & " FROM egov_permitreviewtypes F, egov_permittypes_to_permitreviewtypes T "
	sSql = sSql & " WHERE F.permitreviewtypeid = T.permitreviewtypeid AND T.permittypeid = " & iPermitTypeId
	sSql = sSql & " ORDER BY T.revieworder, F.permitreviewtypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iRowCount = iRowCount + 1
		If oRs("isrequired") Then
			iIsRequired = 1
			iIsIncluded = 1
		Else
			iIsRequired = 0
			iIsIncluded = 0
		End If 
		If oRs("notifyonrelease") Then 
			iNotifyOnRelease = 1
		Else
			iNotifyOnRelease = 0
		End If 
		sSql = "INSERT INTO egov_permitreviews (orgid, permitid, permittypeid, permitreviewtypeid, permitreviewtype,  "
		sSql = sSql & " reviewdescription, revieweruserid, revieworder, isrequired, isincluded, reviewstatusid, notifyonrelease ) VALUES ( " & session("orgid")
		sSql = sSql & ", " & iPermitId & ", " & iPermitTypeId & ", " & oRs("permitreviewtypeid") &  ", '" & dbsafe(oRs("permitreviewtype"))
		sSql = sSql & "', '" & dbsafe(oRs("reviewdescription")) & "', " & oRs("revieweruserid") & ", " & iRowCount
		sSql = sSql & ", " & iIsRequired & ", " & iIsIncluded & ", " & iInitialStatusid & ", " & iNotifyOnRelease & " )"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitInspections iPermitId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CreatePermitInspections( ByVal iPermitId, ByVal iPermitTypeId )
	Dim sSql, oRs, iRowCount, iIsRequired, iIsIncluded, iIsFinal, iInitialStatusid, iPermitInspectorId

	iRowCount = 0
	iInitialStatusid = GetInspectionStatusId( "isinitialstatus" )	' in permitcommonfunctions.asp

	sSql = "SELECT F.permitinspectiontypeid, F.permitinspectiontype, F.inspectiondescription, ISNULL(T.inspectoruserid,0) AS inspectoruserid, T.isrequired, T.isfinal "
	sSql = sSql & " FROM egov_permitinspectiontypes F, egov_permittypes_to_permitinspectiontypes T "
	sSql = sSql & " WHERE F.permitinspectiontypeid = T.permitinspectiontypeid AND T.permittypeid = " & iPermitTypeId
	sSql = sSql & " ORDER BY T.inspectionorder, F.permitinspectiontypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iRowCount = iRowCount + 1
		If oRs("isrequired") Then
			iIsRequired = 1
			iIsIncluded = 1
		Else
			iIsRequired = 0
			iIsIncluded = 1		' Always want them included. They can remove them if they do not want them.
		End If 
		If oRs("isfinal") Then
			iIsFinal = 1
		Else
			iIsFinal = 0
		End If 
		If CLng(oRs("inspectoruserid")) = CLng(0) Then
			iPermitInspectorId = "NULL"
		Else
			iPermitInspectorId = oRs("inspectoruserid")
		End If 

		sSql = "INSERT INTO egov_permitinspections (orgid, permitid, permittypeid, permitinspectiontypeid, permitinspectiontype,  "
		sSql = sSql & " inspectiondescription, inspectoruserid, inspectionorder, isrequired, isfinal, isincluded, inspectionstatusid, routeorder ) VALUES ( " & session("orgid")
		sSql = sSql & ", " & iPermitId & ", " & iPermitTypeId & ", " & oRs("permitinspectiontypeid") &  ", '" & dbsafe(oRs("permitinspectiontype"))
		sSql = sSql & "', '" & dbsafe(oRs("inspectiondescription")) & "', " & iPermitInspectorId & ", " & iRowCount
		sSql = sSql & ", " & iIsRequired & ", " & iIsFinal & ", " & iIsIncluded & ", " & iInitialStatusid & ", 999 )"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitRequiredLicenses iPermitId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CreatePermitRequiredLicenses( ByVal iPermitId, ByVal iPermitTypeId )
	Dim sSql, oRs, sLicenseType, iDisplayOrder, iIsRequired

	sSql = "SELECT P.licensetypeid, L.licensetype, L.displayorder, P.isrequired "
	sSql = sSql & " FROM egov_permitlicensetypes L, egov_permittypes_to_permitlicensetypes P "
	sSql = sSql & " WHERE L.licensetypeid = P.licensetypeid AND P.permittypeid = " & iPermitTypeId
	sSql = sSql & " ORDER BY L.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If oRs("isrequired") Then
			iIsRequired = 1
		Else
			iIsRequired = 0
		End If 
		sSql = "INSERT INTO egov_permits_to_permitlicensetypes ( permitid, permittypeid, licensetypeid, orgid, licensetype, displayorder, isrequired ) VALUES ( " 
		sSql = sSql & iPermitId & ", " & iPermitTypeid & ", " & oRs("licensetypeid") & ", " & session("orgid") & ", '" & oRs("licensetype") & "', " & oRs("displayorder") & ", " & iIsRequired & " )"
		RunSQL sSql
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitAddress iPermitId, iPermitAddressTypeId, sLocationRequired
'-------------------------------------------------------------------------------------------------
Sub CreatePermitAddress( ByVal iPermitId, ByVal iPermitAddressTypeId, ByVal sLocationRequired )
	Dim sSql, oRs, sResidentstreetnumber, sResidentunit, sResidentstreetprefix, sSortstreetname
	Dim sResidentstreetname, sResidentcity, sResidentstate, sResidentzip, sResidenttype, sParcelidnumber
	Dim sLegaldescription, sLatitude, sLongitude, sStreetDirection, sListedowner, sRegisteredUserId
	Dim sStreetsuffix, sPropertyTaxNumber, sLotNumber, sLotWidth, sLotLength, sBlockNumber, sSubdivision
	Dim sSection, sTownship, sRange, sPermanentRealEstateIndexNumber, sCollectorsTaxBillVolumeNumber

	If sLocationRequired = "address" Then 
		sSql = "SELECT residentaddressid, residentstreetnumber, residentunit, residentstreetprefix, "
		sSql = sSql & " residentstreetname, sortstreetname, residentcity, residentstate, residentzip, county, "
		sSql = sSql & " residenttype, latitude, longitude, parcelidnumber, legaldescription, "
		sSql = sSql & " streetdirection, listedowner, registereduserid, streetsuffix, propertytaxnumber, "
		sSql = sSql & " lotnumber, lotwidth, lotlength, blocknumber, subdivision, section, township, "
		sSql = sSql & " range, permanentrealestateindexnumber, collectorstaxbillvolumenumber "
		sSql = sSql & " FROM egov_residentaddresses WHERE residentaddressid = " & iPermitAddressTypeId
		
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			If IsNull(oRs("residentstreetnumber")) Then
				sResidentstreetnumber = "NULL"
			Else
				sResidentstreetnumber = "'" & dbsafe(oRs("residentstreetnumber")) & "'"
			End If 
			If IsNull(oRs("residentunit")) Then
				sResidentunit = "NULL"
			Else
				sResidentunit = "'" & dbsafe(oRs("residentunit")) & "'"
			End If
			If IsNull(oRs("residentstreetprefix")) Then
				sResidentstreetprefix = "NULL"
			Else
				sResidentstreetprefix = "'" & dbsafe(oRs("residentstreetprefix")) & "'"
			End If
			If IsNull(oRs("residentstreetname")) Then
				sResidentstreetname = "NULL"
			Else
				sResidentstreetname = "'" & dbsafe(oRs("residentstreetname")) & "'"
			End If 
			If IsNull(oRs("sortstreetname")) Then
				sSortstreetname = "NULL"
			Else
				sSortstreetname = "'" & dbsafe(oRs("sortstreetname")) & "'"
			End If
			If IsNull(oRs("residentcity")) Then
				sResidentcity = "NULL"
			Else
				sResidentcity = "'" & dbsafe(oRs("residentcity")) & "'"
			End If
			If IsNull(oRs("residentstate")) Then
				sResidentstate = "NULL"
			Else
				sResidentstate = "'" & dbsafe(oRs("residentstate")) & "'"
			End If
			If IsNull(oRs("residentzip")) Then
				sResidentzip = "NULL"
			Else
				sResidentzip = "'" & dbsafe(oRs("residentzip")) & "'"
			End If
			If IsNull(oRs("county")) Then
				sCounty = "NULL"
			Else
				sCounty = "'" & dbsafe(oRs("county")) & "'"
			End If
			If IsNull(oRs("residenttype")) Then
				sResidenttype = "NULL"
			Else
				sResidenttype = "'" & dbsafe(oRs("residenttype")) & "'"
			End If
			If IsNull(oRs("latitude")) Then
				sLatitude = "NULL"
			Else
				sLatitude = oRs("latitude")
			End If
			If IsNull(oRs("longitude")) Then
				sLongitude = "NULL"
			Else
				sLongitude = oRs("longitude")
			End If
			If IsNull(oRs("parcelidnumber")) Then
				sParcelidnumber = "NULL"
			Else
				sParcelidnumber = "'" & dbsafe(oRs("parcelidnumber")) & "'"
			End If
			If IsNull(oRs("legaldescription")) Then
				sLegaldescription = "NULL"
			Else
				sLegaldescription = "'" & dbsafe(oRs("legaldescription")) & "'"
			End If
			If IsNull(oRs("listedowner")) Then
				sListedowner = "NULL"
			Else
				sListedowner = "'" & dbsafe(oRs("listedowner")) & "'"
			End If
			If IsNull(oRs("registereduserid")) Then
				sRegisteredUserId = "NULL"
			Else
				sRegisteredUserId = oRs("registereduserid")
			End If
			If IsNull(oRs("streetsuffix")) Then
				sStreetsuffix = "NULL"
			Else
				sStreetsuffix = "'" & dbsafe(oRs("streetsuffix")) & "'"
			End If
			If IsNull(oRs("streetdirection")) Then
				sStreetDirection = "NULL"
			Else
				sStreetDirection = "'" & dbsafe(oRs("streetdirection")) & "'"
			End If

			If IsNull(oRs("propertytaxnumber")) Then
				sPropertyTaxNumber = "NULL"
			Else
				sPropertyTaxNumber = "'" & dbsafe(oRs("propertytaxnumber")) & "'"
			End If 

			If IsNull(oRs("lotnumber")) Then
				sLotNumber = "NULL"
			Else
				sLotNumber = "'" & dbsafe(oRs("lotnumber")) & "'"
			End If 

			If IsNull(oRs("lotwidth")) Then
				sLotWidth = "NULL"
			Else
				sLotWidth = "'" & dbsafe(oRs("lotwidth")) & "'"
			End If 

			If IsNull(oRs("lotlength")) Then
				sLotLength = "NULL"
			Else
				sLotLength = "'" & dbsafe(oRs("lotlength")) & "'"
			End If 

			If IsNull(oRs("blocknumber")) Then
				sBlockNumber = "NULL"
			Else
				sBlockNumber = "'" & dbsafe(oRs("blocknumber")) & "'"
			End If 

			If IsNull(oRs("subdivision")) Then
				sSubdivision = "NULL"
			Else
				sSubdivision = "'" & dbsafe(oRs("subdivision")) & "'"
			End If 

			If IsNull(oRs("section")) Then
				sSection = "NULL"
			Else
				sSection = "'" & dbsafe(oRs("section")) & "'"
			End If 

			If IsNull(oRs("township")) Then
				sTownship = "NULL"
			Else
				sTownship = "'" & dbsafe(oRs("township")) & "'"
			End If 

			If IsNull(oRs("range")) Then
				sRange = "NULL"
			Else
				sRange = "'" & dbsafe(oRs("range")) & "'"
			End If 

			If IsNull(oRs("permanentrealestateindexnumber")) Then
				sPermanentRealEstateIndexNumber = "NULL"
			Else
				sPermanentRealEstateIndexNumber = "'" & dbsafe(oRs("permanentrealestateindexnumber")) & "'"
			End If 

			If IsNull(oRs("collectorstaxbillvolumenumber")) Then
				sCollectorsTaxBillVolumeNumber = "NULL"
			Else
				sCollectorsTaxBillVolumeNumber = "'" & dbsafe(oRs("collectorstaxbillvolumenumber")) & "'"
			End If 

			sSql = "INSERT INTO egov_permitaddress (permitid, residentaddressid, residentstreetnumber, residentunit, "
			sSql = sSql & " residentstreetprefix, residentstreetname, sortstreetname, residentcity, residentstate, "
			sSql = sSql & " residentzip, county, orgid, residenttype, latitude, longitude, parcelidnumber, legaldescription, "
			sSql = sSql & " listedowner, registereduserid, streetsuffix, streetdirection, propertytaxnumber, lotnumber, "
			sSql = sSql & " lotwidth, lotlength, blocknumber, subdivision, section, township, range, "
			sSql = sSql & " permanentrealestateindexnumber, collectorstaxbillvolumenumber ) VALUES ( "
			sSql = sSql & iPermitId & ", " & oRs("residentaddressid") & ", " & sResidentstreetnumber & ", "
			sSql = sSql & sResidentunit & ", " & sResidentstreetprefix & ", " & sResidentstreetname & ", " 
			sSql = sSql & sSortstreetname & ", " & sResidentcity & ", " & sResidentstate & ", "
			sSql = sSql & sResidentzip & ", " & sCounty & ", " & session("orgid") & ", " & sResidenttype & ", " & sLatitude & ", "
			sSql = sSql & sLongitude & ", " & sParcelidnumber & ", " & sLegaldescription & ", " & sListedowner & ", "
			sSql = sSql & sRegisteredUserId & ", " & sStreetsuffix & ", " & sStreetDirection & ", "
			sSql = sSql & sPropertyTaxNumber & ", " & sLotNumber & ", " & sLotWidth & ", " & sLotLength & ", "
			sSql = sSql & sBlockNumber & ", " & sSubdivision & ", " & sSection & ", " & sTownship & ", " & sRange & ", "
			sSql = sSql & sPermanentRealEstateIndexNumber & ", " & sCollectorsTaxBillVolumeNumber & " )"
			RunSQL sSql
		End If 

		oRs.Close
		Set oRs = Nothing 
	Else
		' Put in a blank row for sLocationRequired = location or none
		sSql = "INSERT INTO egov_permitaddress ( permitid, orgid ) VALUES ( "
		sSql = sSql & iPermitId & ", " & session("orgid") & " )"
		RunSQL sSql
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' integer GetPermitTypeCategoryId( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitTypeCategoryId( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitcategoryid FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTypeCategoryId = CLng(oRs("permitcategoryid"))
	Else 
		GetPermitTypeCategoryId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitTypeLocationRequirementId( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitTypeLocationRequirementId( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitlocationrequirementid FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTypeLocationRequirementId = CLng(oRs("permitlocationrequirementid"))
	Else 
		GetPermitTypeLocationRequirementId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void CreatePermitCustomFields iPermitId, iPermitTypeId
'-------------------------------------------------------------------------------------------------
Sub CreatePermitCustomFields( ByVal iPermitId, ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT F.customfieldtypeid, F.orgid, F.fieldtypeid, F.fieldname, F.pdffieldname, F.prompt, "
	sSql = sSql & "ISNULL(F.valuelist,'') AS valuelist, ISNULL(F.fieldsize,0) AS fieldsize, P.customfieldorder "
	sSql = sSql & "FROM egov_permitcustomfieldtypes F, egov_permittypes_to_permitcustomfieldtypes P "
	sSql = sSql & "WHERE F.customfieldtypeid = P.customfieldtypeid AND P.permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		' insert into the permit custom fields table
		sSql = "INSERT INTO egov_permitcustomfields ( permitid, orgid, fieldtypeid, fieldname, pdffieldname, "
		sSql = sSql & "prompt, valuelist, fieldsize, displayorder, customfieldtypeid ) VALUES ( " & iPermitId & ", "
		sSql = sSql & oRs("orgid") & ", " & oRs("fieldtypeid") & ", '" & dbsafe(oRs("fieldname")) & "', '"
		sSql = sSql & dbsafe(oRs("pdffieldname")) & "', '" & dbsafe(oRs("prompt")) & "', '"
		sSql = sSql & dbsafe(oRs("valuelist")) & "', " & oRs("fieldsize") & ", " & oRs("customfieldorder") & ", "
		sSql = sSql & oRs("customfieldtypeid") & " )"

		RunSQL sSql

		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
