<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypecopy.asp
' AUTHOR: Steve Loar
' CREATED: 10/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This copies permit types
'
' MODIFICATION HISTORY
' 1.0   10/10/2008	Steve Loar - INITIAL VERSION
' 1.1	10/27/2010	Steve Loar - Changes to allow any type of permits
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitTypeid, sSql, oRs, iNewPermitTypeId

iPermitTypeId = CLng(request("permittypeid"))

iNewPermitTypeId = CopyPermitType( iPermitTypeId )

CopyPermitTypeFees iNewPermitTypeId, iPermitTypeId

CopyPermitTypeReviews iNewPermitTypeId, iPermitTypeId

CopyPermitTypeInspections iNewPermitTypeId, iPermitTypeId

CopyPermitTypeAlerts iNewPermitTypeId, iPermitTypeId

CopyPermitDetailsFields iNewPermitTypeId, iPermitTypeId

CopyPermitTypeCustomFields iNewPermitTypeId, iPermitTypeId


response.redirect "permittypeedit.asp?permittypeid=" & iNewPermitTypeId & "&success=Copy%20Succeeded"


'-------------------------------------------------------------------------------------------------
' integer CopyPermitType( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function CopyPermitType( ByVal iPermitTypeId )
	Dim sSql, oRs, iIsBuildingPermitType, iIsAutoApproved, sExpirationDays, sPublicDescription
	Dim sPermitSubTitle, sPermitRightTitle, sPermitTitleBottom, sPermitFooter, sListFixtures
	Dim sShowConstructionType, sShowFeeTotal, sShowOccupancyType, sShowJobValue, sShowWorkDesc
	Dim sShowFootages, sShowProposedUse, sAdditionalFooterInfo, sPermitTitle, sApprovingOfficial
	Dim iNewPermitTypeId, sPermitLogo, sGroupByInvoiceCategories, sInvoiceLogo, sInvoiceHeader
	Dim sShowCounty, sShowParcelid, sShowPlansBy, sShowPrimaryContact, sHasTempCo, sHasCo, sShowApprovedAsOnTCO, sShowApprovedAsOnCO
	Dim sShowConstTypeOnTCO, sShowConstTypeOnCO, sShowOccTypeonTCO, sShowOccTypeonCO, sShowOccupantsOnTCO, sShowOccupantsOnCO
	Dim sTempCOLogo, sCoLogo, sTempCOTitle, sTempCOSubTitle, sCOTitle, sCOSubTitle, sTempCOAddress, sCOAddress
	Dim sTempCOTopText, sCOTopText, sTempCOBottomText, sCOBottomText, sTempCOCodeRef, sCOCodeRef
	Dim sTempCOApproval, sCOApproval, sTempCOFooter, sCOFooter, sTempCOSubFooter, sCOSubFooter
	Dim sShowTotalSqFt, sShowApprovedAs, sShowFeeTypeTotals, sShowOccupancyUse, iPermitCategoryId
	Dim iPermitLocationRequirementId

	iNewPermitTypeId = 0 

	sSql = "SELECT ISNULL(permittype, '') AS permittype, ISNULL(permittypedesc, '') AS permittypedesc, permitcategoryid, "
	sSql = sSql & " expirationdays, isautoapproved, ISNULL(permitnumberprefix, 'B') AS permitnumberprefix, publicdescription, "
	sSql = sSql & " permittitle, additionalfooterinfo, approvingofficial, permitsubtitle, permitrighttitle, ISNULL(permitlogo,'') AS permitlogo, "
	sSql = sSql & " permittitlebottom, permitfooter, permitsubfooter, listfixtures, showconstructiontype, showfeetotal, "
	sSql = sSql & " showoccupancytype, showjobvalue, showworkdesc, showfootages, showproposeduse, groupbyinvoicecategories, "
	sSql = sSql & " ISNULL(invoicelogo,'') AS invoicelogo, ISNULL(invoiceheader,'') AS invoiceheader, showelectricalcontractor, "
	sSql = sSql & " showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, showcounty, showparcelid, showplansby, "
	sSql = sSql & " showprimarycontact, usetypeid, hastempco, hasco, showapprovedasontco, showconsttypeontco, showocctypeontco, "
	sSql = sSql & " showoccupantsontco, showapprovedasonco, showconsttypeonco, showocctypeonco, showoccupantsonco, "
	sSql = sSql & " ISNULL(tempcologo,'') AS tempcologo, ISNULL(cologo,'') AS cologo, ISNULL(tempcotitle,'') AS tempcotitle, "
	sSql = sSql & " ISNULL(tempcosubtitle,'') AS tempcosubtitle, ISNULL(cotitle,'') AS cotitle, ISNULL(cosubtitle,'') AS cosubtitle, "
	sSql = sSql & " ISNULL(tempcoaddress,'') AS tempcoaddress, ISNULL(coaddress,'') AS coaddress, ISNULL(tempcotoptext,'') AS tempcotoptext, "
	sSql = sSql & " ISNULL(cotoptext,'') AS cotoptext, ISNULL(tempcobottomtext,'') AS tempcobottomtext, ISNULL(cobottomtext,'') AS cobottomtext, "
	sSql = sSql & " ISNULL(tempcocoderef,'') AS tempcocoderef, ISNULL(cocoderef,'') AS cocoderef, ISNULL(tempcoapproval,'') AS tempcoapproval, "
	sSql = sSql & " ISNULL(coapproval,'') AS coapproval, ISNULL(tempcofooter,'') AS tempcofooter, ISNULL(tempcosubfooter,'') AS tempcosubfooter, "
	sSql = sSql & " ISNULL(cofooter,'') AS cofooter, ISNULL(cosubfooter,'') AS cosubfooter, showtotalsqft, showapprovedas, "
	sSql = sSql & " showfeetypetotals, showoccupancyuse, permitlocationrequirementid "
	sSql = sSql & " FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

'	response.write sSql & "<br />"
'	response.End 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iPermitCategoryId = oRs("permitcategoryid")

		iPermitLocationRequirementId = oRs("permitlocationrequirementid")

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

		If IsNull(oRs("permitlogo")) Then
			sPermitLogo = "NULL"
		Else
			sPermitLogo = "'" & DBsafeWithHTML(oRs("permitlogo")) & "'"
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
		If oRs("groupbyinvoicecategories") Then 
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

		sSql = "INSERT INTO egov_permittypes ( orgid, permittype, permittypedesc, "
		sSql = sSql & " permitcategoryid, expirationdays, isautoapproved, displayorder, permitnumberprefix, "
		sSql = sSql & " publicdescription, permittitle, additionalfooterinfo, permitlogo, approvingofficial, "
		sSql = sSql & " permitsubtitle, permitrighttitle, permittitlebottom, permitfooter, permitsubfooter, listfixtures, "
		sSql = sSql & " showconstructiontype, showfeetotal, showoccupancytype, showjobvalue, showworkdesc, showfootages, "
		sSql = sSql & " showproposeduse, groupbyinvoicecategories, invoicelogo, invoiceheader, showelectricalcontractor, "
		sSql = sSql & " showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, showcounty, showparcelid, "
		sSql = sSql & " showplansby, showprimarycontact, usetypeid, hastempco, hasco, showapprovedasontco, showconsttypeontco, "
		sSql = sSql & " showocctypeontco, showoccupantsontco, showapprovedasonco, showconsttypeonco, showocctypeonco, "
		sSql = sSql & " showoccupantsonco, tempcologo, cologo, tempcotitle, tempcosubtitle, cotitle, cosubtitle, tempcoaddress, "
		sSql = sSql & " coaddress, tempcotoptext, cotoptext, tempcobottomtext, cobottomtext, tempcocoderef, cocoderef, "
		sSql = sSql & " tempcoapproval, coapproval, tempcofooter, tempcosubfooter, cofooter, cosubfooter, showtotalsqft, "
		sSql = sSql & " showapprovedas, showfeetypetotals, showoccupancyuse, permitlocationrequirementid ) VALUES ( "
		sSql = sSql & session("orgid") & ", 'Copy of " & Left(oRs("permittype"),42) & "', '"
		sSql = sSql & oRs("permittypedesc") & "', " & iPermitCategoryId & ", " & sExpirationDays & ", " & iIsAutoApproved
		sSql = sSql & ", 1, '" & oRs("permitnumberprefix") & "', " & sPublicDescription & ", " & sPermitTitle 
		sSql = sSql & ", " & sAdditionalFooterInfo & ", " & sPermitLogo & ", " & sApprovingOfficial & ", " & sPermitSubTitle & ", " & sPermitRightTitle
		sSql = sSql & ", " & sPermitTitleBottom & ", " & sPermitFooter & ", " & sPermitSubFooter & ", " & sListFixtures
		sSql = sSql & ", " & sShowConstructionType & ", " & sShowFeeTotal & ", " & sShowOccupancyType & ", " & sShowJobValue
		sSql = sSql & ", " & sShowWorkDesc & ", " & sShowFootages & ", " & sShowProposedUse & ", " & sGroupByInvoiceCategories
		sSql = sSql & ", " & sInvoiceLogo & ", " & sInvoiceHeader & ", " & sShowElectricalContractor 
		sSql = sSql & ", " & sShowMechanicalContractor & ", " & sShowPlumbingContractor & ", " & sShowApplicantLicense
		sSql = sSql & ", " & sShowCounty & ", " & sShowParcelid & ", " & sShowPlansBy & ", " & sShowPrimaryContact & ", " & iUseTypeId
		sSql = sSql & ", " & sHasTempCo & ", " & sHasCO & ", " & sShowApprovedAsOnTCO & ", " & sShowConstTypeOnTCO & ", " & sShowOccTypeonTCO
		sSql = sSql & ", " & sShowOccupantsOnTCO & ", " & sShowApprovedAsOnCO & ", " & sShowConstTypeOnCO & ", " & sShowOccTypeonCO
		sSql = sSql & ", " & sShowOccupantsOnCO & ", " & sTempCOLogo & ", " & sCOLogo & ", " & sTempCOTitle & ", " & sTempCOSubTitle
		sSql = sSql & ", " & sCOTitle & ", " & sCOSubTitle & ", " & sTempCOAddress & ", " & sCOAddress & ", " & sTempCOTopText
		sSql = sSql & ", " & sCOTopText & ", " & sTempCOBottomText & ", " & sCOBottomText & ", " & sTempCOCodeRef & ", " & sCOCodeRef
		sSql = sSql & ", " & sTempCOApproval & ", " & sCOApproval & ", " & sTempCOFooter & ", " & sTempCOSubFooter & ", " & sCOFooter
		sSql = sSql & ", " & sCOSubFooter & ", " & sShowTotalSqFt & ", " & sShowApprovedAs & ", " & sShowFeeTypeTotals
		sSql = sSql & ", " & sShowOccupancyUse & ", " & iPermitLocationRequirementId & " )"
		'response.write sSql & "<br />"

		iNewPermitTypeId = RunIdentityInsert( sSql )
	End If 

	oRs.Close
	Set oRs = Nothing 

	CopyPermitType = iNewPermitTypeId

End Function  


'-------------------------------------------------------------------------------------------------
' void CopyPermitTypeFees iNewPermitTypeId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CopyPermitTypeFees( ByVal iNewPermitTypeId, ByVal iPermitTypeId )
	Dim sSql, oRs, iIsrequired

	sSql = "SELECT permitfeetypeid, isrequired, displayorder "
	sSql = sSql & "FROM egov_permittypes_to_permitfeetypes "
	sSql = sSql & "WHERE permittypeid = " & iPermitTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If oRs("isrequired") Then
			iIsrequired = 1
		Else
			iIsrequired = 0
		End If 

		sSql = "INSERT INTO egov_permittypes_to_permitfeetypes ( permittypeid, permitfeetypeid, isrequired, displayorder ) VALUES ( "
		sSql = sSql & iNewPermitTypeId & ", " & oRs("permitfeetypeid") & ", " & iIsrequired & ", " & oRs("displayorder") & " )"
		
		RunSql sSql 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CopyPermitTypeReviews iNewPermitTypeId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CopyPermitTypeReviews( ByVal iNewPermitTypeId, ByVal iPermitTypeId )
	Dim sSql, oRs, iIsrequired, iReviewerUserId

	sSql = "SELECT permitreviewtypeid, isrequired, revieweruserid, revieworder "
	sSql = sSql & "FROM egov_permittypes_to_permitreviewtypes "
	sSql = sSql & "WHERE permittypeid = " & iPermitTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If oRs("isrequired") Then
			iIsrequired = 1
		Else
			iIsrequired = 0
		End If 

		If IsNull(oRs("revieweruserid")) Then
			iReviewerUserId = "NULL"
		Else
			iReviewerUserId = oRs("revieweruserid")
		End If 

		sSql = "INSERT INTO egov_permittypes_to_permitreviewtypes ( permittypeid, permitreviewtypeid, isrequired, revieweruserid, revieworder ) VALUES ( "
		sSql = sSql & iNewPermitTypeId & ", " & oRs("permitreviewtypeid") & ", " & iIsrequired & ", " & iReviewerUserId & ", " & oRs("revieworder") & " )"
		
		RunSql sSql 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CopyPermitTypeInspections iNewPermitTypeId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CopyPermitTypeInspections( ByVal iNewPermitTypeId, ByVal iPermitTypeId )
	Dim sSql, oRs, iIsrequired, iInspectorUserId, iIsFinal

	sSql = "SELECT permitinspectiontypeid, isrequired, isfinal, inspectionorder, inspectoruserid "
	sSql = sSql & "FROM egov_permittypes_to_permitinspectiontypes "
	sSql = sSql & "WHERE permittypeid = " & iPermitTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If oRs("isrequired") Then
			iIsrequired = 1
		Else
			iIsrequired = 0
		End If 

		If oRs("isfinal") Then
			iIsFinal = 1
		Else
			iIsFinal = 0
		End If

		If IsNull(oRs("inspectoruserid")) Then
			iInspectorUserId = "NULL"
		Else
			iInspectorUserId = oRs("inspectoruserid")
		End If 

		sSql = "INSERT INTO egov_permittypes_to_permitinspectiontypes ( permittypeid, permitinspectiontypeid, isrequired, isfinal, inspectionorder, inspectoruserid ) VALUES ( "
		sSql = sSql & iNewPermitTypeId & ", " & oRs("permitinspectiontypeid") & ", " & iIsrequired & ", " & iIsFinal & ", " & oRs("inspectionorder") & ", " & iInspectorUserId & " )"
		
		RunSql sSql 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CopyPermitTypeAlerts iNewPermitTypeId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CopyPermitTypeAlerts( ByVal iNewPermitTypeId, ByVal iPermitTypeId )
	Dim sSql, oRs, iIsForReviews, iIsForInspections, iNotifyUserId

	sSql = "SELECT orgid, permitalerttypeid, notifyuserid, isforreviews, isforinspections "
	sSql = sSql & "FROM egov_permittypes_to_permitalerttypes "
	sSql = sSql & "WHERE permittypeid = " & iPermitTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If IsNull(oRs("notifyuserid")) Then
			iNotifyUserId = "NULL"
		Else
			iNotifyUserId = oRs("notifyuserid")
		End If 

		If oRs("isforreviews") Then
			iIsForReviews = 1
		Else
			iIsForReviews = 0
		End If 

		If oRs("isforinspections") Then
			iIsForInspections = 1
		Else
			iIsForInspections = 0
		End If 

		sSql = "INSERT INTO egov_permittypes_to_permitalerttypes ( permittypeid, orgid, permitalerttypeid, notifyuserid, isforreviews, isforinspections ) VALUES ( "
		sSql = sSql & iNewPermitTypeId & ", " & oRs("orgid") & ", " & oRs("permitalerttypeid") & ", "
		sSql = sSql & iNotifyUserId & ", " & iIsForReviews & ", " & iIsForInspections & " )"
		
		RunSql sSql 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CopyPermitDetailsFields iNewPermitTypeId, iPermitTypeId 
'-------------------------------------------------------------------------------------------------
Sub CopyPermitDetailsFields( ByVal iNewPermitTypeId, ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT orgid, detailfieldid "
	sSql = sSql & "FROM egov_permittypes_to_permitdetailfields "
	sSql = sSql & "WHERE permittypeid = " & iPermitTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sSql = "INSERT INTO egov_permittypes_to_permitdetailfields ( permittypeid, orgid, detailfieldid ) VALUES ( "
		sSql = sSql & iNewPermitTypeId & ", " & oRs("orgid") & ", " & oRs("detailfieldid") & " )"
		RunSql sSql 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CopyPermitTypeCustomFields iNewPermitTypeId, iPermitTypeId
'-------------------------------------------------------------------------------------------------
Sub CopyPermitTypeCustomFields( ByVal iNewPermitTypeId, ByVal iPermitTypeId )
	Dim sSql, oRs, sIncludeOnReport

	sSql = "SELECT customfieldtypeid, customfieldorder, includeonreport "
	sSql = sSql & "FROM egov_permittypes_to_permitcustomfieldtypes "
	sSql = sSql & "WHERE permittypeid = " & iPermitTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If request("includeonreport") then
			sIncludeOnReport = "1"
		Else
			sIncludeOnReport = "0"
		End If 

		sSql = "INSERT INTO egov_permittypes_to_permitcustomfieldtypes ( permittypeid, customfieldtypeid, customfieldorder, includeonreport ) VALUES ( "
		sSql = sSql & iNewPermitTypeId & ", " & oRs("customfieldtypeid") & ", " & oRs("customfieldorder") & ", " & sIncludeOnReport & " )"
		
		RunSql sSql 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
