<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 01/21/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit types
'
' MODIFICATION HISTORY
' 1.0   01/21/2008   Steve Loar - INITIAL VERSION
' 1.1	07/25/2008	Steve Loar - Inspectors of unassigned added
' 2.0	10/27/2010	Steve Loar - Changes to allow any type of permits
' 2.1	01/11/2011	Steve Loar - Added flag to notify reviewers of attachments
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitTypeid, sSql, isBuildingPermitType, iExpirationDays, sExpirationDays, sPermitTypeDesc
Dim iMaxFeeRows, iMaxInspRows, iMaxReviewRows, x, iIsRequired, sScheduledDaysOut, iIsFinal
Dim sPublicDescription, sPermitNumberPrefix, iMaxReviewAlertRows, iMaxInspectionAlertRows
Dim iPermitInspectorId, iIsFinalPick, sAdditionalFooterInfo, sPermitTitle, sApprovingOfficial
Dim sPermitSubTitle, sPermitRightTitle, sPermitTitleBottom, sPermitFooter, oRs, sListFixtures
Dim sShowConstructionType, sShowFeeTotal, sShowOccupancyType, sShowJobValue, sShowWorkDesc
Dim sShowFootages, sShowProposedUse,sSuccessMsg, sShowOtherContacts, sPermitLogo
Dim sGroupByInvoiceCategories, sInvoiceLogo, sInvoiceHeader, sShowElectricalContractor
Dim sShowMechanicalContractor, sShowPlumbingContractor, sShowApplicantLicense
Dim sShowCounty, sShowParcelid, sShowPlansBy, sShowPrimaryContact, iUseTypeId, sHasTempCo
Dim sHasCo, sShowApprovedAsOnTCO, sShowApprovedAsOnCO, sShowConstTypeOnTCO, sShowConstTypeOnCO
Dim sShowOccTypeonTCO, sShowOccTypeonCO, sShowOccupantsOnTCO, sShowOccupantsOnCO, sTempCOLogo
Dim sCoLogo, sTempCOTitle, sTempCOSubTitle, sCOTitle, sCOSubTitle, sTempCOAddress, sCOAddress
Dim sTempCOTopText, sCOTopText, sTempCOBottomText, sCOBottomText, sTempCOCodeRef, sCOCodeRef
Dim sTempCOApproval, sCOApproval, sTempCOFooter, sCOFooter, sTempCOSubFooter, sCOSubFooter
Dim sShowTotalSqFt, sShowApprovedAs, sShowFeeTypeTotals, sShowOccupancyUse, iLicenseTypeId
Dim sLicenseType, iDisplayOrder, sShowPayments, iNotifyOnRelease, iDocumentId, iMaxDocRows
Dim iDetailFieldId, iPermitCategoryId, iMaxCustomFieldRows, iPermitLocationRequirementId
Dim sIncludeOnReport, sAttachmentReviewerAlert

iPermitTypeid = CLng(request("permittypeid") )

'If request("isbuildingpermittype") = "on" Then
'	isBuildingPermitType = 1
'Else
'	isBuildingPermitType = 0
'End If 

iPermitCategoryId = CLng(request("permitcategoryid"))

iPermitLocationRequirementId = CLng(request("permitlocationrequirementid"))

If request("permittypedesc") = "" Then
	sPermitTypeDesc = "NULL"
Else
	sPermitTypeDesc = "'" & dbsafe(request("permittypedesc")) & "'"
End If 

If request("expirationdays") = "" Then
	sExpirationDays = "NULL"
Else
	sExpirationDays = request("expirationdays")
End If 

If request("publicdescription") = "" Then
	sPublicDescription = "NULL"
Else
	sPublicDescription = "'" & DBsafeWithHTML(request("publicdescription")) & "'"
End If 

If request("isfinal") <> "" Then 
	iIsFinalPick = request("isfinal")
Else
	iIsFinalPick = -1
End If 

sPermitNumberPrefix = "'" & request("permitnumberprefix") & "'"

If request("permittitle") = "" Then
	sPermitTitle = "NULL"
Else
	sPermitTitle =  "'" & DBsafeWithHTML(request("permittitle")) & "'"
End If 

If request("additionalfooterinfo") = "" Then
	sAdditionalFooterInfo = "NULL"
Else
	sAdditionalFooterInfo =  "'" & DBsafeWithHTML(request("additionalfooterinfo")) & "'"
End If 

If request("approvingofficial") = "" Then
	sApprovingOfficial = "NULL"
Else
	sApprovingOfficial =  "'" & DBsafeWithHTML(request("approvingofficial")) & "'"
End If 

If request("permitsubtitle") = "" Then
	sPermitSubTitle = "NULL"
Else
	sPermitSubTitle =  "'" & DBsafeWithHTML(request("permitsubtitle")) & "'"
End If 

If request("permitrighttitle") = "" Then
	sPermitRightTitle = "NULL"
Else
	sPermitRightTitle =  "'" & DBsafeWithHTML(request("permitrighttitle")) & "'"
End If 

If request("permittitlebottom") = "" Then
	sPermitTitleBottom = "NULL"
Else
	sPermitTitleBottom =  "'" & DBsafeWithHTML(request("permittitlebottom")) & "'"
End If 

If request("permitfooter") = "" Then
	sPermitFooter = "NULL"
Else
	sPermitFooter =  "'" & DBsafeWithHTML(request("permitfooter")) & "'"
End If 

If request("permitsubfooter") = "" Then
	sPermitSubFooter = "NULL"
Else
	sPermitSubFooter =  "'" & DBsafeWithHTML(request("permitsubfooter")) & "'"
End If 

If request("permitlogo") = "" Then
	sPermitLogo = "NULL"
Else
	sPermitLogo =  "'" & DBsafeWithHTML(request("permitlogo")) & "'"
End If

If request("listfixtures") = "on" Then 
	sListFixtures = 1
Else 
	sListFixtures = 0
End If 
If request("showconstructiontype") = "on" Then 
	sShowConstructionType = 1
Else 
	sShowConstructionType = 0
End If 
If request("showfeetotal") = "on" Then 
	sShowFeeTotal = 1
Else 
	sShowFeeTotal = 0
End If 
If request("showoccupancytype") = "on" Then 
	sShowOccupancyType = 1
Else 
	sShowOccupancyType = 0
End If 
If request("showjobvalue") = "on" Then 
	sShowJobValue = 1
Else 
	sShowJobValue = 0
End If 
If request("showworkdesc") = "on" Then 
	sShowWorkDesc = 1
Else 
	sShowWorkDesc = 0
End If 
If request("showfootages") = "on" Then 
	sShowFootages = 1
Else 
	sShowFootages = 0
End If 
If request("showproposeduse") = "on" Then 
	sShowProposedUse = 1
Else 
	sShowProposedUse = 0
End If 
If request("showothercontacts") = "on" Then 
	sShowOtherContacts = 1
Else 
	sShowOtherContacts = 0
End If 
If request("groupbyinvoicecategories") = "on" Then 
	sGroupByInvoiceCategories = 1
Else 
	sGroupByInvoiceCategories = 0
End If 
If request("invoicelogo") = "" Then
	sInvoiceLogo = "NULL"
Else
	sInvoiceLogo =  "'" & DBsafeWithHTML(request("invoicelogo")) & "'"
End If
If request("invoiceheader") = "" Then
	sInvoiceHeader = "NULL"
Else
	sInvoiceHeader =  "'" & DBsafeWithHTML(request("invoiceheader")) & "'"
End If
If request("showelectricalcontractor") = "on" Then 
	sShowElectricalContractor = 1
Else 
	sShowElectricalContractor = 0
End If 
If request("showmechanicalcontractor") = "on" Then 
	sShowMechanicalContractor = 1
Else 
	sShowMechanicalContractor = 0
End If 
If request("showplumbingcontractor") = "on" Then 
	sShowPlumbingContractor = 1
Else 
	sShowPlumbingContractor = 0
End If 
If request("showapplicantlicense") = "on" Then 
	sShowApplicantLicense = 1
Else 
	sShowApplicantLicense = 0
End If 
If request("showcounty") = "on" Then 
	sShowCounty = 1
Else 
	sShowCounty = 0
End If
If request("showparcelid") = "on" Then 
	sShowParcelid = 1
Else 
	sShowParcelid = 0
End If
If request("showplansby") = "on" Then 
	sShowPlansBy = 1
Else 
	sShowPlansBy = 0
End If

If request("showprimarycontact") = "on" Then 
	sShowPrimaryContact = 1
Else 
	sShowPrimaryContact = 0
End If

iUseTypeId = CLng(request("usetypeid"))

If request("hastempco") = "on" Then 
	sHasTempCo = 1
Else 
	sHasTempCo = 0
End If 

If request("hasco") = "on" Then 
	sHasCo = 1
Else 
	sHasCo = 0
End If 

If request("showapprovedasontco") = "on" Then 
	sShowApprovedAsOnTCO = 1
Else 
	sShowApprovedAsOnTCO = 0
End If 

If request("showconsttypeontco") = "on" Then 
	sShowConstTypeOnTCO = 1
Else 
	sShowConstTypeOnTCO = 0
End If 

If request("showocctypeontco") = "on" Then 
	sShowOccTypeonTCO = 1
Else 
	sShowOccTypeonTCO = 0
End If 

If request("showoccupantsontco") = "on" Then 
	sShowOccupantsOnTCO = 1
Else 
	sShowOccupantsOnTCO = 0
End If 

If request("showapprovedasonco") = "on" Then 
	sShowApprovedAsOnCO = 1
Else 
	sShowApprovedAsOnCO = 0
End If 

If request("showconsttypeonco") = "on" Then 
	sShowConstTypeOnCO = 1
Else 
	sShowConstTypeOnCO = 0
End If 

If request("showocctypeonco") = "on" Then 
	sShowOccTypeonCO = 1
Else 
	sShowOccTypeonCO = 0
End If 

If request("showoccupantsonco") = "on" Then 
	sShowOccupantsOnCO = 1
Else 
	sShowOccupantsOnCO = 0
End If 
If request("tempcologo") = "" Then
	sTempCOLogo = "NULL"
Else
	sTempCOLogo =  "'" & DBsafeWithHTML(request("tempcologo")) & "'"
End If
If request("cologo") = "" Then
	sCOLogo = "NULL"
Else
	sCOLogo =  "'" & DBsafeWithHTML(request("cologo")) & "'"
End If

If request("tempcotitle") = "" Then
	sTempCOTitle = "NULL"
Else
	sTempCOTitle =  "'" & DBsafeWithHTML(request("tempcotitle")) & "'"
End If
If request("tempcosubtitle") = "" Then
	sTempCOSubTitle = "NULL"
Else
	sTempCOSubTitle =  "'" & DBsafeWithHTML(request("tempcosubtitle")) & "'"
End If
If request("cotitle") = "" Then
	sCOTitle = "NULL"
Else
	sCOTitle =  "'" & DBsafeWithHTML(request("cotitle")) & "'"
End If
If request("cosubtitle") = "" Then
	sCOSubTitle = "NULL"
Else
	sCOSubTitle =  "'" & DBsafeWithHTML(request("cosubtitle")) & "'"
End If
If request("tempcoaddress") = "" Then
	sTempCOAddress = "NULL"
Else
	sTempCOAddress =  "'" & DBsafeWithHTML(request("tempcoaddress")) & "'"
End If
If request("coaddress") = "" Then
	sCOAddress = "NULL"
Else
	sCOAddress =  "'" & DBsafeWithHTML(request("coaddress")) & "'"
End If
If request("tempcotoptext") = "" Then
	sTempCOTopText = "NULL"
Else
	sTempCOTopText =  "'" & DBsafeWithHTML(request("tempcotoptext")) & "'"
End If
If request("cotoptext") = "" Then
	sCOTopText = "NULL"
Else
	sCOTopText =  "'" & DBsafeWithHTML(request("cotoptext")) & "'"
End If
If request("tempcobottomtext") = "" Then
	sTempCOBottomText = "NULL"
Else
	sTempCOBottomText =  "'" & DBsafeWithHTML(request("tempcobottomtext")) & "'"
End If
If request("cobottomtext") = "" Then
	sCOBottomText = "NULL"
Else
	sCOBottomText =  "'" & DBsafeWithHTML(request("cobottomtext")) & "'"
End If
If request("tempcocoderef") = "" Then
	sTempCOCodeRef = "NULL"
Else
	sTempCOCodeRef =  "'" & DBsafeWithHTML(request("tempcocoderef")) & "'"
End If
If request("cocoderef") = "" Then
	sCOCodeRef = "NULL"
Else
	sCOCodeRef =  "'" & DBsafeWithHTML(request("cocoderef")) & "'"
End If
If request("tempcoapproval") = "" Then
	sTempCOApproval = "NULL"
Else
	sTempCOApproval =  "'" & DBsafeWithHTML(request("tempcoapproval")) & "'"
End If
If request("coapproval") = "" Then
	sCOApproval = "NULL"
Else
	sCOApproval =  "'" & DBsafeWithHTML(request("coapproval")) & "'"
End If
If request("tempcofooter") = "" Then
	sTempCOFooter = "NULL"
Else
	sTempCOFooter =  "'" & DBsafeWithHTML(request("tempcofooter")) & "'"
End If 

If request("tempcosubfooter") = "" Then
	sTempCOSubFooter = "NULL"
Else
	sTempCOSubFooter =  "'" & DBsafeWithHTML(request("tempcosubfooter")) & "'"
End If 
If request("cofooter") = "" Then
	sCOFooter = "NULL"
Else
	sCOFooter =  "'" & DBsafeWithHTML(request("cofooter")) & "'"
End If 

If request("cosubfooter") = "" Then
	sCOSubFooter = "NULL"
Else
	sCOSubFooter =  "'" & DBsafeWithHTML(request("cosubfooter")) & "'"
End If 

If request("showtotalsqft") = "on" Then
	sShowTotalSqFt = 1
Else
	sShowTotalSqFt = 0
End If 
If request("showapprovedas") = "on" Then
	sShowApprovedAs = 1
Else
	sShowApprovedAs = 0
End If 
If request("showfeetypetotals") = "on" Then
	sShowFeeTypeTotals = 1
Else
	sShowFeeTypeTotals = 0
End If
If request("showoccupancyuse") = "on" Then
	sShowOccupancyUse = 1
Else
	sShowOccupancyUse = 0
End If
If request("showpayments") = "on" Then
	sShowPayments = 1
Else
	sShowPayments = 0
End If

If request("attachmentrevieweralert") = "on" Then
	sAttachmentReviewerAlert = 1
Else
	sAttachmentReviewerAlert = 0
End If 

iDocumentId = request("documentid")

If iPermitTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permittypes ( orgid, permittypedesc, permittype, permitcategoryid, "
	sSql = sSql & " expirationdays, permitnumberprefix, publicdescription, permittitle, additionalfooterinfo, "
	sSql = sSql & " approvingofficial, permitsubtitle, permitrighttitle, permittitlebottom, permitfooter, "
	sSql = sSql & " permitsubfooter, permitlogo, listfixtures, showconstructiontype, showfeetotal, "
	sSql = sSql & " showoccupancytype, showjobvalue, showworkdesc, showfootages, showproposeduse, showothercontacts, "
	sSql = sSql & " groupbyinvoicecategories, invoicelogo, invoiceheader, showelectricalcontractor, showmechanicalcontractor, "
	sSql = sSql & " showplumbingcontractor, showapplicantlicense, showcounty, showparcelid, showplansby, showprimarycontact, usetypeid, "
	sSql = sSql & " hastempco, hasco, showapprovedasontco, showconsttypeontco, showocctypeontco, showoccupantsontco, "
	sSql = sSql & " showapprovedasonco, showconsttypeonco, showocctypeonco, showoccupantsonco, tempcologo, cologo, "
	sSql = sSql & " tempcotitle, tempcosubtitle, cotitle, cosubtitle, tempcoaddress, coaddress, tempcotoptext, cotoptext, "
	sSql = sSql & " tempcobottomtext, cobottomtext, tempcocoderef, cocoderef, tempcoapproval, coapproval, tempcofooter, "
	sSql = sSql & " tempcosubfooter, cofooter, cosubfooter, showtotalsqft, showapprovedas, showfeetypetotals, showoccupancyuse, "
	sSql = sSql & " showpayments, documentid, permitlocationrequirementid, attachmentrevieweralert ) "
	sSql = sSql & " VALUES ( " & session("orgid") & ", " & sPermitTypeDesc 
	sSql = sSql & ", '" & dbsafe(request("permittype")) & "', " & iPermitCategoryId & ", "
	sSql = sSql & sExpirationDays & ", " & sPermitNumberPrefix & ", " & sPublicDescription & ", " & sPermitTitle & ", " 
	sSql = sSql & sAdditionalFooterInfo & ", " & sApprovingOfficial & ", " & sPermitSubTitle & ", " & sPermitRightTitle
	sSql = sSql & ", " & sPermitTitleBottom & ", " & sPermitFooter & ", " & sPermitSubFooter & ", " & sPermitLogo & ", " & sListFixtures
	sSql = sSql & ", " & sShowConstructionType & ", " & sShowFeeTotal & ", " & sShowOccupancyType & ", " & sShowJobValue
	sSql = sSql & ", " & sShowWorkDesc & ", " & sShowFootages & ", " & sShowProposedUse & ", " & sShowOtherContacts
	sSql = sSql & ", " & sGroupByInvoiceCategories & ", " & sInvoiceLogo & ", " & sInvoiceHeader & ", " & sShowElectricalContractor 
	sSql = sSql & ", " & sShowMechanicalContractor & ", " & sShowPlumbingContractor & ", " & sShowApplicantLicense
	sSql = sSql & ", " & sShowCounty & ", " & sShowParcelid & ", " & sShowPlansBy & ", " & sShowPrimaryContact & ", " & iUseTypeId
	sSql = sSql & ", " & sHasTempCo & ", " & sHasCO & ", " & sShowApprovedAsOnTCO & ", " & sShowConstTypeOnTCO & ", " & sShowOccTypeonTCO
	sSql = sSql & ", " & sShowOccupantsOnTCO & ", " & sShowApprovedAsOnCO & ", " & sShowConstTypeOnCO & ", " & sShowOccTypeonCO
	sSql = sSql & ", " & sShowOccupantsOnCO & ", " & sTempCOLogo & ", " & sCOLogo & ", " & sTempCOTitle & ", " & sTempCOSubTitle
	sSql = sSql & ", " & sCOTitle & ", " & sCOSubTitle & ", " & sTempCOAddress & ", " & sCOAddress & ", " & sTempCOTopText
	sSql = sSql & ", " & sCOTopText & ", " & sTempCOBottomText & ", " & sCOBottomText & ", " & sTempCOCodeRef & ", " & sCOCodeRef
	sSql = sSql & ", " & sTempCOApproval & ", " & sCOApproval & ", " & sTempCOFooter & ", " & sTempCOSubFooter & ", " & sCOFooter
	sSql = sSql & ", " & sCOSubFooter & ", " & sShowTotalSqFt & ", " & sShowApprovedAs & ", " & sShowFeeTypeTotals
	sSql = sSql & ", " & sShowOccupancyUse & ", " & sShowPayments & ", " & iDocumentId & ", " & iPermitLocationRequirementId
	sSql = sSql & ", " & sAttachmentReviewerAlert & " )"
	iPermitTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Permit Type Created"
Else 
	sSql = "UPDATE egov_permittypes SET permittypedesc = " & sPermitTypeDesc
	sSql = sSql & ", permittype = '" & dbsafe(request("permittype"))
	sSql = sSql & "', permitcategoryid = " & iPermitCategoryId
	sSql = sSql & ", expirationdays = " & sExpirationDays
	sSql = sSql & ", permitnumberprefix = " & sPermitNumberPrefix
	sSql = sSql & ", publicdescription = " & sPublicDescription
	sSql = sSql & ", permittitle = " & sPermitTitle
	sSql = sSql & ", additionalfooterinfo = " & sAdditionalFooterInfo
	sSql = sSql & ", approvingofficial = " & sApprovingOfficial
	sSql = sSql & ", permitsubtitle = " & sPermitSubTitle
	sSql = sSql & ", permitrighttitle = " & sPermitRightTitle
	sSql = sSql & ", permittitlebottom = " & sPermitTitleBottom
	sSql = sSql & ", permitfooter = " & sPermitFooter
	sSql = sSql & ", permitsubfooter = " & sPermitSubFooter
	sSql = sSql & ", permitlogo = " & sPermitLogo
	sSql = sSql & ", listfixtures = " & sListFixtures
	sSql = sSql & ", showconstructiontype = " & sShowConstructionType
	sSql = sSql & ", showfeetotal = " & sShowFeeTotal
	sSql = sSql & ", showoccupancytype = " & sShowOccupancyType
	sSql = sSql & ", showjobvalue = " & sShowJobValue
	sSql = sSql & ", showworkdesc = " & sShowWorkDesc
	sSql = sSql & ", showfootages = " & sShowFootages
	sSql = sSql & ", showproposeduse = " & sShowProposedUse
	sSql = sSql & ", showothercontacts = " & sShowOtherContacts
	sSql = sSql & ", groupbyinvoicecategories = " & sGroupByInvoiceCategories
	sSql = sSql & ", invoicelogo = " & sInvoiceLogo
	sSql = sSql & ", invoiceheader = " & sInvoiceHeader
	sSql = sSql & ", showelectricalcontractor = " & sShowElectricalContractor
	sSql = sSql & ", showmechanicalcontractor = " & sShowMechanicalContractor
	sSql = sSql & ", showplumbingcontractor = " & sShowPlumbingContractor
	sSql = sSql & ", showapplicantlicense = " & sShowApplicantLicense
	sSql = sSql & ", showcounty = " & sShowCounty
	sSql = sSql & ", showparcelid = " & sShowParcelid
	sSql = sSql & ", showplansby = " & sShowPlansBy
	sSql = sSql & ", showprimarycontact = " & sShowPrimaryContact
	sSql = sSql & ", usetypeid = " & iUseTypeId
	sSql = sSql & ", hastempco = " & sHasTempCo
	sSql = sSql & ", hasco = " & sHasCo
	sSql = sSql & ", showapprovedasontco = " & sShowApprovedAsOnTCO
	sSql = sSql & ", showconsttypeontco = " & sShowConstTypeOnTCO
	sSql = sSql & ", showocctypeontco = " & sShowOccTypeonTCO
	sSql = sSql & ", showoccupantsontco = " & sShowOccupantsOnTCO
	sSql = sSql & ", showapprovedasonco = " & sShowApprovedAsOnCO
	sSql = sSql & ", showconsttypeonco = " & sShowConstTypeOnCO
	sSql = sSql & ", showocctypeonco = " & sShowOccTypeonCO
	sSql = sSql & ", showoccupantsonco = " & sShowOccupantsOnCO
	sSql = sSql & ", tempcologo = " & sTempCOLogo
	sSql = sSql & ", cologo = " & sCOLogo
	sSql = sSql & ", tempcotitle = " & sTempCOTitle    
	sSql = sSql & ", tempcosubtitle = " & sTempCOSubTitle
	sSql = sSql & ", cotitle = " & sCOTitle
	sSql = sSql & ", cosubtitle = " & sCOSubTitle
	sSql = sSql & ", tempcoaddress = " & sTempCOAddress
	sSql = sSql & ", coaddress = " & sCOAddress
	sSql = sSql & ", tempcotoptext = " & sTempCOTopText
	sSql = sSql & ", cotoptext = " & sCOTopText
	sSql = sSql & ", tempcobottomtext = " & sTempCOBottomText
	sSql = sSql & ", cobottomtext = " & sCOBottomText
	sSql = sSql & ", tempcocoderef = " & sTempCOCodeRef
	sSql = sSql & ", cocoderef = " & sCOCodeRef
	sSql = sSql & ", tempcoapproval = " & sTempCOApproval
	sSql = sSql & ", coapproval = " & sCOApproval
	sSql = sSql & ", tempcofooter = " & sTempCOFooter
	sSql = sSql & ", tempcosubfooter = " & sTempCOSubFooter
	sSql = sSql & ", cofooter = " & sCOFooter
	sSql = sSql & ", cosubfooter = " & sCOSubFooter
	sSql = sSql & ", showtotalsqft = " & sShowTotalSqFt
	sSql = sSql & ", showapprovedas = " & sShowApprovedAs
	sSql = sSql & ", showfeetypetotals = " & sShowFeeTypeTotals
	sSql = sSql & ", showoccupancyuse = " & sShowOccupancyUse
	sSql = sSql & ", showpayments = " & sShowPayments
	sSql = sSql & ", documentid = " & iDocumentId
	sSql = sSql & ", attachmentrevieweralert = " & sAttachmentReviewerAlert
	sSql = sSql & ", permitlocationrequirementid = " & iPermitLocationRequirementId
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permittypeid = " & iPermitTypeid
	RunSQL sSql 

	sSuccessMsg = "Changes Saved"

	' Update any active permits, that is the ones not completed and not voided
	sSql = "SELECT P.permitid FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & "WHERE P.isvoided = 0 AND P.permitstatusid = S.permitstatusid AND "
	sSql = sSql & "S.iscompletedstatus = 0 AND permittypeid = " & iPermitTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	Do While Not oRs.EOF
		sSql = "UPDATE egov_permitpermittypes SET "
		sSql = sSql & " permitcategoryid = " & iPermitCategoryId
		sSql = sSql & ", expirationdays = " & sExpirationDays
		sSql = sSql & ", permitnumberprefix = " & sPermitNumberPrefix
		sSql = sSql & ", publicdescription = " & sPublicDescription
		sSql = sSql & ", permittitle = " & sPermitTitle
		sSql = sSql & ", additionalfooterinfo = " & sAdditionalFooterInfo
		sSql = sSql & ", approvingofficial = " & sApprovingOfficial
		sSql = sSql & ", permitsubtitle = " & sPermitSubTitle
		sSql = sSql & ", permitrighttitle = " & sPermitRightTitle
		sSql = sSql & ", permittitlebottom = " & sPermitTitleBottom
		sSql = sSql & ", permitfooter = " & sPermitFooter
		sSql = sSql & ", permitsubfooter = " & sPermitSubFooter
		sSql = sSql & ", permitlogo = " & sPermitLogo
		sSql = sSql & ", listfixtures = " & sListFixtures
		sSql = sSql & ", showconstructiontype = " & sShowConstructionType
		sSql = sSql & ", showfeetotal = " & sShowFeeTotal
		sSql = sSql & ", showoccupancytype = " & sShowOccupancyType
		sSql = sSql & ", showjobvalue = " & sShowJobValue
		sSql = sSql & ", showworkdesc = " & sShowWorkDesc
		sSql = sSql & ", showfootages = " & sShowFootages
		sSql = sSql & ", showproposeduse = " & sShowProposedUse
		sSql = sSql & ", showothercontacts = " & sShowOtherContacts
		sSql = sSql & ", groupbyinvoicecategories = " & sGroupByInvoiceCategories
		sSql = sSql & ", invoicelogo = " & sInvoiceLogo
		sSql = sSql & ", invoiceheader = " & sInvoiceHeader
		sSql = sSql & ", showelectricalcontractor = " & sShowElectricalContractor
		sSql = sSql & ", showmechanicalcontractor = " & sShowMechanicalContractor
		sSql = sSql & ", showplumbingcontractor = " & sShowPlumbingContractor
		sSql = sSql & ", showapplicantlicense = " & sShowApplicantLicense
		sSql = sSql & ", showcounty = " & sShowCounty
		sSql = sSql & ", showparcelid = " & sShowParcelid
		sSql = sSql & ", showplansby = " & sShowPlansBy
		sSql = sSql & ", showprimarycontact = " & sShowPrimaryContact
		sSql = sSql & ", hastempco = " & sHasTempCo
		sSql = sSql & ", hasco = " & sHasCo
		sSql = sSql & ", showapprovedasontco = " & sShowApprovedAsOnTCO
		sSql = sSql & ", showconsttypeontco = " & sShowConstTypeOnTCO
		sSql = sSql & ", showocctypeontco = " & sShowOccTypeonTCO
		sSql = sSql & ", showoccupantsontco = " & sShowOccupantsOnTCO
		sSql = sSql & ", showapprovedasonco = " & sShowApprovedAsOnCO
		sSql = sSql & ", showconsttypeonco = " & sShowConstTypeOnCO
		sSql = sSql & ", showocctypeonco = " & sShowOccTypeonCO
		sSql = sSql & ", showoccupantsonco = " & sShowOccupantsOnCO
		sSql = sSql & ", tempcologo = " & sTempCOLogo
		sSql = sSql & ", cologo = " & sCOLogo
		sSql = sSql & ", tempcotitle = " & sTempCOTitle    
		sSql = sSql & ", tempcosubtitle = " & sTempCOSubTitle
		sSql = sSql & ", cotitle = " & sCOTitle
		sSql = sSql & ", cosubtitle = " & sCOSubTitle
		sSql = sSql & ", tempcoaddress = " & sTempCOAddress
		sSql = sSql & ", coaddress = " & sCOAddress
		sSql = sSql & ", tempcotoptext = " & sTempCOTopText
		sSql = sSql & ", cotoptext = " & sCOTopText
		sSql = sSql & ", tempcobottomtext = " & sTempCOBottomText
		sSql = sSql & ", cobottomtext = " & sCOBottomText
		sSql = sSql & ", tempcocoderef = " & sTempCOCodeRef
		sSql = sSql & ", cocoderef = " & sCOCodeRef
		sSql = sSql & ", tempcoapproval = " & sTempCOApproval
		sSql = sSql & ", coapproval = " & sCOApproval
		sSql = sSql & ", tempcofooter = " & sTempCOFooter
		sSql = sSql & ", tempcosubfooter = " & sTempCOSubFooter
		sSql = sSql & ", cofooter = " & sCOFooter
		sSql = sSql & ", cosubfooter = " & sCOSubFooter
		sSql = sSql & ", showtotalsqft = " & sShowTotalSqFt
		sSql = sSql & ", showapprovedas = " & sShowApprovedAs
		sSql = sSql & ", showfeetypetotals = " & sShowFeeTypeTotals
		sSql = sSql & ", showoccupancyuse = " & sShowOccupancyUse
		sSql = sSql & ", showpayments = " & sShowPayments
		sSql = sSql & ", documentid = " & iDocumentId
		sSql = sSql & ", attachmentrevieweralert = " & sAttachmentReviewerAlert
		sSql = sSql & " WHERE permitid = " & oRs("permitid") & " AND permittypeid = " & iPermitTypeid
		RunSQL sSql 

'		If CLng(iUseTypeId) > CLng(0) Then 
'			sSql = "UPDATE egov_permits SET usetypeid = " & iUseTypeId
'			sSql = sSql & " WHERE permitid = " & oRs("permitid")
'			RunSQL sSql 
'		End If 

		' Change the Required Licenses and Show Licenses rows for each active permit
		sSql = "DELETE FROM egov_permits_to_permitlicensetypes WHERE permitid = " & oRs("permitid")
		RunSQL sSql
		
		For Each iLicenseTypeId In Request("licensetypeid")
			GetPermitTypeLicenseDetails iLicenseTypeId, sLicenseType, iDisplayOrder
			sSql = "INSERT INTO egov_permits_to_permitlicensetypes ( permitid, permittypeid, licensetypeid, isrequired, orgid, licensetype, displayorder ) VALUES ( " 
			sSql = sSql & oRs("permitid") & ", " & iPermitTypeid & ", " & iLicenseTypeId & ", 1, " & session("orgid") & ", '" & sLicenseType & "', " & iDisplayOrder & " )"
			RunSQL sSql
		Next 

		For Each iLicenseTypeId In Request("showlicensetypeid")
			GetPermitTypeLicenseDetails iLicenseTypeId, sLicenseType, iDisplayOrder
			sSql = "INSERT INTO egov_permits_to_permitlicensetypes ( permitid, permittypeid, licensetypeid, isrequired, orgid, licensetype, displayorder ) VALUES ( " 
			sSql = sSql & oRs("permitid") & ", " & iPermitTypeid & ", " & iLicenseTypeId & ", 0, " & session("orgid") & ", '" & sLicenseType & "', " & iDisplayOrder & " )"
			RunSQL sSql
		Next 

		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	' Delete any required licenses, display licenses, fees, inspections, reviews, documents, detail fields and custom fields
	sSql = "DELETE FROM egov_permittypes_to_permitfeetypes WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitinspectiontypes WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitreviewtypes WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitalerttypes WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitlicensetypes WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitdocuments WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitdetailfields WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql
	sSql = "DELETE FROM egov_permittypes_to_permitcustomfieldtypes WHERE permittypeid = " & iPermitTypeid
	RunSQL sSql

End If 

' Add any required licenses, fees, inspections, reviews and custom fields
iMaxFeeRows = CLng(request("maxfeerows"))
iMaxInspRows = CLng(request("maxinspectionrows"))
iMaxReviewRows = CLng(request("maxreviewrows"))
iMaxReviewAlertRows = CLng(request("maxreviewalertrows"))
iMaxInspectionAlertRows = CLng(request("maxinspectionalertrows"))
iMaxDocRows = CLng(request("maxdocumentrows"))
iMaxCustomFieldRows = CLng(request("maxcustomfieldrows"))

For x = 0 To iMaxDocRows
	' If they picked adocument and gave it a label then save the data
	If CLng(request("documentid" & x)) <> CLng(0) And request("documentlabel" & x) <> "" Then 
		sSql = "INSERT INTO egov_permittypes_to_permitdocuments (permittypeid, documentid, orgid, documentlabel) VALUES ( "
		sSql = sSql & iPermitTypeid & ", " & CLng(request("documentid" & x)) & ", " & session("orgid") & ", '" & dbsafe(request("documentlabel" & x)) & "' )"
		RunSQL sSql
	End If 
Next 

For x = 0 To iMaxFeeRows
	' See if the fee type data exists
	If request("permitfeetypeid" & x) <> "" Then 
		' If they picked an actual fee type then save the data
		If CLng(request("permitfeetypeid" & x)) <> CLng(0) Then 
			If request("isrequired" & x) = "on" Then 
				iIsRequired = 1
			Else
				iIsRequired = 0
			End If 
			sSql = "INSERT INTO egov_permittypes_to_permitfeetypes (permittypeid, permitfeetypeid, isrequired, displayorder) VALUES ( "
			sSql = sSql & iPermitTypeid & ", " & CLng(request("permitfeetypeid" & x)) & ", " & iIsRequired & ", " & request("displayorder" & x) & " )"
			RunSQL sSql
		End If 
	End If 
Next 

ReorderFeeRows iPermitTypeid

For x = 0 To iMaxInspRows
	' See if the fee type data exists
	If request("permitinspectiontypeid" & x) <> "" Then 
		' If they picked an actual inspection type then save the data
		If CLng(request("permitinspectiontypeid" & x)) <> CLng(0) Then 
			If request("inspectionisrequired" & x) = "on" Then 
				iIsRequired = 1
			Else
				iIsRequired = 0
			End If 
'			If request("isfinal" & x) = "on" Then 
'				iIsFinal = 1
'			Else
'				iIsFinal = 0
'			End If 
			If CLng(iIsFinalPick) = CLng(x) Then 
				iIsFinal = 1
			Else
				iIsFinal = 0
			End If 
			If CLng(request("permitinspectorid" & x)) = CLng(0) Then
				iPermitInspectorId = "NULL"
			Else
				iPermitInspectorId = request("permitinspectorid" & x)
			End If 
			sSql = "INSERT INTO egov_permittypes_to_permitinspectiontypes ( permittypeid, permitinspectiontypeid, isrequired, inspectionorder, isfinal, inspectoruserid ) VALUES ( "
			sSql = sSql & iPermitTypeid & ", " & CLng(request("permitinspectiontypeid" & x)) & ", " & iIsRequired & ", " & request("inspectionorder" & x) & ", " & iIsFinal & ", " & iPermitInspectorId & " )"
			RunSQL sSql
		End If 
	End If 
Next 

ReorderInspectionRows iPermitTypeid 

For x = 0 To iMaxReviewRows
	' See if the fee type data exists
	If request("permitreviewtypeid" & x) <> "" Then 
		' If they picked an actual review type then save the data
		If CLng(request("permitreviewtypeid" & x)) <> CLng(0) Then 
			If request("reviewisrequired" & x) = "on" Then 
				iIsRequired = 1
			Else
				iIsRequired = 0
			End If 
			If request("notifyonrelease" & x) = "on" Then 
				iNotifyOnRelease = 1
			Else
				iNotifyOnRelease = 0
			End If 
			sSql = "INSERT INTO egov_permittypes_to_permitreviewtypes (permittypeid, permitreviewtypeid, isrequired, notifyonrelease, revieworder, revieweruserid ) VALUES ( "
			sSql = sSql & iPermitTypeid & ", " & CLng(request("permitreviewtypeid" & x)) & ", " & iIsRequired & ", " & iNotifyOnRelease & ", " & request("revieworder" & x) & ", " & request("permitreviewerid" & x) & " )"
			RunSQL sSql
		End If 
	End If 
Next 

ReorderReviewRows iPermitTypeid 

' Add in the review alert rows
For x = 0 To iMaxReviewAlertRows
	' See if the alert type was picked
	If CLng(request("permitalerttypeid" & x)) <> CLng(0) Then 
		sSql = "INSERT INTO egov_permittypes_to_permitalerttypes (orgid, permittypeid, permitalerttypeid, notifyuserid, isforreviews) VALUES ( "
		sSql = sSql & session("orgid") & ", " & iPermitTypeid & ", " & CLng(request("permitalerttypeid" & x)) & ", " & CLng(request("notifyuserid" & x)) & ", 1 )"
		RunSQL sSql
	End If 
Next 

' Add in the Inspection alert rows
For x = 0 To iMaxInspectionAlertRows
	' See if the alert type was picked
	If CLng(request("permitinspectionalerttypeid" & x)) <> CLng(0) Then 
		sSql = "INSERT INTO egov_permittypes_to_permitalerttypes (orgid, permittypeid, permitalerttypeid, notifyuserid, isforinspections) VALUES ( "
		sSql = sSql & session("orgid") & ", " & iPermitTypeid & ", " & CLng(request("permitinspectionalerttypeid" & x)) & ", " & CLng(request("notifyinspectoruserid" & x)) & ", 1 )"
		RunSQL sSql
	End If 
Next 

' Add in the Required Licenses rows
For Each iLicenseTypeId In Request("licensetypeid")
	sSql = "INSERT INTO egov_permittypes_to_permitlicensetypes ( permittypeid, licensetypeid, isrequired ) VALUES ( " 
	sSql = sSql & iPermitTypeid & ", " & iLicenseTypeId & ", 1 )"
	RunSQL sSql
Next 

For Each iLicenseTypeId In Request("showlicensetypeid")
	sSql = "INSERT INTO egov_permittypes_to_permitlicensetypes ( permittypeid, licensetypeid, isrequired ) VALUES ( " 
	sSql = sSql & iPermitTypeid & ", " & iLicenseTypeId & ", 0 )"
	RunSQL sSql
Next 

' Add in the detail tab fields
For Each iDetailFieldId In request("detailfieldid")
	sSql = "INSERT INTO egov_permittypes_to_permitdetailfields ( permittypeid, detailfieldid, orgid ) VALUES ( " 
	sSql = sSql & iPermitTypeid & ", " & iDetailFieldId & ", " & session("orgid") & " )"
	RunSQL sSql
Next 

' Add the custom fields
For x = 0 To iMaxCustomFieldRows
	If request("customfieldtypeid" & x) <> "" Then 
		If CLng(request("customfieldtypeid" & x)) <> CLng(0) Then 
			If request("includeonreport" & x) = "on" then
				sIncludeOnReport = "1"
			Else
				sIncludeOnReport = "0"
			End If 
			sSql = "INSERT INTO egov_permittypes_to_permitcustomfieldtypes ( permittypeid, customfieldtypeid, customfieldorder, includeonreport ) VALUES ( "
			sSql = sSql & iPermitTypeid & ", " & CLng(request("customfieldtypeid" & x)) & ", " & request("customfieldorder" & x) & ", "
			sSql = sSql & sIncludeOnReport & " )"
			RunSQL sSql
		End If 
	End If 
Next 

ReorderCustomFieldRows iPermitTypeid 


response.redirect "permittypeedit.asp?permittypeid=" & iPermitTypeid & "&activetab=" & request("activetab") & "&success=" & sSuccessMsg


'-------------------------------------------------------------------------------------------------
' void ReorderFeeRows( iPermitTypeid )
'-------------------------------------------------------------------------------------------------
Sub ReorderFeeRows( ByVal iPermitTypeid )
	Dim iNewOrder, oRs, sSql

	iNewOrder = CLng(0)
	
	sSql = "SELECT permittypeid, permitfeetypeid, displayorder FROM egov_permittypes_to_permitfeetypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid & " ORDER BY displayorder, permitfeetypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 1, 3

	Do While Not oRs.EOF
		iNewOrder = iNewOrder + CLng(1)
		oRs("displayorder") = iNewOrder
		oRs.Update
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void ReorderInspectionRows( iPermitTypeid )
'-------------------------------------------------------------------------------------------------
Sub ReorderInspectionRows( ByVal iPermitTypeid )
	Dim iNewOrder, oRs, sSql

	iNewOrder = CLng(0)
	
	sSql = "SELECT permittypeid, permitinspectiontypeid, inspectionorder FROM egov_permittypes_to_permitinspectiontypes WHERE permittypeid = " & iPermitTypeid & " ORDER BY inspectionorder, permitinspectiontypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 1, 3

	Do While Not oRs.EOF
		iNewOrder = iNewOrder + CLng(1)
		oRs("inspectionorder") = iNewOrder
		oRs.Update
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void ReorderReviewRows( iPermitTypeid )
'-------------------------------------------------------------------------------------------------
Sub ReorderReviewRows( ByVal iPermitTypeid )
	Dim iNewOrder, oRs, sSql

	iNewOrder = CLng(0)
	
	sSql = "SELECT permittypeid, permitreviewtypeid, revieworder FROM egov_permittypes_to_permitreviewtypes WHERE permittypeid = " & iPermitTypeid & " ORDER BY revieworder, permitreviewtypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 1, 3

	Do While Not oRs.EOF
		iNewOrder = iNewOrder + CLng(1)
		oRs("revieworder") = iNewOrder
		oRs.Update
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void ReorderCustomFieldRows( iPermitTypeid )
'-------------------------------------------------------------------------------------------------
Sub ReorderCustomFieldRows( ByVal iPermitTypeid )
	Dim iNewOrder, oRs, sSql

	iNewOrder = CLng(0)
	
	sSql = "SELECT permittypeid, customfieldtypeid, customfieldorder "
	sSql = sSql & "FROM egov_permittypes_to_permitcustomfieldtypes "
	ssql = sSql & "WHERE permittypeid = " & iPermitTypeid & " ORDER BY customfieldorder, customfieldtypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 1, 3

	Do While Not oRs.EOF
		iNewOrder = iNewOrder + CLng(1)
		oRs("customfieldorder") = iNewOrder
		oRs.Update
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
