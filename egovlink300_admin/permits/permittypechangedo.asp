<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypechangedo.asp
' AUTHOR: Steve Loar
' CREATED: 08/18/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This changes the type of a permit. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/18/2010   Steve Loar - INITIAL VERSION
' 1.1	10/27/2010	Steve Loar - Changes to allow any type of permits
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitTypeId, iOriginalPermitTypeId, sSql, oRs, iIsBuildingPermitType, iIsAutoApproved
Dim sExpirationDays, sPublicDescription, sPermitSubTitle, sPermitRightTitle, sPermitTitleBottom, sPermitFooter
Dim sListFixtures, sShowConstructionType, sShowFeeTotal, sShowOccupancyType, sShowJobValue, sShowWorkDesc
Dim sShowFootages, sShowProposedUse, sAdditionalFooterInfo, sPermitTitle, sApprovingOfficial
Dim sPermitLogo, sGroupByInvoiceCategories, sInvoiceLogo, sInvoiceHeader, sShowElectricalContractor
Dim sShowMechanicalContractor, sShowPlumbingContractor, sShowApplicantLicense
Dim sShowCounty, sShowParcelid, sShowPlansBy, sShowPrimaryContact, sHasTempCo, sHasCo, sShowApprovedAsOnTCO, sShowApprovedAsOnCO
Dim sShowConstTypeOnTCO, sShowConstTypeOnCO, sShowOccTypeonTCO, sShowOccTypeonCO, sShowOccupantsOnTCO, sShowOccupantsOnCO
Dim sTempCOLogo, sCoLogo, sTempCOTitle, sTempCOSubTitle, sCOTitle, sCOSubTitle, sTempCOAddress, sCOAddress
Dim sTempCOTopText, sCOTopText, sTempCOBottomText, sCOBottomText, sTempCOCodeRef, sCOCodeRef
Dim sTempCOApproval, sCOApproval, sTempCOFooter, sCOFooter, sTempCOSubFooter, sCOSubFooter
Dim sShowTotalSqFt, sShowApprovedAs, sShowFeeTypeTotals, sShowOccupancyUse, sShowPayments, iDocumentId
Dim sActivityComment, iPermitStatusId, sCurrentPermitType, sNewPermitType

iPermitId = CLng(request("permitid"))
iPermitTypeId = CLng(request("permittypeid"))
iOriginalPermitTypeId = CLng(request("originalpermittypeid"))
sCurrentPermitType = GetPermitTypeDesc( iPermitId, True )

' Get the values for the new permit type
sSql = "SELECT ISNULL(permittype, '') AS permittype, ISNULL(permittypedesc, '') AS permittypedesc,  "
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
sSql = sSql & " showoccupancyuse, showpayments, documentid "
sSql = sSql & " FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	
	If oRs("showoccupancyuse") Then
		sShowOccupancyUse = 1
	Else
		sShowOccupancyUse = 0
	End If
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

	' Update the permit type on the permit table
	sSql = "UPDATE egov_permits"
	sSql = sSql & " SET permittypeid = " & iPermitTypeId
	sSql = sSql & " WHERE permitid = " & iPermitId
	sSql = sSql & " AND permittypeid = " & iOriginalPermitTypeId
	sSql = sSql & " AND orgid = " & session("orgid")
	'response.write sSql & "<br /><br />"
	RunSQL sSql 

	' Update the permit type information in the permitpermittype table
	sSql = "UPDATE egov_permitpermittypes "
	sSql = sSql & "SET permittypeid = " & iPermitTypeId & ", "
	sSql = sSql & "permittype = '" & dbsafe(oRs("permittype")) & "', "
	sSql = sSql & "permittypedesc = '" & dbsafe(oRs("permittypedesc")) & "', "
	sSql = sSql & "expirationdays = " & sExpirationDays & ", "
	sSql = sSql & "isautoapproved = " & iIsAutoApproved & ", "
	sSql = sSql & "permitnumberprefix = '" & oRs("permitnumberprefix") & "', "
	sSql = sSql & "publicdescription = " & sPublicDescription & ", "
	sSql = sSql & "permittitle = " & sPermitTitle & ", "
	sSql = sSql & "additionalfooterinfo = " & sAdditionalFooterInfo & ", "
	sSql = sSql & "approvingofficial = " & sApprovingOfficial & ", "
	sSql = sSql & "permitsubtitle = " & sPermitSubTitle & ", "
	sSql = sSql & "permitrighttitle = " & sPermitRightTitle & ", "
	sSql = sSql & "permittitlebottom = " & sPermitTitleBottom & ", "
	sSql = sSql & "permitfooter = " & sPermitFooter & ", "
	sSql = sSql & "permitsubfooter = " & sPermitSubFooter & ", "
	sSql = sSql & "listfixtures = " & sListFixtures & ", "
	sSql = sSql & "showconstructiontype = " & sShowConstructionType & ", "
	sSql = sSql & "showfeetotal = " & sShowFeeTotal & ", "
	sSql = sSql & "showoccupancytype = " & sShowOccupancyType & ", "
	sSql = sSql & "showjobvalue = " & sShowJobValue & ", "
	sSql = sSql & "showworkdesc = " & sShowWorkDesc & ", "
	sSql = sSql & "showfootages = " & sShowFootages & ", "
	sSql = sSql & "showproposeduse = " & sShowProposedUse & ", "
	sSql = sSql & "permitlogo = " & sPermitLogo & ", "
	sSql = sSql & "groupbyinvoicecategories = " & sGroupByInvoiceCategories & ", "
	sSql = sSql & "invoicelogo = " & sInvoiceLogo & ", "
	sSql = sSql & "invoiceheader = " & sInvoiceHeader & ", "
	sSql = sSql & "showelectricalcontractor = " & sShowElectricalContractor & ", "
	sSql = sSql & "showmechanicalcontractor = " & sShowMechanicalContractor & ", "
	sSql = sSql & "showplumbingcontractor = " & sShowPlumbingContractor & ", "
	sSql = sSql & "showapplicantlicense = " & sShowApplicantLicense & ", "
	sSql = sSql & "showcounty = " & sShowCounty & ", "
	sSql = sSql & "showparcelid = " & sShowParcelid & ", "
	sSql = sSql & "showplansby = " & sShowPlansBy & ", "
	sSql = sSql & "showprimarycontact = " & sShowPrimaryContact & ", "
	sSql = sSql & "hastempco = " & sHasTempCo & ", "
	sSql = sSql & "hasco = " & sHasCO & ", "
	sSql = sSql & "showapprovedasontco = " & sShowApprovedAsOnTCO & ", "
	sSql = sSql & "showconsttypeontco = " & sShowConstTypeOnTCO & ", "
	sSql = sSql & "showocctypeontco = " & sShowOccTypeonTCO & ", "
	sSql = sSql & "showoccupantsontco = " & sShowOccupantsOnTCO & ", "
	sSql = sSql & "showapprovedasonco = " & sShowApprovedAsOnCO & ", "
	sSql = sSql & "showconsttypeonco = " & sShowConstTypeOnCO & ", "
	sSql = sSql & "showocctypeonco = " & sShowOccTypeonCO & ", "
	sSql = sSql & "showoccupantsonco = " & sShowOccupantsOnCO & ", "
	sSql = sSql & "tempcologo = " & sTempCOLogo & ", "
	sSql = sSql & "cologo = " & sCOLogo & ", "
	sSql = sSql & "tempcotitle = " & sTempCOTitle & ", "
	sSql = sSql & "tempcosubtitle = " & sTempCOSubTitle & ", "
	sSql = sSql & "cotitle = " & sCOTitle & ", "
	sSql = sSql & "cosubtitle = " & sCOSubTitle & ", "
	sSql = sSql & "tempcoaddress = " & sTempCOAddress & ", "
	sSql = sSql & "coaddress = " & sCOAddress & ", "
	sSql = sSql & "tempcotoptext = " & sTempCOTopText & ", "
	sSql = sSql & "cotoptext = " & sCOTopText & ", "
	sSql = sSql & "tempcobottomtext = " & sTempCOBottomText & ", "
	sSql = sSql & "cobottomtext = " & sCOBottomText & ", "
	sSql = sSql & "tempcocoderef = " & sTempCOCodeRef & ", "
	sSql = sSql & "cocoderef = " & sCOCodeRef & ", "
	sSql = sSql & "tempcoapproval = " & sTempCOApproval & ", "
	sSql = sSql & "coapproval = " & sCOApproval & ", "
	sSql = sSql & "tempcofooter = " & sTempCOFooter & ", "
	sSql = sSql & "tempcosubfooter = " & sTempCOSubFooter & ", "
	sSql = sSql & "cofooter = " & sCOFooter & ", "
	sSql = sSql & "cosubfooter = " & sCOSubFooter & ", "
	sSql = sSql & "showtotalsqft = " & sShowTotalSqFt & ", "
	sSql = sSql & "showapprovedas = " & sShowApprovedAs & ", "
	sSql = sSql & "showfeetypetotals = " & sShowFeeTypeTotals & ", "
	sSql = sSql & "showoccupancyuse = " & sShowOccupancyUse & ", "
	sSql = sSql & "showpayments = " & sShowPayments & ", "
	sSql = sSql & "documentid = " & iDocumentId 
	sSql = sSql & " WHERE permitid = " & iPermitId
	sSql = sSql & " AND permittypeid = " & iOriginalPermitTypeId
	sSql = sSql & " AND orgid = " & session("orgid")
	'response.write sSql & "<br /><br />"
	RunSQL sSql 
	
End If 

oRs.Close
Set oRs = Nothing 

iPermitStatusId = GetPermitStatusId( iPermitId )
sNewPermitType = GetPermitTypeDesc( iPermitId, True )
sActivityComment = "'Permit type was changed from ''" & dbsafe(sCurrentPermitType) & "'' to ''" & dbsafe(sNewPermitType) & "'' '"

' Put an entry in the log for the change
MakeAPermitLogEntry iPermitid, "'Permit Type Change'", sActivityComment, "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL" 


' return
response.write "SUCCESS"


%>

