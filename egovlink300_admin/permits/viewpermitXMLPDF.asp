<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewpermitXMLPDF.asp
' AUTHOR: Steve Loar
' CREATED: 03/18/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays a PDF document for a permit.
'
' MODIFICATION HISTORY
' 1.0   03/18/2010	Steve Loar - Proof of concept - with hardcoded data
' 1.1	03/23/2010	Steve Loar - Initial Version with real data pulled in
' 1.2	04/06/2010	Steve Loar - Pulled in more fields for Somers documents
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sPermitLocation, sPDFPath, iPermitDocumentId, bIsPermitDocument

iPermitId = CLng(request("permitid"))
iPermitDocumentId = CLng(request("permitdocumentid"))
If clng(request("permitdoc")) = clng(1) Then
	bIsPermitDocument = True 
Else
	bIsPermitDocument = False 
End If 

sPDFPath = GetPermitDocumentInfo( iPermitid, iPermitDocumentId, bIsPermitDocument )

if instr(sPDFPath, "http://www.egovlink.com") > 0 then
	sPDFPath = replace(sPDFPath,"http:","https:")
end if

'if iPermitDocumentId = 5977 or iPermitDocumentId = 5979 or iPermitDocumentId = 5978 then
if session("orgid") = 76 or session("orgid") = 181 or session("orgid") = 200 or session("orgid") = 129 or session("orgid") = 205 or session("orgid") = 8 or session("orgid") = 139 then
	response.redirect replace(lcase(sPDFPath),".pdf",".asp") & "?" & request.servervariables("QUERY_STRING")
end if



If sPDFPath <> "" Then 
	ShowForm sPDFPath, iPermitId
Else
    response.write "<html>" & vbcrlf
    response.write "<head>" & vbcrlf
    response.write "  <script language=""javascript"">" & vbcrlf
    response.write "    function closeWindow() {" & vbcrlf
    response.write "      alert('Could not find the related PDF file." & iPermitDocumentId & "');" & vbcrlf
    response.write "      parent.close();" & vbcrlf
    response.write "    }" & vbcrlf
    response.write "  </script>" & vbcrlf
    response.write "</head>" & vbcrlf
    response.write "<body onload=""closeWindow()"">" & vbcrlf
    response.write "</body>" & vbcrlf
    response.write "</html>" & vbcrlf
End If 



'--------------------------------------------------------------------------------------------------
' void ShowForm( sPDFPath, iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowForm( ByVal sPDFPath, ByVal iPermitId )
	Dim sValue, sLegalDescription, sListedOwner, iPermitAddressId, sPermitLocation, sCounty, sParcelid
	Dim sApplicantAddressLabel, sAttachmentDate, sAttachmentName, sAttachmentDescription, sJobValue
	Dim sReviewList, sReviewStatus, sReviewDate, sReviewer, sReviewNotes, sApplicantName, sTotalFees
	Dim sApplicantEmail, sApplicantPhone, sApplicantCell, sApplicantFax, sInspectionsDue, sInspectorList
	Dim sInspectionList, sInspectionStatus, sInspectionDate, sInspector, sInspectionNotes, sTotalPaid
	Dim sPaymentDate1, sMethod1, sAmount1, sPaymentDate2, sMethod2, sAmount2, sPermitNotes, sIssuedDate
	Dim sReviewList1, sReviewStatus1, sReviewDate1, sReviewer1, sReviewList2, sReviewStatus2, sReviewDate2
	Dim sReviewer2, sReviewList3, sReviewStatus3, sReviewDate3, sReviewer3, sReviewList4, sReviewStatus4
	Dim sReviewDate4, sReviewer4, sReviewList5, sReviewStatus5, sReviewDate5, sReviewer5, sReviewList6
	Dim sReviewStatus6, sReviewDate6, sReviewer6, sInspection1, sInspectionStatus1, sInspectionDate1, sInspector1
	Dim sInspection2, sInspectionStatus2, sInspectionDate2, sInspector2, sInspection3, sInspectionStatus3
	Dim sInspectionDate3, sInspector3, sInspection4, sInspectionStatus4, sInspectionDate4, sInspector4
	Dim sInspection5, sInspectionStatus5, sInspectionDate5, sInspector5, sInspection6, sInspectionStatus6
	Dim sInspectionDate6, sInspector6, iPermitContactId, sPrimaryContractor, sPrimaryContractorLicense
	Dim sCompletedDate, sDimensions, sPlansBy, sUseGroup, sConstructionType, sFinalInspectionDate
	Dim dWaterImpactFee, dRoadImpactFee, dPermitFee, dZoneFee, dBBSFee, dRecImpactFee, dTodaysDate
	Dim sApplicantAddress, sApplicantCity, sApplicantState, sApplicantZip, dCofOFee, dNoZoneTotalFees
	Dim dZoneFeeAmount, dTotalFees

'	This toggle is to switch between a PDF form and viewing the XML for debugging
'	10171 is for Wyoming(200)
'	3102
	If session("userid") = 10171 then
		bShow = True 
	Else 
		bShow = True 
	End If 

	' Get an array of sysbols to swap for the XML values 
	Dim symbolArray()
	getSymbols symbolArray 

	If bShow Then 
		Response.ContentType = "application/vnd.adobe.xdp+xml"
		Response.Charset = "UTF-8"
		response.write "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf  
		response.write "<?xfa generator='AdobeDesigner_V7.0' APIVersion='2.2.4333.0'?>" & vbcrlf
		response.write "<xdp:xdp xmlns:xdp='http://ns.adobe.com/xdp/'>" & vbcrlf
		response.write "<xfa:datasets xmlns:xfa='http://www.xfa.org/schema/xfa-data/1.0/'>" & vbcrlf
		response.write "<xfa:data>" & vbcrlf
	End If 
	response.write "<form1>" & vbcrlf

	' Get the needed data here

	WriteXMLLine symbolArray, "permitno", GetPermitNumber( iPermitId )
	WriteXMLLine symbolArray, "permittypedesc", GetPermitTypeDescForPDF( iPermitId, True )
	WriteXMLLine symbolArray, "status", GetPermitStatusByPermitId( iPermitId )
	WriteXMLLine symbolArray, "statusdate", GetPermitCurrentStatusDate( iPermitId )
	WriteXMLLine symbolArray, "todaysdate",  FormatDateTime(Date(),2)

	' Issued Date
	sIssuedDate = GetPermitIssuedDate( iPermitId )
	if session("orgid") <> "129" then
		WriteXMLLine symbolArray, "issueddate", sIssuedDate
	else
		WriteXMLLine symbolArray, "issueddate", ""
	end if

	' Completed Date
	sCompletedDate = GetPermitDate( "completeddate", iPermitId )
	WriteXMLLine symbolArray, "completeddate", sCompletedDate

	' Location Information
	sPermitLocation = GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, True )
	ReplaceBreaksForPDFs sPermitLocation
	WriteXMLLine symbolArray, "jobsite", sPermitLocation
	WriteXMLLine symbolArray, "owner", sListedOwner
	WriteXMLLine symbolArray, "listedowner", sListedOwner
	WriteXMLLine symbolArray, "county", sCounty
	WriteXMLLine symbolArray, "parcelid", sParcelid

	WriteXMLLine symbolArray, "jobstreetaddress", GetPermitStreetAddress( iPermitId )

	' Applicant address for envelope window
	sApplicantAddressLabel = GetPermitApplicantAddressLabel( iPermitId )
	ReplaceBreaksForPDFs sApplicantAddressLabel
	WriteXMLLine symbolArray, "applicantinfo", sApplicantAddressLabel

	' Applicant contact information
	GetPermitApplicantInfo iPermitId, sApplicantName, sApplicantEmail, sApplicantPhone, sApplicantCell, sApplicantFax, sApplicantAddress, sApplicantCity, sApplicantState, sApplicantZip
	WriteXMLLine symbolArray, "applicantname", sApplicantName
	WriteXMLLine symbolArray, "applicantphone", sApplicantPhone
	WriteXMLLine symbolArray, "applicantcell", sApplicantCell
	WriteXMLLine symbolArray, "applicantfax", sApplicantFax
	WriteXMLLine symbolArray, "applicantemail", sApplicantEmail
	WriteXMLLine symbolArray, "applicantaddress", sApplicantAddress
	WriteXMLLine symbolArray, "applicantcity", sApplicantCity
	WriteXMLLine symbolArray, "applicantstate", sApplicantState
	WriteXMLLine symbolArray, "applicantzip", sApplicantZip

	' Primary Contractor
	sPrimaryContractor = ShowPrimaryContractorForPermit( iPermitId )
	ReplaceBreaksForPDFs sPrimaryContractor
	WriteXMLLine symbolArray, "primarycontractor", sPrimaryContractor

	sPlansBy = GetPermitPlansBy( iPermitid )
	ReplaceBreaksForPDFs sPlansBy
	WriteXMLLine symbolArray, "plansby", sPlansBy

	' Primary Contractor License
	sPrimaryContractorLicense = GetPrimaryContactLicense( iPermitid )
	WriteXMLLine symbolArray, "primarycontractorlicense", sPrimaryContractorLicense

	' Get the 4 contact types here and their licenses
	WriteXMLLine symbolArray, "electricalcompany", GetPermitContractorCompanyByType( iPermitId, "electrical", iPermitContactId )
	WriteXMLLine symbolArray, "electricallicense", GetPermitLicenseByType( iPermitId, "electrical", iPermitContactId )

	WriteXMLLine symbolArray, "plumbingcompany", GetPermitContractorCompanyByType( iPermitId, "plumbing", iPermitContactId )
	WriteXMLLine symbolArray, "plumbinglicense", GetPermitLicenseByType( iPermitId, "plumbing", iPermitContactId )

	WriteXMLLine symbolArray, "mechanicalcompany", GetPermitContractorCompanyByType( iPermitId, "mechanical", iPermitContactId )
	WriteXMLLine symbolArray, "mechanicallicense", GetPermitLicenseByType( iPermitId, "mechanical", iPermitContactId )

	WriteXMLLine symbolArray, "insulationcompany", GetPermitContractorCompanyByType( iPermitId, "insulation", iPermitContactId )
	WriteXMLLine symbolArray, "insulationlicense", GetPermitLicenseByType( iPermitId, "insulation", iPermitContactId )

	' Description of work
	WriteXMLLine symbolArray, "descriptionofwork", GetPermitDetailItemAsString( iPermitId, "descriptionofwork" )

	' Type of construction
	WriteXMLLine symbolArray, "typeofconstruction", GetPermitConstructionType( iPermitId )

	' Use Group, or Occupancy Type, or Occupancy Use - depends on who you ask
	WriteXMLLine symbolArray, "occupancyuse", GetPermitOccupancyTypeGroup( iPermitId ) & " " & GetPermitOccupancyType( iPermitId )

	' Approved As
	WriteXMLLine symbolArray, "approvedas", GetPermitDetailItemAsString( iPermitId, "approvedas" )

	' Existing Use
	WriteXMLLine symbolArray, "existinguse", GetPermitDetailItemAsString( iPermitId, "existinguse" )

	' Proposed Use
	WriteXMLLine symbolArray, "proposeduse", GetPermitDetailItemAsString( iPermitId, "proposeduse" )

	' Zoning
	WriteXMLLine symbolArray, "zoning", GetPermitDetailItemAsString( iPermitId, "zoning" )

	' Residential Units
	WriteXMLLine symbolArray, "residentialunits", GetPermitDetailItemAsNumber( iPermitId, "residentialunits", "integer" )

	' Occupants
	WriteXMLLine symbolArray, "occupants", GetPermitDetailItemAsNumber( iPermitId, "occupants", "integer" )

	' Total Sq Ft
	WriteXMLLine symbolArray, "totalsqft", GetPermitDetailItemAsNumber( iPermitId, "totalsqft", "double" )

	' Permit Issued By
	WriteXMLLine symbolArray, "permitissuedby", GetPermitIssuedBy( iPermitId )

	' Certificate of Occupancy Stipulations, Conditions, Variances
	WriteXMLLine symbolArray, "conditions", GetPermitDetailItemAsString( iPermitId, "conotes" )
	WriteXMLLine symbolArray, "conotes", GetPermitDetailItemAsString( iPermitId, "conotes" )

	WriteXMLLine symbolArray, "tempconotes", GetPermitDetailItemAsString( iPermitId, "tempconotes" )

	For x = 1 To 4 
		GetPermitAttachments iPermitid, x, sAttachmentDate, sAttachmentName, sAttachmentDescription
		WriteXMLLine symbolArray, "attachmentdate" & x, sAttachmentDate
		WriteXMLLine symbolArray, "attachmentname" & x, sAttachmentName
		WriteXMLLine symbolArray, "attachmentdescription" & x, sAttachmentDescription
	Next 

	WriteXMLLine symbolArray, "usetype", GetPermitUseType( iPermitId )
	WriteXMLLine symbolArray, "useclass", GetPermitUseClass( iPermitId )
	WriteXMLLine symbolArray, "workscope", GetPermitWorkScope( iPermitId )
	WriteXMLLine symbolArray, "workclass", GetPermitWorkClass( iPermitId )

	' Job Value
	sJobValue = GetPermitDetailItemAsNumber( iPermitId, "jobvalue", "currency" )
	WriteXMLLine symbolArray, "jobvalue", sJobValue

	' Get the total of all fees
	'sTotalFees = GetPermitDetailItemAsNumber( iPermitId, "feetotal", "currency" )
	dTotalFees = GetPermitDetailItemAsNumber( iPermitId, "feetotal", "double" )
	sTotalFees = FormatCurrency(dTotalFees,2)
	WriteXMLLine symbolArray, "totalfees", sTotalFees

	sFinalInspectionDate = GetPermitFinalInspectionDate( iPermitId )
	WriteXMLLine symbolArray, "finalinspectiondate", sFinalInspectionDate

	dZoneFeeAmount = CDbl(0.00)

	' See if the city has a bbs fee - That makes it Loveland and has all the fee things below
	If OrgHasFeeType( "isbbs" ) Then 
		dFeesSum = CDbl(0.00)

		dZoneFee = GetPermitFeeTypeTotal( iPermitId, "iszone" )
		If CDbl(dZoneFee) = CDbl(0.00) Then
			dZoneFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dZoneFee)
			dZoneFee = FormatCurrency(dZoneFee,2)
			dZoneFeeAmount = CDbl(dZoneFee)
		End If 
		WriteXMLLine symbolArray, "zoningfee", dZoneFee

		dBBSFee = GetPermitFeeTypeTotal( iPermitId, "isbbs" )
		If CDbl(dBBSFee) = CDbl(0.00) Then
			dBBSFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dBBSFee)
			dBBSFee = FormatCurrency(dBBSFee,2)
		End If 
		WriteXMLLine symbolArray, "bbsfee", dBBSFee

		dRecImpactFee = GetPermitFeeTypeTotal( iPermitId, "isrecreationimpact" )
		If CDbl(dRecImpactFee) = CDbl(0.00) Then
			dRecImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dRecImpactFee)
			dRecImpactFee = FormatCurrency(dRecImpactFee,2)
		End If 
		WriteXMLLine symbolArray, "recimpactfee", dRecImpactFee

		dWaterImpactFee = GetPermitFeeTypeTotal( iPermitId, "iswaterimpact" )
		If CDbl(dWaterImpactFee) = CDbl(0.00) Then
			dWaterImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dWaterImpactFee)
			dWaterImpactFee = FormatCurrency(dWaterImpactFee,2)
		End If 
		WriteXMLLine symbolArray, "waterimpactfee", dWaterImpactFee

		dWaterMeterFee = GetPermitFeeTypeTotal( iPermitId, "iswatermeter" )
		If CDbl(dWaterMeterFee) = CDbl(0.00) Then
			dWaterMeterFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dWaterMeterFee)
			dWaterMeterFee = FormatCurrency(dWaterMeterFee,2)
		End If 
		WriteXMLLine symbolArray, "watermeterfee", dWaterMeterFee

		dRoadImpactFee = GetPermitFeeTypeTotal( iPermitId, "isroadimpact" )
		If CDbl(dRoadImpactFee) = CDbl(0.00) Then
			dRoadImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dRoadImpactFee)
			dRoadImpactFee = FormatCurrency(dRoadImpactFee,2)
		End If 
		WriteXMLLine symbolArray, "roadimpactfee", dRoadImpactFee

		dCofOFee = GetPermitFeeTypeTotal( iPermitId, "iscertofoccupancy" )
		If CDbl(dCofOFee) = CDbl(0.00) Then
			dCofOFee = ""
		Else
			' check if the C of O fee is to be removed from the total permit fees
			' This is for Milford, not for Loveland
			If OrgHasFeature( "exclude c of o fee" ) Then 
				dFeesSum = dFeesSum + CDbl(dCofOFee)
			End If 
			dCofOFee = FormatCurrency(dCofOFee,2)
		End If 
		WriteXMLLine symbolArray, "certofoccupancyfee", dCofOFee

		'dTotalFees = GetPermitDetailItemAsNumber( iPermitId, "feetotal", "double" )
		dPermitFee = CDbl(dTotalFees) - dFeesSum
		dPermitFee = FormatCurrency(dPermitFee,2)
		WriteXMLLine symbolArray, "permitfees", dPermitFee
	End If 

	' For Milford, get the total fees less the zoning fees
	dNoZoneTotalFees = CDbl(dTotalFees) - dZoneFeeAmount
	dNoZoneTotalFees = FormatCurrency(dNoZoneTotalFees,2)
	WriteXMLLine symbolArray, "totalnozonefees", dNoZoneTotalFees

	' Get the total paid
	'sTotalPaid = GetPermitDetailItemAsNumber( iPermitId, "totalpaid", "currency" )
	sTotalPaid = GetPaidTotal( iPermitId )
	WriteXMLLine symbolArray, "totalpaid", sTotalPaid

	sDimensions = FormatNumber(GetPermitDetailItemAsNumber( iPermitId, "totalsqft", "integer" ),0) & " sq.ft."
	WriteXMLLine symbolArray, "dimensions", sDimensions

	' Structure Length
	WriteXMLLine symbolArray, "structurelength", GetPermitDetailItemAsString( iPermitId, "structurelength" )

	' Structure width
	WriteXMLLine symbolArray, "structurewidth", GetPermitDetailItemAsString( iPermitId, "structurewidth" )

	' Structure height
	WriteXMLLine symbolArray, "structureheight", GetPermitDetailItemAsString( iPermitId, "structureheight" )

	sConstructionType = GetPermitConstructionType( iPermitId )
	WriteXMLLine symbolArray, "constructiontype", sConstructionType

	sUseGroup = GetPermitOccupancyTypeGroup( iPermitId )
	WriteXMLLine symbolArray, "usegroup", sUseGroup

	' Payment rows here
	GetPaymentForPermit iPermitId, sPaymentDate1, sMethod1, sAmount1, sPaymentDate2, sMethod2, sAmount2
	WriteXMLLine symbolArray, "paymentdate1", sPaymentDate1
	WriteXMLLine symbolArray, "method1", sMethod1
	WriteXMLLine symbolArray, "amount1", sAmount1
	WriteXMLLine symbolArray, "paymentdate2", sPaymentDate2
	WriteXMLLine symbolArray, "method2", sMethod2
	WriteXMLLine symbolArray, "amount2", sAmount2

	' Permit Notes
	sPermitNotes = GetPermitNotes( iPermitId )
	ReplaceBreaksForPDFs sPermitNotes
	WriteXMLLine symbolArray, "permitnotes", sPermitNotes

	' Review Information - This is for Piqua on their plan review and the permit document
	GetPermitReviewListForXML iPermitid, sReviewList1, sReviewStatus1, sReviewDate1, sReviewer1, sReviewList2, sReviewStatus2, sReviewDate2, sReviewer2, sReviewList3, sReviewStatus3, sReviewDate3, sReviewer3, sReviewList4, sReviewStatus4, sReviewDate4, sReviewer4, sReviewList5, sReviewStatus5, sReviewDate5, sReviewer5, sReviewList6, sReviewStatus6, sReviewDate6, sReviewer6, sReviewNotes 
	If sReviewList1 <> "" Then
		ReplaceBreaksForPDFs sReviewList1
		WriteXMLLine symbolArray, "reviewlist1", sReviewList1
		ReplaceBreaksForPDFs sReviewStatus1
		WriteXMLLine symbolArray, "reviewstatus1", sReviewStatus1
		ReplaceBreaksForPDFs sReviewDate1
		WriteXMLLine symbolArray, "reviewdate1", sReviewDate1
		ReplaceBreaksForPDFs sReviewer1
		WriteXMLLine symbolArray, "reviewer1", sReviewer1
	End If 
	If sReviewList2 <> "" Then 
		ReplaceBreaksForPDFs sReviewList2
		WriteXMLLine symbolArray, "reviewlist2", sReviewList2
		ReplaceBreaksForPDFs sReviewStatus2
		WriteXMLLine symbolArray, "reviewstatus2", sReviewStatus2
		ReplaceBreaksForPDFs sReviewDate2
		WriteXMLLine symbolArray, "reviewdate2", sReviewDate2
		ReplaceBreaksForPDFs sReviewer2
		WriteXMLLine symbolArray, "reviewer2", sReviewer2
	End If 
	If sReviewList3 <> "" Then 
		ReplaceBreaksForPDFs sReviewList3
		WriteXMLLine symbolArray, "reviewlist3", sReviewList3
		ReplaceBreaksForPDFs sReviewStatus3
		WriteXMLLine symbolArray, "reviewstatus3", sReviewStatus3
		ReplaceBreaksForPDFs sReviewDate3
		WriteXMLLine symbolArray, "reviewdate3", sReviewDate3
		ReplaceBreaksForPDFs sReviewer3
		WriteXMLLine symbolArray, "reviewer3", sReviewer3
	End If 
	If sReviewList4 <> "" Then 
		ReplaceBreaksForPDFs sReviewList4
		WriteXMLLine symbolArray, "reviewlist4", sReviewList4
		ReplaceBreaksForPDFs sReviewStatus4
		WriteXMLLine symbolArray, "reviewstatus4", sReviewStatus4
		ReplaceBreaksForPDFs sReviewDate4
		WriteXMLLine symbolArray, "reviewdate4", sReviewDate4
		ReplaceBreaksForPDFs sReviewer4
		WriteXMLLine symbolArray, "reviewer4", sReviewer4
	End If 
	If sReviewList5 <> "" Then 
		ReplaceBreaksForPDFs sReviewList5
		WriteXMLLine symbolArray, "reviewlist5", sReviewList5
		ReplaceBreaksForPDFs sReviewStatus5
		WriteXMLLine symbolArray, "reviewstatus5", sReviewStatus5
		ReplaceBreaksForPDFs sReviewDate5
		WriteXMLLine symbolArray, "reviewdate5", sReviewDate5
		ReplaceBreaksForPDFs sReviewer5
		WriteXMLLine symbolArray, "reviewer5", sReviewer5
	End If 
	If sReviewList6 <> "" Then 
		ReplaceBreaksForPDFs sReviewList6
		WriteXMLLine symbolArray, "reviewlist6", sReviewList6
		ReplaceBreaksForPDFs sReviewStatus6
		WriteXMLLine symbolArray, "reviewstatus6", sReviewStatus6
		ReplaceBreaksForPDFs sReviewDate6
		WriteXMLLine symbolArray, "reviewdate6", sReviewDate6
		ReplaceBreaksForPDFs sReviewer6
		WriteXMLLine symbolArray, "reviewer6", sReviewer6
	End If 
	If sReviewNotes <> "" Then 
		ReplaceBreaksForPDFs sReviewNotes
		WriteXMLLine symbolArray, "reviewnotes", sReviewNotes
	End If 

	' Inspections due list
	GetPermitInspectionsList iPermitId, sInspectionsDue, sInspectorList 
	ReplaceBreaksForPDFs sInspectionsDue
	WriteXMLLine symbolArray, "inspectionsdue", sInspectionsDue
	ReplaceBreaksForPDFs sInspectorList
	WriteXMLLine symbolArray, "inspectorlist", sInspectorList

	' Inspection Information
'	GetPermitInspectionsPDF iPermitid, sInspectionList, sInspectionStatus, sInspectionDate, sInspector, sInspectionNotes
'	ReplaceBreaksForPDFs sInspectionList
'	WriteXMLLine symbolArray, "inspectionlist", sInspectionList
'	ReplaceBreaksForPDFs sInspectionStatus
'	WriteXMLLine symbolArray, "inspectionstatus", sInspectionStatus
'	ReplaceBreaksForPDFs sInspectionDate
'	WriteXMLLine symbolArray, "inspectiondate", sInspectionDate
'	ReplaceBreaksForPDFs sInspector
'	WriteXMLLine symbolArray, "inspector", sInspector
'	ReplaceBreaksForPDFs sInspectionNotes
'	WriteXMLLine symbolArray, "inspectionnotes", sInspectionNotes

	' Inspection Information of Piqua for Inspection Statement and Certificate of Compliance
	GetPermitInspectionsXML iPermitid, sInspection1, sInspectionStatus1, sInspectionDate1, sInspector1, sInspection2, sInspectionStatus2, sInspectionDate2, sInspector2, sInspection3, sInspectionStatus3, sInspectionDate3, sInspector3, sInspection4, sInspectionStatus4, sInspectionDate4, sInspector4, sInspection5, sInspectionStatus5, sInspectionDate5, sInspector5, sInspection6, sInspectionStatus6, sInspectionDate6, sInspector6, sInspectionNotes
	If sInspection1 <> "" Then
		ReplaceBreaksForPDFs sInspection1
		WriteXMLLine symbolArray, "inspection1", sInspection1
		ReplaceBreaksForPDFs sInspectionStatus1
		WriteXMLLine symbolArray, "inspectionstatus1", sInspectionStatus1
		ReplaceBreaksForPDFs sInspectionDate1
		WriteXMLLine symbolArray, "inspectiondate1", sInspectionDate1
		ReplaceBreaksForPDFs sInspector1
		WriteXMLLine symbolArray, "inspector1", sInspector1
	End If 
	If sInspection2 <> "" Then
		ReplaceBreaksForPDFs sInspection2
		WriteXMLLine symbolArray, "inspection2", sInspection2
		ReplaceBreaksForPDFs sInspectionStatus2
		WriteXMLLine symbolArray, "inspectionstatus2", sInspectionStatus2
		ReplaceBreaksForPDFs sInspectionDate2
		WriteXMLLine symbolArray, "inspectiondate2", sInspectionDate2
		ReplaceBreaksForPDFs sInspector2
		WriteXMLLine symbolArray, "inspector2", sInspector2
	End If 
	If sInspection3 <> "" Then
		ReplaceBreaksForPDFs sInspection3
		WriteXMLLine symbolArray, "inspection3", sInspection3
		ReplaceBreaksForPDFs sInspectionStatus3
		WriteXMLLine symbolArray, "inspectionstatus3", sInspectionStatus3
		ReplaceBreaksForPDFs sInspectionDate3
		WriteXMLLine symbolArray, "inspectiondate3", sInspectionDate3
		ReplaceBreaksForPDFs sInspector3
		WriteXMLLine symbolArray, "inspector3", sInspector3
	End If 
	If sInspection4 <> "" Then
		ReplaceBreaksForPDFs sInspection4
		WriteXMLLine symbolArray, "inspection4", sInspection4
		ReplaceBreaksForPDFs sInspectionStatus4
		WriteXMLLine symbolArray, "inspectionstatus4", sInspectionStatus4
		ReplaceBreaksForPDFs sInspectionDate4
		WriteXMLLine symbolArray, "inspectiondate4", sInspectionDate4
		ReplaceBreaksForPDFs sInspector4
		WriteXMLLine symbolArray, "inspector4", sInspector4
	End If 
	If sInspection5 <> "" Then
		ReplaceBreaksForPDFs sInspection5
		WriteXMLLine symbolArray, "inspection5", sInspection5
		ReplaceBreaksForPDFs sInspectionStatus5
		WriteXMLLine symbolArray, "inspectionstatus5", sInspectionStatus5
		ReplaceBreaksForPDFs sInspectionDate5
		WriteXMLLine symbolArray, "inspectiondate5", sInspectionDate5
		ReplaceBreaksForPDFs sInspector5
		WriteXMLLine symbolArray, "inspector5", sInspector5
	End If 
	If sInspection6 <> "" Then
		ReplaceBreaksForPDFs sInspection6
		WriteXMLLine symbolArray, "inspection6", sInspection6
		ReplaceBreaksForPDFs sInspectionStatus6
		WriteXMLLine symbolArray, "inspectionstatus6", sInspectionStatus6
		ReplaceBreaksForPDFs sInspectionDate6
		WriteXMLLine symbolArray, "inspectiondate6", sInspectionDate6
		ReplaceBreaksForPDFs sInspector6
		WriteXMLLine symbolArray, "inspector6", sInspector6
	End If 
	ReplaceBreaksForPDFs sInspectionNotes
	WriteXMLLine symbolArray, "inspectionsnotes", sInspectionNotes


	' Get the Custom Field
	IncludeCustomFields symbolArray, iPermitId


	response.write "</form1>" & vbcrlf
	If bShow Then
		response.write "</xfa:data>" & vbcrlf
		response.write "</xfa:datasets>" & vbcrlf

		response.write "<pdf href='" & sPDFPath & "' xmlns='http://ns.adobe.com/xdp/pdf/' />" & vbcrlf
		response.write "</xdp:xdp>" & vbcrl
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' void WriteXMLLine sNodeName, sValue
'--------------------------------------------------------------------------------------------------
Sub WriteXMLLine( ByRef symbolArray, ByVal sNodeName, ByVal sValue )

	' handle reserved XML characters
'	sValue = Replace(sValue, "&", "&amp;")
'	sValue = Replace(sValue, ">", "&gt;")
'	sValue = Replace(sValue, "<", "&lt;")
'	sValue = Replace(sValue, "'", "&apos;")
'	sValue = replace(sValue, "’", "&apos;")
'	sValue = replace(sValue, "‘", "&apos;")
'	sValue = Replace(sValue, "%", "&#37;")
	sValue = Trim(sValue)

	' loop trough the values and replace any symbols we have defined
	For x = 0 To UBound(symbolArray,1)
		sValue = Replace( sValue, symbolArray(x,0), symbolArray(x,1) )
	Next 

	' This part is for the node name characters
	sNodeName = Replace(sNodeName, "(", "")
	sNodeName = Replace(sNodeName, ")", "")
	sNodeName = Replace(sNodeName, "-", "")
	sNodeName = Replace(sNodeName, " ", "")
	sNodeName = Replace(sNodeName, "<", "")
	sNodeName = Replace(sNodeName, ">", "")
	sNodeName = Replace(sNodeName, "&", "")
	sNodeName = Replace(sNodeName, "'", "")
	sNodeName = Replace(sNodeName, "%", "")
	sNodeName = Trim(sNodeName)

	response.write "<" & sNodeName & ">" & sValue & "</" & sNodeName & ">" & vbcrlf

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetPermitContractorCompanyByType( iPermitid, sContractorType, iPermitContactId )
'--------------------------------------------------------------------------------------------------
Function GetPermitContractorCompanyByType( ByVal iPermitId, ByVal sContractorType, ByRef iPermitContactId )
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(P.company,'') AS company, ISNULL(P.firstname,'') AS firstname, ISNULL(P.lastname,'') AS lastname, P.permitcontactid "
	sSql = sSql & "FROM egov_permitcontacts P, egov_permitcontractortypes C "
	sSql = sSql & "WHERE P.contractortypeid = C.contractortypeid AND P.permitid = " & iPermitId
	sSql = sSql & "AND ispriorcontact = 0 AND UPPER(C.contractortype) = UPPER('" & sContractorType & "')"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("company") <> "" Then 
			GetPermitContractorCompanyByType = oRs("company")
		Else
			GetPermitContractorCompanyByType = Trim(oRs("firstname") & " " & oRs("lastname"))
		End If 
		iPermitContactId = oRs("permitcontactid")
	Else
		GetPermitContractorCompanyByType = ""
		iPermitContactId = 0
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitLicenseByType( iPermitid, sLicenseType, iPermitContactId )
'--------------------------------------------------------------------------------------------------
Function GetPermitLicenseByType( ByVal iPermitId, ByVal sLicenseType, ByVal iPermitContactId )
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(L.licensenumber,'') AS licensenumber "
	sSql = sSql & "FROM egov_permitcontacts P, egov_permitcontacts_licenses L, egov_permitlicensetypes T "
	sSql = sSql & "WHERE P.permitid = L.permitid AND P.permitcontactid = L.permitcontactid AND L.licensetypeid = T.licensetypeid "
	sSql = sSql & "AND licenseenddate >= getdate() AND P.permitid = " & iPermitId
	sSql = sSql & " AND P.permitcontactid = " & iPermitContactId
	sSql = sSql & " AND UPPER(T.licensetype) = UPPER('" & sLicenseType & "')"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitLicenseByType = oRs("licensenumber")
	Else
		GetPermitLicenseByType = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitDocumentInfo( iPermitid, iPermitDocumentId )
'--------------------------------------------------------------------------------------------------
Function GetPermitDocumentInfo( ByVal iPermitid, ByVal iPermitDocumentId, ByVal bIsPermitDocument )
	Dim sSql, oRs 

	If bIsPermitDocument Then 
		sSql = "SELECT document, path FROM egov_permitdocuments "
		sSql = sSql & " WHERE documentid = " & iPermitDocumentId
	Else
		sSql = "SELECT D.document, D.path FROM egov_permitdocuments D, egov_permittypes_to_permitdocuments PD, egov_permits P "
		sSql = sSql & " WHERE PD.permittypeid = P.permittypeid AND PD.documentid = D.documentid AND P.permitid = " & iPermitid
		sSql = sSql & " AND PD.permitdocumentid = " & iPermitDocumentId
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitDocumentInfo = oRs("path") & oRs("document")
	Else
		GetPermitDocumentInfo = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' string GetPrimaryContactLicense( iPermitid )
'--------------------------------------------------------------------------------------------------
Function GetPrimaryContactLicense( ByVal iPermitId )
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(L.licensenumber,'') AS licensenumber "
	sSql = sSql & "FROM egov_permitcontacts P, egov_permitcontacts_licenses L, egov_permitlicensetypes T "
	sSql = sSql & "WHERE P.permitid = L.permitid AND P.permitcontactid = L.permitcontactid AND L.licensetypeid = T.licensetypeid "
	sSql = sSql & "AND licenseenddate >= getdate() AND P.permitid = " & iPermitId
	sSql = sSql & " AND P.isprimarycontractor = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPrimaryContactLicense = oRs("licensenumber")
	Else
		GetPrimaryContactLicense = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ReplaceBreaksForPDFs sFieldValue 
'--------------------------------------------------------------------------------------------------
Sub ReplaceBreaksForPDFs( ByRef sFieldValue )
	
	If Not IsNull(sFieldValue) Then 
		sFieldValue = Replace( sFieldValue, "<br />", vbcrlf )
		sFieldValue = Replace( sFieldValue, "<br/>", vbcrlf )
		sFieldValue = Replace( sFieldValue, "<br>", vbcrlf )
	Else
		sFieldValue = ""
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetPaymentForPermit( ByVal iPermitId, ByRef sPaymentDate1, ByRef sMethod1, ByRef sAmount1, ByRef sPaymentDate2, ByRef sMethod2, ByRef sAmount2 )
'--------------------------------------------------------------------------------------------------
Sub GetPaymentForPermit( ByVal iPermitId, ByRef sPaymentDate1, ByRef sMethod1, ByRef sAmount1, ByRef sPaymentDate2, ByRef sMethod2, ByRef sAmount2 )
	Dim sSql, oRs, iCount

	iCount = clng(0) 

	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND L.ispaymentaccount = 1 AND L.permitid = " & iPermitId
	sSql = sSql & " AND J.paymentid IN (SELECT paymentid FROM egov_permitinvoices WHERE isvoided = 0 AND permitid = " & iPermitId & ")"
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF And iCount < 2
		iCount = iCount + clng(1)
		If iCount = clng(1) Then
			sPaymentDate1 = DateValue(CDate(oRs("paymentdate")))
			sMethod1 = GetFirstPaymentMethod( oRs("paymentid") )
			sAmount1 = FormatCurrency(oRs("paymenttotal"),2)
		Else
			sPaymentDate2 = DateValue(CDate(oRs("paymentdate")))
			sMethod2 = GetFirstPaymentMethod( oRs("paymentid") )
			sAmount2 = FormatCurrency(oRs("paymenttotal"),2)
		End If 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean = getSymbols( symbolArray )
'--------------------------------------------------------------------------------------------------
Sub getSymbols( ByRef symbolArray )
	Dim sSql, oRs, symbolCount

	symbolCount = 0

	' first get the row count since you cannot redimension a multidimension array with preserve
	symbolCount = getSymbolCount()

	symbolCount = symbolCount - 1 ' the array starts at 0 so we need one less than we found

	ReDim symbolArray( symbolCount, 1)

	' now repurpose the symbol count to count what we pull from the DB
	symbolCount = -1

	sSql = "SELECT symbol, adobecode FROM SymbolsToHexCodes ORDER BY processorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		symbolCount = symbolCount + 1
		symbolArray( symbolCount, 0 ) = oRs("symbol")
		symbolArray( symbolCount, 1 ) = oRs("adobecode")
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' int = getSymbolCount( )
'--------------------------------------------------------------------------------------------------
Function getSymbolCount( )
	Dim sSql, oRs, symbolCount

	symbolCount = 0

	sSql = "SELECT COUNT(symbolid) AS symbolcount FROM SymbolsToHexCodes"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		symbolCount = clng(oRs("symbolcount"))
	End If 

	oRs.Close
	Set oRs = Nothing 

	getSymbolCount = symbolCount

End Function 


'--------------------------------------------------------------------------------------------------
' IncludeCustomFields iPermitId
'--------------------------------------------------------------------------------------------------
Sub IncludeCustomFields( ByRef symbolArray, ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT P.customfieldid, P.pdffieldname, F.fieldtypebehavior, P.valuelist, P.fieldsize, "
	sSql = sSql & "ISNULL(P.simpletextvalue,'') AS simpletextvalue, ISNULL(P.largetextvalue,'') AS largetextvalue, "
	sSql = sSql & "P.datevalue, moneyvalue, P.intvalue "
	sSql = sSql & "FROM egov_permitcustomfields P, egov_permitfieldtypes F "
	sSql = sSql & "WHERE P.fieldtypeid = F.fieldtypeid AND P.pdffieldname IS NOT NULL AND P.permitid = " & iPermitid
	sSql = sSql & "ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		Select Case oRs("fieldtypebehavior")
				
			Case "checkbox"
				If oRs("simpletextvalue") <> "" Then
					aChoices = Split(oRs("simpletextvalue"),Chr(10))
			
					For x = 0 To UBound(aChoices)
						sField = Trim(Replace(aChoices(x), Chr(13), ""))
						If sField <> "" Then 
							ReplaceBreaksForPDFs sField
							WriteXMLLine symbolArray, LCase("custom_" & oRs("pdffieldname") & "_" & sField) , LCase(Replace(sField," ", ""))
						End If 
					Next 
				End If 

			Case "textarea"
				sField = oRs("largetextvalue")
				ReplaceBreaksForPDFs sField
				WriteXMLLine symbolArray, "custom_" & oRs("pdffieldname") , sField

			Case "date"
				If Not IsNull(oRs("datevalue")) Then
					sField = oRs("datevalue")
					ReplaceBreaksForPDFs sField
					WriteXMLLine symbolArray, "custom_" & oRs("pdffieldname") , sField
				End If 

			Case "money"
				If Not IsNull(oRs("moneyvalue")) Then
					sField = CStr(oRs("moneyvalue"))
					ReplaceBreaksForPDFs sField
					WriteXMLLine symbolArray, "custom_" & oRs("pdffieldname") , sField
				End If 

			Case "integer"
				If Not IsNull(oRs("intvalue")) Then 
					sField = oRs("intvalue")
					ReplaceBreaksForPDFs sField
					WriteXMLLine symbolArray, "custom_" & oRs("pdffieldname") , sField
				End If 

			Case Else 
				sField = oRs("simpletextvalue")
				ReplaceBreaksForPDFs sField
				WriteXMLLine symbolArray, "custom_" & oRs("pdffieldname") , sField

		End Select 

		oRs.MoveNext
	Loop
	
	oRs.Close 
	Set oRs = Nothing 

End Sub



%>
