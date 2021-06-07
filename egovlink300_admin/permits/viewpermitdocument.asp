<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewpermitdocument.asp
' AUTHOR: Steve Loar
' CREATED: 06/04/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays a document for a permit.
'
' MODIFICATION HISTORY
' 1.0   06/04/2009	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sPermitLocation, sPDFPath, iPermitDocumentId

iPermitId = CLng(request("permitid"))
iPermitDocumentId = CLng(request("permitdocumentid"))

sPDFPath = GetPermitDocumentInfo( iPermitid, iPermitDocumentId )

If sPDFPath <> "" Then 
	ShowForm sPDFPath, iPermitId
Else
    response.write "<html>" & vbcrlf
    response.write "<head>" & vbcrlf
    response.write "  <script language=""javascript"">" & vbcrlf
    response.write "    function closeWindow() {" & vbcrlf
    response.write "      alert('Could not find the related PDF file.');" & vbcrlf
    response.write "      parent.close();" & vbcrlf
    response.write "    }" & vbcrlf
    response.write "  </script>" & vbcrlf
    response.write "</head>" & vbcrlf
    response.write "<body onload=""closeWindow()"">" & vbcrlf
    response.write "</body>" & vbcrlf
    response.write "</html>" & vbcrlf
End If 


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowForm( sPDFPath, iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowForm( ByVal sPDFPath, ByVal iPermitId )
	Dim oPDF, oDocument, r, sPermitNo, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid
	Dim sPermitNotes, sPrimaryContact, sApprovedAs, sPrimaryContractor, sPlansBy, sConstructionType, sUseGroup
	Dim sOccupants, sTotalFees, sIssuedDate, sJobValue, sDimensions, dPermitFee, dZoneFee, dBBSFee, dRecImpactFee
	Dim dWaterImpactFee, dRoadImpactFee, dFeesSum, sTotalPaid, sDescriptionOfWork, sPaymentDate1, sMethod1
	Dim sAmount1, sPaymentDate2, sMethod2, sAmount2, sCONotes, sTempCONotes, sPermitJobAddress, sStatus
	Dim sStatusDate, sApplicantAddressLabel, sAttachmentDate, sAttachmentName, sAttachmentDescription, x
	Dim sInspectionsDue, sInspectorList, sReviewList, sReviewStatus, sReviewDate, sReviewer, sReviewNotes
	Dim sInspectionList, sInspectionStatus, sInspectionDate, sInspector, sInspectionNotes, sPermitTypeDescription

  'Create PDF object
  	set oPDF  = Server.CreateObject("APToolkit.Object")
  	oDocument = oPDF.OpenOutputFile("MEMORY") 'CREATE THE OUTPUT INMEMORY

 	'Build PDF
  	oPDF.OutputPageWidth  = 612 ' 8.5 inches
  	oPDF.OutputPageHeight = 792 ' 11 inches

 	'Add form
	r = oPDF.OpenInputFile( sPDFPath )

	'Add data to form
  	'PopulateFormwithData oPDF,iRequestID 

	sPermitNo = GetPermitNumber( iPermitId )
	setPDFFormFieldData oPDF, "permitno", sPermitNo, 1

	sPermitTypeDescription = GetPermitTypeDescForPDF( iPermitId, True )
	setPDFFormFieldData oPDF, "permittypedesc", sPermitTypeDescription, 1

	sStatus = GetPermitStatusByPermitId( iPermitId )
	setPDFFormFieldData oPDF, "status", sStatus, 1

	sStatusDate = GetPermitCurrentStatusDate( iPermitId )
	setPDFFormFieldData oPDF, "statusdate", sStatusDate, 1

	sPermitLocation = GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, True )
	ReplaceBreaksForPDFs sPermitLocation
	setPDFFormFieldData oPDF, "jobsite", sPermitLocation, 1
	setPDFFormFieldData oPDF, "owner", sListedOwner, 1
	setPDFFormFieldData oPDF, "county", sCounty, 1
	setPDFFormFieldData oPDF, "parcelid", sParcelid, 1

	sPermitJobAddress = GetPermitStreetAddress( iPermitId )
	setPDFFormFieldData oPDF, "jobstreetaddress", sPermitJobAddress, 1

	sApplicantAddressLabel = GetPermitApplicantAddressLabel( iPermitId )
	ReplaceBreaksForPDFs sApplicantAddressLabel
	setPDFFormFieldData oPDF, "applicantinfo", sApplicantAddressLabel, 1

	sPrimaryContractor = ShowPrimaryContractorForPermit( iPermitId )
	ReplaceBreaksForPDFs sPrimaryContractor
	setPDFFormFieldData oPDF, "primarycontractor", sPrimaryContractor, 1

	sPlansBy = GetPermitPlansBy( iPermitid )
	ReplaceBreaksForPDFs sPlansBy
	setPDFFormFieldData oPDF, "plansby", sPlansBy, 1

	sPrimaryContact = GetPermitDetailItemAsString( iPermitId, "primarycontact" )
	ReplaceBreaksForPDFs sPrimaryContact
	setPDFFormFieldData oPDF, "primarycontact", sPrimaryContact, 1

	sConstructionType = GetPermitConstructionType( iPermitId )
	setPDFFormFieldData oPDF, "constructiontype", sConstructionType, 1

	sUseGroup = GetPermitOccupancyTypeGroup( iPermitId )
	setPDFFormFieldData oPDF, "usegroup", sUseGroup, 1

	sOccupants = GetPermitDetailItemAsNumber( iPermitId, "occupants", "integer" )
	setPDFFormFieldData oPDF, "occupants", sOccupants, 1

	sApprovedAs = GetPermitDetailItemAsString( iPermitId, "approvedas" )
	setPDFFormFieldData oPDF, "approvedas", sApprovedAs, 1

	sCONotes = GetPermitDetailItemAsString( iPermitId, "conotes" )
	setPDFFormFieldData oPDF, "conotes", sCONotes, 1

	sTempCONotes = GetPermitDetailItemAsString( iPermitId, "tempconotes" )
	setPDFFormFieldData oPDF, "tempconotes", sTempCONotes, 1

	sPermitNotes = GetPermitNotes( iPermitId )
	ReplaceBreaksForPDFs sPermitNotes
	setPDFFormFieldData oPDF, "permitnotes", sPermitNotes, 1

	sIssuedDate = GetPermitIssuedDate( iPermitId )
	setPDFFormFieldData oPDF, "issueddate", sIssuedDate, 1

	sFinalInspectionDate = GetPermitFinalInspectionDate( iPermitId )
	setPDFFormFieldData oPDF, "finalinspectiondate", sFinalInspectionDate, 1

	' Get the total of all fees
	sTotalFees = GetPermitDetailItemAsNumber( iPermitId, "feetotal", "currency" )
	setPDFFormFieldData oPDF, "totalfees", sTotalFees, 1

	' See if the city has a bbs fee - That makes it Loveland and has all the fee things below
	If OrgHasFeeType( "isbbs" ) Then 
		dFeesSum = CDbl(0.00)

		dZoneFee = GetPermitFeeTypeTotal( iPermitId, "iszone" )
		If CDbl(dZoneFee) = CDbl(0.00) Then
			dZoneFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dZoneFee)
			dZoneFee = FormatCurrency(dZoneFee,2)
		End If 
		setPDFFormFieldData oPDF, "zoningfee", dZoneFee, 1

		dBBSFee = GetPermitFeeTypeTotal( iPermitId, "isbbs" )
		If CDbl(dBBSFee) = CDbl(0.00) Then
			dBBSFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dBBSFee)
			dBBSFee = FormatCurrency(dBBSFee,2)
		End If 
		setPDFFormFieldData oPDF, "bbsfee", dBBSFee, 1

		dRecImpactFee = GetPermitFeeTypeTotal( iPermitId, "isrecreationimpact" )
		If CDbl(dRecImpactFee) = CDbl(0.00) Then
			dRecImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dRecImpactFee)
			dRecImpactFee = FormatCurrency(dRecImpactFee,2)
		End If 
		setPDFFormFieldData oPDF, "recimpactfee", dRecImpactFee, 1

		dWaterImpactFee = GetPermitFeeTypeTotal( iPermitId, "iswaterimpact" )
		If CDbl(dWaterImpactFee) = CDbl(0.00) Then
			dWaterImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dWaterImpactFee)
			dWaterImpactFee = FormatCurrency(dWaterImpactFee,2)
		End If 
		setPDFFormFieldData oPDF, "waterimpactfee", dWaterImpactFee, 1

		dWaterMeterFee = GetPermitFeeTypeTotal( iPermitId, "iswatermeter" )
		If CDbl(dWaterMeterFee) = CDbl(0.00) Then
			dWaterMeterFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dWaterMeterFee)
			dWaterMeterFee = FormatCurrency(dWaterMeterFee,2)
		End If 
		setPDFFormFieldData oPDF, "watermeterfee", dWaterMeterFee, 1

		dRoadImpactFee = GetPermitFeeTypeTotal( iPermitId, "isroadimpact" )
		If CDbl(dRoadImpactFee) = CDbl(0.00) Then
			dRoadImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dRoadImpactFee)
			dRoadImpactFee = FormatCurrency(dRoadImpactFee,2)
		End If 
		setPDFFormFieldData oPDF, "roadimpactfee", dRoadImpactFee, 1

		sTotalFees = GetPermitDetailItemAsNumber( iPermitId, "feetotal", "double" )
		dPermitFee = CDbl(sTotalFees) - dFeesSum
		dPermitFee = FormatCurrency(dPermitFee,2)
		setPDFFormFieldData oPDF, "permitfees", dPermitFee, 1
	End If 

	' Payment rows here
	GetPaymentForPermit iPermitId, sPaymentDate1, sMethod1, sAmount1, sPaymentDate2, sMethod2, sAmount2
	setPDFFormFieldData oPDF, "paymentdate1", sPaymentDate1, 1
	setPDFFormFieldData oPDF, "method1", sMethod1, 1
	setPDFFormFieldData oPDF, "amount1", sAmount1, 1
	setPDFFormFieldData oPDF, "paymentdate2", sPaymentDate2, 1
	setPDFFormFieldData oPDF, "method2", sMethod2, 1
	setPDFFormFieldData oPDF, "amount2", sAmount2, 1

	' Get the total paid
	'sTotalPaid = GetPermitDetailItemAsNumber( iPermitId, "totalpaid", "currency" )
	sTotalPaid = GetPaidTotal( iPermitId )
	setPDFFormFieldData oPDF, "totalpaid", sTotalPaid, 1

	sDimensions = FormatNumber(GetPermitDetailItemAsNumber( iPermitId, "totalsqft", "integer" ),0) & " sq.ft."
	setPDFFormFieldData oPDF, "dimensions", sDimensions, 1

	sJobValue = GetPermitDetailItemAsNumber( iPermitId, "jobvalue", "currency" )
	setPDFFormFieldData oPDF, "jobvalue", sJobValue, 1

	sDescriptionOfWork = GetPermitDetailItemAsString( iPermitId, "descriptionofwork" )
	setPDFFormFieldData oPDF, "descriptionofwork", sDescriptionOfWork, 1

	For x = 1 To 4 
		GetPermitAttachments iPermitid, x, sAttachmentDate, sAttachmentName, sAttachmentDescription
		setPDFFormFieldData oPDF, "attachmentdate" & x, sAttachmentDate, 1
		setPDFFormFieldData oPDF, "attachmentname" & x, sAttachmentName, 1
		setPDFFormFieldData oPDF, "attachmentdescription" & x, sAttachmentDescription, 1
	Next 

	GetPermitInspectionsList iPermitId, sInspectionsDue, sInspectorList 
	ReplaceBreaksForPDFs sInspectionsDue
	setPDFFormFieldData oPDF, "inspectionsdue", sInspectionsDue, 1
	ReplaceBreaksForPDFs sInspectorList
	setPDFFormFieldData oPDF, "inspectorlist", sInspectorList, 1

	GetPermitReviewList iPermitid, sReviewList, sReviewStatus, sReviewDate, sReviewer, sReviewNotes 
	ReplaceBreaksForPDFs sReviewList
	setPDFFormFieldData oPDF, "reviewlist", sReviewList, 1

	' Add these for the old way of doint it.
	sReviewStatus = sReviewStatus & "<br />" & sReviewNotes
	ReplaceBreaksForPDFs sReviewStatus
	setPDFFormFieldData oPDF, "reviewstatus", sReviewStatus, 1

'	ReplaceBreaksForPDFs sReviewNotes
'	setPDFFormFieldData oPDF, "reviewnotes", sReviewNotes, 1
	'If InStr(sPDFPath, "GeneralPermitE-GovPermit") > 0 Then
'		Set oField = oPDF.FieldInfo("reviewnotes",1)
'		iWidth = oField.Width
'		oField.Width = (iWidth - 100)
		'setPDFFormFieldData oPDF, "reviewnotes", oField.Width, 1
		'oField.Value = "Something Here"
'		Set oField = Nothing 
	'End If 

	ReplaceBreaksForPDFs sReviewDate
	setPDFFormFieldData oPDF, "reviewdate", sReviewDate, 1
	ReplaceBreaksForPDFs sReviewer
	setPDFFormFieldData oPDF, "reviewer", sReviewer, 1

	GetPermitInspectionsPDF iPermitid, sInspectionList, sInspectionStatus, sInspectionDate, sInspector, sInspectionNotes
	ReplaceBreaksForPDFs sInspectionList
	setPDFFormFieldData oPDF, "inspectionlist", sInspectionList, 1
	ReplaceBreaksForPDFs sInspectionStatus
	setPDFFormFieldData oPDF, "inspectionstatus", sInspectionStatus, 1
	ReplaceBreaksForPDFs sInspectionDate
	setPDFFormFieldData oPDF, "inspectiondate", sInspectionDate, 1
	ReplaceBreaksForPDFs sInspector
	setPDFFormFieldData oPDF, "inspector", sInspector, 1
	ReplaceBreaksForPDFs sInspectionNotes
	setPDFFormFieldData oPDF, "inspectionnotes", sInspectionNotes, 1

  	oPDF.FlattenRemainingFormFields = True 
  	r = oPDF.CopyForm(0, 0)

 	'Close PDF
  	oPDF.CloseOutputFile
  	oDocument = oPDF.binaryImage 

 	'Stream PDF to browser
  	response.expires = 0
  	response.Clear
  	response.ContentType = "application/pdf"
  	response.AddHeader "Content-Type", "application/pdf"
  	response.AddHeader "Content-Disposition", "inline;filename=Permit.pdf"
  	response.BinaryWrite oDocument  

 	'Destory Objects
  	Set oPDF      = Nothing 
  	Set oDocument = Nothing 

 End Sub 


'--------------------------------------------------------------------------------------------------
' Sub setPDFFormFieldData( oPDF, sFieldName, sFieldValue, iReadOnly )
'--------------------------------------------------------------------------------------------------
Sub setPDFFormFieldData( oPDF, sFieldName, sFieldValue, iReadOnly )
	Dim r
	
	'Object Properties: object.SetFormFieldData "FieldName", "FieldData", LeaveReadOnlyFlag
	r = oPDF.SetFormFieldData( sFieldName, sFieldValue, iReadOnly )

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetPermitDocumentInfo( iPermitid, iPermitDocumentId, sPDFPath )
'--------------------------------------------------------------------------------------------------
Function GetPermitDocumentInfo( ByVal iPermitid, ByVal iPermitDocumentId )
	Dim sSql, oRs 

	sSql = "SELECT D.document, D.path FROM egov_permitdocuments D, egov_permittypes_to_permitdocuments PD, egov_permits P "
	sSql = sSql & " WHERE PD.permittypeid = P.permittypeid AND PD.documentid = D.documentid AND P.permitid = " & iPermitid
	sSql = sSql & " AND PD.permitdocumentid = " & iPermitDocumentId

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
' void ReplaceBreaksForPDFs( ByRef sFieldValue )
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
' Sub GetPaymentForPermit( ByVal iPermitId, ByRef sPaymentDate1, ByRef sMethod1, ByRef sAmount1, ByRef sPaymentDate2, ByRef sMethod2, ByRef sAmount2 )
'--------------------------------------------------------------------------------------------------
Sub GetPaymentForPermit( ByVal iPermitId, ByRef sPaymentDate1, ByRef sMethod1, ByRef sAmount1, ByRef sPaymentDate2, ByRef sMethod2, ByRef sAmount2 )
	Dim sSql, oRs, iCount

	iCount = clng(0) 

'	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
'	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_permitinvoices I "
'	sSql = sSql & " WHERE L.paymentid = J.paymentid AND L.ispaymentaccount = 1 AND L.permitid = " & iPermitId
'	sSql = sSql & " AND J.paymentid = I.paymentid AND I.isvoided = 0 GROUP BY L.paymentid, J.paymentdate"

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



%>
