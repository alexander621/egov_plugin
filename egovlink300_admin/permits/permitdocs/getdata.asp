<%

		sSQL = "SELECT pt.permittype + CASE WHEN pt.permittypedesc IS NULL THEN '' ELSE ' - ' + pt.permittypedesc END as permittypedesc, " _
			& " ps.permitstatus, approvedas, " _
			& " CASE ps.statusdatedisplayed " _
			& " WHEN 'applieddate' THEN p.applieddate " _
			& " WHEN 'releaseddate' THEN p.releaseddate " _
			& " WHEN 'approveddate' THEN p.approveddate " _
			& " WHEN 'issueddate' THEN p.issueddate " _
			& " WHEN 'completeddate' THEN p.completeddate " _
			& " END as statusdate, p.issueddate, p.completeddate, p.applieddate, p.expirationdate, pa.residentstreetnumber, pa.residentstreetprefix, pa.residentstreetname, " _
			& " pa.streetsuffix, pa.streetdirection, pa.residentunit, pa.residentcity, pa.residentstate, pa.legaldescription, pa.listedowner,  " _
			& " pa.residentzip, pa.county, pa.parcelidnumber,  " _
			& "  " _
			& " pcapp.firstname as appfirstname, pcapp.lastname as applastname, pcapp.company as appcompany, pcapp.address as appaddress, pcapp.city as appcity,  " _
			& " pcapp.state as appstate, pcapp.zip as appzip, pcapp.phone as appphone, pcapp.contacttype as appcontacttype, pcapp.email as appemail, " _
			& " pcapp.cell as appcell, pcapp.fax as appfax, pccon.isprimarycontractor, " _
			& "  " _
			& " pccon.firstname as confirstname, pccon.lastname as conlastname, pccon.company as concompany, pccon.address as conaddress, pccon.city as concity,  " _
			& " pccon.state as constate, pccon.zip as conzip, pccon.phone as conphone, pccon.contacttype as concontacttype, pccon.email as conemail, " _
			& " pccon.cell as concell, pccon.fax as confax, " _
			& " p.descriptionofwork, ct.constructiontype, ot.usegroupcode,ot.occupancytype, p.existinguse, p.proposeduse, p.zoning, p.residentialunits, p.occupants,  " _
			& " IsNull(p.totalsqft,0) as totalsqft, p.conotes, p.tempconotes, put.usetype, puc.useclass, pws.workscope, pwc.workclass, ISNULL(p.feetotal,0) as feetotal, " _
			& " ISNULL(p.jobvalue,0) as jobvalue, const.constructiontype, structurelength, structurewidth, structureheight, isnull(permitnotes,'<br /><br />') as permitnotes, " _
			& " p.primarycontact " _
			& " FROM egov_permits p " _
			& " INNER JOIN egov_permitpermittypes pt ON pt.permitid = p.permitid " _
			& " INNER JOIN egov_permitstatuses ps ON ps.permitstatusid = p.permitstatusid " _
			& " INNER JOIN egov_permitaddress pa ON pa.permitid = p.permitid " _
			& " INNER JOIN egov_permitcontacts pcapp ON pcapp.permitid = p.permitid AND pcapp.isapplicant = 1 and pcapp.ispriorcontact = 0 " _
			& " LEFT JOIN egov_permitcontacts pccon ON pccon.permitid = p.permitid AND pccon.isprimarycontractor = 1 and pcapp.ispriorcontact = 0 " _
			& " LEFT JOIN egov_constructiontypes ct ON ct.constructiontypeid = p.constructiontypeid " _
			& " LEFT JOIN egov_occupancytypes ot ON p.occupancytypeid = ot.occupancytypeid " _
			& " LEFT JOIN egov_permitusetypes put ON put.usetypeid = p.usetypeid " _
			& " LEFT JOIN egov_permituseclasses puc ON puc.useclassid = p.useclassid " _
			& " LEFT JOIN egov_permitworkscope pws ON pws.workscopeid = p.workscopeid " _
			& " LEFT JOIN egov_permitworkclasses pwc ON pwc.workclassid = p.workclassid " _
			& " LEFT JOIN egov_constructiontypes const ON const.constructiontypeid = p.constructiontypeid " _
			& " WHERE p.permitid = " & intPermitID
		Set oRs = Server.CreateObject("ADODB.RecordSet")
		oRs.Open sSQL, Application("DSN"), 3, 1

		strPermitNumber = GetPermitNumber(intPermitID)
		strPermitIssuedBy = GetPermitIssuedBy(intPermitID)

		sSql = "SELECT TOP 4 dateadded, attachmentname, description FROM egov_permitattachments WHERE permitid = " & intPermitID  & " ORDER BY dateadded, attachmentname"
		Set oRsA = Server.CreateObject("ADODB.RecordSet")
		oRsA.Open sSQL, Application("DSN"), 3, 1

		sSQL = "SELECT ISNULL(I.permitinspectiontype,'') + ' - ' + ISNULL(I.inspectiondescription,'') as inspection, S.inspectionstatus, " _
			& " I.inspecteddate, u.FirstName + ' ' + u.LastName as inspector, " _
			& " ISNULL(u.BusinessNumber,'') as inspectorphone, i.permitinspectionid " _
			& " FROM egov_permitinspections I " _
			& " INNER JOIN  egov_inspectionstatuses S ON I.inspectionstatusid = S.inspectionstatusid " _
			& " LEFT JOIN Users u ON u.UserID = i.inspectoruserid " _
			& " WHERE  I.permitid = " & intPermitID & " " _
			& " ORDER BY I.inspectionorder "
		Set oRsI = Server.CreateObject("ADODB.RecordSet")
		oRsI.Open sSQL, Application("DSN"), 3, 1

		sSql = "SELECT TOP 6 R.permitreviewid, R.permitreviewtype, S.reviewstatus, S.shownotes, "
		sSql = sSql & " ISNULL(R.revieweruserid,0) AS revieweruserid, R.reviewed, " _
			& "  u.FirstName + ' ' + u.LastName as reviewer, " _
			& " ISNULL(u.BusinessNumber,'') as reviewerphone " 
		sSql = sSql & " FROM egov_permitreviews R "
		sSql = sSql & " INNER JOIN egov_reviewstatuses S ON R.reviewstatusid = S.reviewstatusid "
		sSql = sSql & " LEFT JOIN Users u ON u.UserID = R.revieweruserid "
		sSql = sSql & " WHERE R.permitid = " & intPermitID
		sSql = sSql & " ORDER BY R.reviewed"
		Set oRsPR = Server.CreateObject("ADODB.RecordSet")
		oRsPR.Open sSQL, Application("DSN"), 3, 1


		sReviewNotes = ""
		if not oRsPR.EOF then
			Do While Not oRsPR.EOF
				sLastNote = GetPermitReviewNotes( oRsPR("permitreviewid") )
				If sLastNote <> "" Then 
					sReviewNotes = sReviewNotes & "<br /><br />" & oRsPR("permitreviewtype") & " Notes and Conditions:<br />" & sLastNote & " <br />"
				End If 
				oRsPR.MoveNext
			loop
			oRsPR.MoveFirst
		end if

Function GetPrimaryContactLicense( ByVal iPermitId )
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(L.licensenumber,'') AS licensenumber "
	sSql = sSql & "FROM egov_permitcontacts P, egov_permitcontacts_licenses L, egov_permitlicensetypes T "
	sSql = sSql & "WHERE P.permitid = L.permitid AND P.permitcontactid = L.permitcontactid AND L.licensetypeid = T.licensetypeid "
	sSql = sSql & "AND licenseenddate >= getdate() AND P.permitid = " & iPermitId
	sSql = sSql & " AND P.isprimarycontractor = 1"

	Set oRsPCL = Server.CreateObject("ADODB.Recordset")
	oRsPCL.Open sSql, Application("DSN"), 3, 1

	If Not oRsPCL.EOF Then 
		GetPrimaryContactLicense = oRsPCL("licensenumber")
	Else
		GetPrimaryContactLicense = ""
	End If
	
	oRsPCL.Close
	Set oRsPCL = Nothing 

End Function 

	strUdonumber = ""
	strFloodzone = ""
	strFrontsetback = ""
	strRearsetback = ""
	strSidesetback = ""
	strPropertysize = ""
	strLotnumber = ""
	strOwnerName = ""
	strMaxOcc = ""
	strPeriodOfTime = ""
	strPortionOfBuilding = ""

	sSql = "SELECT P.customfieldid, P.pdffieldname, F.fieldtypebehavior, P.valuelist, P.fieldsize, "
	sSql = sSql & "ISNULL(P.simpletextvalue,'') AS simpletextvalue, ISNULL(P.largetextvalue,'') AS largetextvalue, "
	sSql = sSql & "P.datevalue, moneyvalue, P.intvalue "
	sSql = sSql & "FROM egov_permitcustomfields P, egov_permitfieldtypes F "
	sSql = sSql & "WHERE P.fieldtypeid = F.fieldtypeid AND P.pdffieldname IS NOT NULL AND P.permitid = " & intPermitID
	sSql = sSql & "ORDER BY P.displayorder"

	Set oRsCF = Server.CreateObject("ADODB.Recordset")
	oRsCF.Open sSql, Application("DSN"), 3, 1
	Do While not oRsCF.EOF
		Select Case oRsCF("pdffieldname")
			case "udonumber"
				strUdonumber = oRsCF("simpletextvalue")
			case "floodzone"
				strFloodzone = oRsCF("simpletextvalue")
			case "frontsetback"
				strFrontsetback = oRsCF("simpletextvalue")
			case "rearsetback"
				strRearsetback = oRsCF("simpletextvalue")
			case "sidesetback"
				strSidesetback = oRsCF("simpletextvalue")
			case "propertysize"
				strPropertysize = oRsCF("simpletextvalue")
			case "lotnumber"
				strLotnumber = oRsCF("simpletextvalue")
			case "plansapproved"
				strPlansApproved = oRsCF("simpletextvalue")
			case "plansapprovedother"
				strPlansApprovedOther = oRsCF("simpletextvalue")
			case "primaryaddress"
				strPrimaryAddress = replace(oRsCF("largetextvalue"),vbcrlf,"<br />")
			case "owneraddress"
				strOwnerAddress = replace(oRsCF("largetextvalue") & oRsCF("simpletextvalue"),vbcrlf,"<br />")
			case "permitconditions"
				strPermitConditions = replace(oRsCF("largetextvalue"),vbcrlf,"<br />")
			case "ownername"
				strOwnerName = oRsCF("simpletextvalue")
			case "ownerphone"
				strOwnerPhone = oRsCF("simpletextvalue")
			case "owneremail"
				strOwnerEmail = oRsCF("simpletextvalue")
			case "automaticsprinklers"
				strAutoSprinklers = oRsCF("simpletextvalue")
			case "hazardclassification"
				strHazard = oRsCF("simpletextvalue")
			case "zaddcomments"
				strZAddComments = oRsCF("simpletextvalue")
			case "bldqcode"
				strBLDQCode = oRsCF("simpletextvalue")
			case "descocc"
				strDESCOCC = oRsCF("simpletextvalue")
			case "maxocc"
				strMaxOcc = oRsCF("simpletextvalue")
			case "periodoftime"
				strPeriodOfTime = oRsCF("simpletextvalue")
			case "portionofbuilding"
				strPortionOfBuilding = oRsCF("simpletextvalue")

		end select
		oRsCF.MoveNext
	loop
	oRsCF.Close
	Set oRsCF = Nothing

	blnApprovedOhio = false
	if instr(strPlansApproved,"Ohio") > 0 then blnApprovedOhio = true

	blnApprovedWyoming = false
	if instr(strPlansApproved,"Wyoming") > 0 then blnApprovedWyoming = true

	'Get Contractors
	strEleConName = ""
	strEleConLic = ""
	strMechConName = ""
	strMechConLic = ""
	strPlumbConName = ""
	strPlumbConLic = ""
	strInsuConName = ""
	strInsuConLic = ""
	sSQL = "SELECT DISTINCT p.permitid, l.licensenumber, t.contractortype, " _
		& " CASE WHEN p.company <> '' and p.company IS NOT NULL THEN p.company ELSE ISNULL(P.firstname,'') + ' ' + ISNULL(P.lastname,'') END as Name " _
		& " FROM egov_permitcontacts P " _
		& " INNER JOIN egov_permitcontractortypes t ON t.contractortypeid = p.contractortypeid " _
		& " LEFT JOIN egov_permitcontacts_licenses L ON l.permitcontactid = p.permitcontactid and licenseenddate >= GETDATE() " _
		& " WHERE  ispriorcontact = 0 AND contractortype <> 'General' AND P.permitid = " & intPermitID
	Set oRsCon = Server.CreateObject("ADODB.Recordset")
	oRsCon.Open sSql, Application("DSN"), 3, 1
	Do While not oRsCon.EOF
		Select Case oRsCon("contractortype")
			case "Electrical"
				strEleConName = oRsCon("Name")
				strEleConLic = oRsCon("licensenumber")
			case "Mechanical"
				strMechConName = oRsCon("Name")
				strMechConLic = oRsCon("licensenumber")
			case "Plumbing"
				strPlumbConName = oRsCon("Name")
				strPlumbConLic = oRsCon("licensenumber")
			case "Insulation"
				strInsuConName = oRsCon("Name")
				strInsuConLic = oRsCon("licensenumber")
		end select
		oRsCon.MoveNext
	loop
	oRsCon.Close
	Set oRsCon = Nothing

	dZoneFeeAmount = CDbl(0.00)

	' See if the city has a bbs fee - That makes it Loveland and has all the fee things below
	If OrgHasFeeType( "isbbs" ) Then 

		dFeesSum = CDbl(0.00)

		dZoneFee = GetPermitFeeTypeTotal( intPermitId, "iszone" )
		If CDbl(dZoneFee) = CDbl(0.00) Then
			dZoneFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dZoneFee)
			dZoneFee = FormatCurrency(dZoneFee,2)
			dZoneFeeAmount = CDbl(dZoneFee)
		End If 
		response.write "<!--ZONEFEE:" & dZoneFee & "-->" & vbcrlf

		dBBSFee = GetPermitFeeTypeTotal( intPermitId, "isbbs" )
		If CDbl(dBBSFee) = CDbl(0.00) Then
			dBBSFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dBBSFee)
			dBBSFee = FormatCurrency(dBBSFee,2)
		End If 
		response.write "<!--BBSFEE:" & dBBSFee & "-->" & vbcrlf

		dRecImpactFee = GetPermitFeeTypeTotal( intPermitId, "isrecreationimpact" )
		If CDbl(dRecImpactFee) = CDbl(0.00) Then
			dRecImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dRecImpactFee)
			dRecImpactFee = FormatCurrency(dRecImpactFee,2)
		End If 
		response.write "<!--RECIMPACTFEE:" & dRecImpactFee & "-->" & vbcrlf

		dWaterImpactFee = GetPermitFeeTypeTotal( intPermitId, "iswaterimpact" )
		If CDbl(dWaterImpactFee) = CDbl(0.00) Then
			dWaterImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dWaterImpactFee)
			dWaterImpactFee = FormatCurrency(dWaterImpactFee,2)
		End If 
		response.write "<!--WaterImpactFEE:" & dWaterImpactFee & "-->" & vbcrlf

		dWaterMeterFee = GetPermitFeeTypeTotal( intPermitId, "iswatermeter" )
		If CDbl(dWaterMeterFee) = CDbl(0.00) Then
			dWaterMeterFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dWaterMeterFee)
			dWaterMeterFee = FormatCurrency(dWaterMeterFee,2)
		End If 
		response.write "<!--WaterMeterFEE:" & dWaterMeterFee & "-->" & vbcrlf

		dRoadImpactFee = GetPermitFeeTypeTotal( intPermitId, "isroadimpact" )
		If CDbl(dRoadImpactFee) = CDbl(0.00) Then
			dRoadImpactFee = ""
		Else
			dFeesSum = dFeesSum + CDbl(dRoadImpactFee)
			dRoadImpactFee = FormatCurrency(dRoadImpactFee,2)
		End If 
		response.write "<!--RoadImpactFEE:" & dRoadImpactFee & "-->" & vbcrlf

		dCofOFee = GetPermitFeeTypeTotal( intPermitId, "iscertofoccupancy" )
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
		response.write "<!--CofOFEE:" & dCofOFee & "-->" & vbcrlf

		dPermitFee = CDbl(dTotalFees) - dFeesSum
		dPermitFee = FormatCurrency(dPermitFee,2)
	End If 
	
	dTotalFees = GetPermitDetailItemAsNumber( intPermitId, "feetotal", "double" )
	sTotalFees = FormatCurrency(dTotalFees,2)
	response.write "<!--TOTALFEES" & dTotalFees & "-->"
	response.write "<!--FEESSUM" & dFeesSum & "-->"

	dPermitFee = CDbl(dTotalFees) - dFeesSum
	dPermitFee = FormatCurrency(dPermitFee,2)

	dNoZoneTotalFees = CDbl(dTotalFees) - dZoneFeeAmount
	dNoZoneTotalFees = FormatCurrency(dNoZoneTotalFees,2)

	dTotalFees = FormatCurrency(dTotalFees,2)
	
		'-Plans By
		sPlansBy = GetPermitPlansBy( intPermitid )
	
		'-totalpaid
		sTotalPaid = GetPaidTotal( intPermitid )

		
		'-Payment Info
		GetPaymentForPermit intPermitid, sPaymentDate1, sMethod1, sAmount1, sPaymentDate2, sMethod2, sAmount2



		sDimensions = FormatNumber(GetPermitDetailItemAsNumber( intPermitID, "totalsqft", "integer" ),0) & " sq.ft."

		strFinalInspectionDate = GetPermitFinalInspectionDate( intPermitId )

		'NEED
		'-certofoccupancyfee
		'-totalnozonefees
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
%>
