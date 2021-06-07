<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitupdate.asp
' AUTHOR: Steve Loar
' CREATED: 03/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates permits
'
' MODIFICATION HISTORY
' 1.0   03/14/2008	Steve Loar - INITIAL VERSION
' 1.1	01/18/2010	Steve Loar - Added notification of applicant of inspection scheduling
' 1.2	10/27/2010	Steve Loar - Changes to allow any type of permits
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitid, iUseTypeId, sSql, sDescriptionofwork, iWorkclassid, iConstructiontypeid, iOccupancytypeid
Dim iConstructiontyperate, sExistinguse, sProposeduse, iContactTypeId, sCompany, sFirstname
Dim sLastname, sAddress, sCity, sState, sZip, sEmail, sPhone, sCell, sFax, iContactUserId, iPrimaryContactUserId
Dim sJobValue, sTotalSqFt, sFinishedSqFt, sUnFinishedSqFt, sOtherSqFt, sHours, iWaiveAllFees
Dim sInternalNotes, sPublicNotes, iPermitStatusId, sPermitNotes, sResidentialUnits
Dim sPassword, sWorkPhone, sEmergencyContact, sEmergencyPhone, iNeighborhoodid, sResidentType, sBusinessAddress
Dim sUserUnit, sEmailnotavailable, sResidencyVerified, sSuccessMsg, iContractorTypeId, iPlansByContactId
Dim sApprovedAs, sOccupants, sTempCONotes, sCONotes, sPrimaryContact, iUseClassId, iWorkScopeId
Dim sStructureLength, sStructureWidth, sStructureHeight, sZoning, sPlanNumber, sAlertApplicant
Dim sDemolishExistingStructure, sLandFillName, sLandFillCity, sLandFillPhone, iBusinessTypeId
Dim sStateLicense, sEmployeeCount, sReference1, sAutoInsurancePhone, sBondAgent, sBondAgentPhone
Dim sReference2, sReference3, sOtherLicensedCity1, sOtherLicensedCity2, sGeneralLiabilityAgent
Dim sGeneralLiabilityPhone, sWorkersCompAgent, sWorkersCompPhone, sAutoInsuranceAgent, sPermitLocation 
Dim sNewValue, Item

iPermitid = CLng(request("permitid"))

'response.write "iPermitid = " & iPermitid & "<br />"

sCompany = "NULL"
sFirstname = "NULL"
sLastname = "NULL"
sAddress = "NULL"
sCity = "NULL"
sState = "NULL"
sZip = "NULL"
sEmail = "NULL"
sPhone = "NULL"
sCell = "NULL"
sFax = "NULL"
sPassword = "NULL"
sWorkPhone = "NULL"
sEmergencyContact = "NULL"
sEmergencyPhone = "NULL"
iNeighborhoodid = "NULL"
sResidentType = "NULL"
sBusinessAddress = "NULL"
sUserUnit = "NULL"
sEmailnotavailable = 0
sResidencyVerified = 0
sSuccessMsg = "Changes Saved"

If request("permitlocation") <> "" Then
	sPermitLocation = "'" & dbsafe(request("permitlocation")) & "'"
Else
	sPermitLocation = "NULL"
End If 

If CLng(request("usetypeid")) > CLng(0) Then
	iUseTypeId = CLng(request("usetypeid"))
Else
	iUseTypeId = "NULL"
End If 

If CLng(request("useclassid")) > CLng(0) Then
	iUseClassId = CLng(request("useclassid"))
Else
	iUseClassId = "NULL"
End If

If CLng(request("workscopeid")) > CLng(0) Then
	iWorkScopeId = CLng(request("workscopeid"))
Else
	iWorkScopeId = "NULL"
End If

If request("descriptionofwork") <> "" Then
	sDescriptionofwork = "'" & dbsafe(request("descriptionofwork")) & "'"
Else
	sDescriptionofwork = "NULL"
End If 

If CLng(request("workclassid")) > CLng(0) Then
	iWorkclassid = CLng(request("workclassid"))
Else
	iWorkclassid = "NULL"
End If 

If CLng(request("constructiontypeid")) > CLng(0) Then
	iConstructiontypeid = CLng(request("constructiontypeid"))
Else
	iConstructiontypeid = "NULL"
End If 

If CLng(request("occupancytypeid")) > CLng(0) Then
	iOccupancytypeid = CLng(request("occupancytypeid"))
Else
	iOccupancytypeid = "NULL"
End If 

If CLng(request("constructiontypeid")) > CLng(0) And CLng(request("occupancytypeid")) > CLng(0) Then 
	iConstructiontyperate = GetConstructionRate( iConstructionTypeId, iOccupancyTypeId )
Else
	iConstructiontyperate = "NULL"
End If 

If request("existinguse") <> "" Then
	sExistinguse = "'" & dbsafe(request("existinguse")) & "'"
Else
	sExistinguse = "NULL"
End If 

If request("proposeduse") <> "" Then
	sProposeduse = "'" & dbsafe(request("proposeduse")) & "'"
Else
	sProposeduse = "NULL"
End If 

If request("jobvalue") <> "" Then 
	sJobValue = FormatNumber(request("jobvalue"),2,,,0)
Else
	sJobValue = 0.00
End If 

If request("othersqft") <> "" Then 
	sOtherSqFt = FormatNumber(request("othersqft"),2,,,0)
Else
	sOtherSqFt = 0.00
End If 

If request("finishedsqft") <> "" Then 
	sFinishedSqFt = FormatNumber(request("finishedsqft"),2,,,0)
Else
	sFinishedSqFt = 0.00
End If 

If request("unfinishedsqft") <> "" Then 
	sUnFinishedSqFt = FormatNumber(request("unfinishedsqft"),2,,,0)
Else
	sUnFinishedSqFt = 0.00
End If 

sTotalSqFt = CDbl(sFinishedSqFt) + CDbl(sUnFinishedSqFt)

If request("examinationhours") <> "" Then
	sHours = FormatNumber(request("examinationhours"),2,,,0)
Else
	sHours = 0.00
End If 

If request("waiveallfees") = "on" Then
	iWaiveAllFees = 1
Else
	iWaiveAllFees = 0
End If 

If request("permitnotes") = "" Then
	sPermitNotes = "NULL"
Else
	'response.write "Permit Notes: " & request("permitnotes") & "<br />"
	sPermitNotes = "'" & dbsafe(request("permitnotes")) & "'"
End If 

If CLng(request("plansbycontactid")) > CLng(0) Then
	iPlansByContactId = CLng(request("plansbycontactid"))
Else
	iPlansByContactId = "NULL"
End If 

If request("residentialunits") <> "" Then
	sResidentialUnits = CLng(request("residentialunits"))
Else
	sResidentialUnits = CLng(0)
End If 

If request("approvedas") <> "" Then
	sApprovedAs = "'" & dbsafe(request("approvedas")) & "'"
Else
	sApprovedAs = "NULL"
End If 

If request("occupants") <> "" Then
	sOccupants = CLng(request("occupants"))
Else
	sOccupants = "NULL"
End If 

If request("tempconotes") <> "" Then
	sTempCONotes = "'" & dbsafe(request("tempconotes")) & "'"
Else
	sTempCONotes = "NULL"
End If 

If request("conotes") <> "" Then
	sCONotes = "'" & dbsafe(request("conotes")) & "'"
Else
	sCONotes = "NULL"
End If 

If request("primarycontact") <> "" Then
	sPrimaryContact = "'" & dbsafe(request("primarycontact")) & "'"
Else
	sPrimaryContact = "NULL"
End If 

If request("structurelength") <> "" Then
	sStructureLength = "'" & dbsafe(request("structurelength")) & "'"
Else
	sStructureLength = "NULL"
End If 

If request("structurewidth") <> "" Then
	sStructureWidth = "'" & dbsafe(request("structurewidth")) & "'"
Else
	sStructureWidth = "NULL"
End If 

If request("structureheight") <> "" Then
	sStructureHeight = "'" & dbsafe(request("structureheight")) & "'"
Else
	sStructureHeight = "NULL"
End If 

If request("zoning") <> "" Then
	sZoning = "'" & dbsafe(request("zoning")) & "'"
Else
	sZoning = "NULL"
End If 

If request("plannumber") <> "" Then
	sPlanNumber = "'" & dbsafe(request("plannumber")) & "'"
Else
	sPlanNumber = "NULL"
End If 

If request("demolishexistingstructure") = "on" Then
	sDemolishExistingStructure = 1
Else
	sDemolishExistingStructure = 0
End If 

If request("landfillname") <> "" Then
	sLandFillName = "'" & dbsafe(request("landfillname")) & "'"
Else
	sLandFillName = "NULL"
End If 

If request("landfillcity") <> "" Then
	sLandFillCity = "'" & dbsafe(request("landfillcity")) & "'"
Else
	sLandFillCity = "NULL"
End If 

If request("landfillphone") <> "" Then
	sLandFillPhone = "'" & dbsafe(request("landfillphone")) & "'"
Else
	sLandFillPhone = "NULL"
End If 

If request("alertapplicantofinspections") = "on" Then
	sAlertApplicant = "1"
Else
	sAlertApplicant = "0"
End If 


sSql = "UPDATE egov_permits SET usetypeid = " & iUseTypeId 
sSql = sSql & ", descriptionofwork = " & sDescriptionofwork
sSql = sSql & ", workclassid = " & iWorkclassid 
sSql = sSql & ", workscopeid = " & iWorkScopeid
sSql = sSql & ", useclassid = " & iUseClassId 
sSql = sSql & ", constructiontypeid = " & iConstructiontypeid
sSql = sSql & ", occupancytypeid = " & iOccupancytypeid
sSql = sSql & ", constructiontyperate = " & iConstructiontyperate
sSql = sSql & ", existinguse = " & sExistinguse
sSql = sSql & ", proposeduse = " & sProposeduse
sSql = sSql & ", jobvalue = " & sJobValue
sSql = sSql & ", totalsqft = " & sTotalSqFt
sSql = sSql & ", finishedsqft = " & sFinishedSqFt
sSql = sSql & ", unfinishedsqft = " & sUnFinishedSqFt
sSql = sSql & ", othersqft = " & sOtherSqFt
sSql = sSql & ", examinationhours = " & sHours
sSql = sSql & ", waiveallfees = " & iWaiveAllFees
sSql = sSql & ", permitnotes = " & sPermitNotes
sSql = sSql & ", plansbycontactid = " & iPlansByContactId
sSql = sSql & ", residentialunits = " & sResidentialUnits
sSql = sSql & ", approvedas = " & sApprovedAs
sSql = sSql & ", occupants = " & sOccupants
sSql = sSql & ", tempconotes = " & sTempCONotes
sSql = sSql & ", conotes = " & sCONotes
sSql = sSql & ", primarycontact = " & sPrimaryContact
sSql = sSql & ", structurelength = " & sStructureLength
sSql = sSql & ", structurewidth = " & sStructureWidth
sSql = sSql & ", structureheight = " & sStructureHeight
sSql = sSql & ", zoning = " & sZoning
sSql = sSql & ", plannumber = " & sPlanNumber
sSql = sSql & ", demolishexistingstructure = " & sDemolishExistingStructure
sSql = sSql & ", landfillname = " & sLandFillName
sSql = sSql & ", landfillcity = " & sLandFillCity
sSql = sSql & ", landfillphone = " & sLandFillPhone
sSql = sSql & ", alertapplicantofinspections = " & sAlertApplicant
sSql = sSql & ", permitlocation = " & sPermitLocation
sSql = sSql & " WHERE permitid = " & iPermitid
RunSQL sSql


' Check the Primary Contact for changes
'If CLng(request("isprimarycontactoriginaluserid")) <> CLng(request("isprimarycontactuserid")) Then 
'	iPrimaryContactUserId = CLng(request("isprimarycontactuserid"))
'
'	' fetch the info for the new contact
'	GetPrimaryContactInfo iPrimaryContactUserId
'
'	If iPrimaryContactUserId = CLng(0) Then
'		iPrimaryContactUserId = "NULL"
'	End If 
'
'	RemoveContact request("isprimarycontactpermitcontactid")
'	CreateNewContactRecord "NULL", "isprimarycontact", iPermitId, iPrimaryContactUserId
'
'End If 

' Check the Billing Contact for changes
If CLng(request("isbillingcontactoriginalpermitcontacttypeid")) <> CLng(request("isbillingcontactpermitcontacttypeid")) Then 
	iContactTypeId = CLng(request("isbillingcontactpermitcontacttypeid"))  ' The new one's type id
	iOldPermitContactId = CLng(request("isbillingcontactpermitcontactid")) ' The old one's id

	' Remove the old one from the permit contacts and license table
	RemoveContact iOldPermitContactId

	If iContactTypeId <> CLng(0) Then
		' Create new contact
		CreateContactInfo iContactTypeId, "isbillingcontact", iPermitId
	End If 
End If 

' Check the Primary Contractor for changes
If CLng(request("isprimarycontractororiginalpermitcontacttypeid")) <> CLng(request("isprimarycontractorpermitcontacttypeid")) Then 
	iContactTypeId = CLng(request("isprimarycontractorpermitcontacttypeid")) ' The new one's type id
	iOldPermitContactId = CLng(request("isprimarycontractorpermitcontactid")) ' The old one's id

	' Remove the old one from the permit contacts and license table
	RemoveContact iOldPermitContactId

	If iContactTypeId <> CLng(0) Then
		' Create new contact
		CreateContactInfo iContactTypeId, "isprimarycontractor", iPermitId
	End If 
End If 

' Check the Architect/Engineer for changes
If CLng(request("isarchitectoriginalpermitcontacttypeid")) <> CLng(request("isarchitectpermitcontacttypeid")) Then 
	iContactTypeId = CLng(request("isarchitectpermitcontacttypeid")) ' The new one's type id
	iOldPermitContactId = CLng(request("isarchitectpermitcontactid")) ' The old one's id

	' Remove the old one from the permit contacts and license table
	RemoveContact iOldPermitContactId

	If iContactTypeId <> CLng(0) Then
		' Create new contact
		CreateContactInfo iContactTypeId, "isarchitect", iPermitId
	End If 
End If 

' Handle the Contractors now
For x = 0 To CLng(request("maxcontractors"))
	' Add any new ones 
	If request("permitcontactid" & x) = "0" Then 
		iContactTypeId = CLng(request("contractor" & x))
		CreateContactInfo iContactTypeId, "iscontractor", iPermitId
	End If 
Next 

' Handle the custom fields
If CLng(request("maxcustompermitfields")) > CLng(0) Then
	For x = 1 To CLng(request("maxcustompermitfields"))
		
		sSql = "UPDATE egov_permitcustomfields "

		Select Case request("fieldtypebehavior" & x)
			Case "date"
				If request("customfield" & x) = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = "'" & request("customfield" & x) & "'"
				End If 
				sSql = sSql & "SET datevalue = " & sNewValue & " " 

			Case "radio"
				' There will always be at least one radio picked
				sSql = sSql & "SET simpletextvalue = '" & request("customfield" & x) & "' "

			Case "select"
				' There will always be at least one selection picked
				sSql = sSql & "SET simpletextvalue = '" & request("customfield" & x) & "' "

			Case "checkbox"
				sNewValue = ""
				For Each Item In Request("customfield" & x)
					sNewValue = sNewValue & Item & vbcrlf
				Next 
				If sNewValue = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = "'" & sNewValue & "'"
				End If 
				sSql = sSql & "SET simpletextvalue = " & sNewValue & " "

			Case "textbox"
				If request("customfield" & x) = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = "'" & dbsafe(request("customfield" & x)) & "'"
				End If 
				sSql = sSql & "SET simpletextvalue = " & sNewValue & " " 

			Case "money"
				If request("customfield" & x) = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = FormatNumber(request("customfield" & x),2,,,0)
				End If 
				sSql = sSql & "SET moneyvalue = " & sNewValue & " "

			Case "integer"
				If request("customfield" & x) = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = request("customfield" & x) 
				End If 
				sSql = sSql & "SET intvalue = " & sNewValue & " "

			Case "textarea"
				If request("customfield" & x) = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = "'" & dbsafe(request("customfield" & x)) & "'"
				End If 
				sSql = sSql & "SET largetextvalue = " & sNewValue & " "

			Case Else 
				If request("customfield" & x) = "" Then
					sNewValue = "NULL"
				Else
					sNewValue = "'" & dbsafe(request("customfield" & x)) & "'"
				End If 
				sSql = sSql & "SET simpletextvalue = " & sNewValue & " "

		End Select 

		sSql = sSql & " WHERE customfieldid = " & request("customfieldid" & x)
		RunSQL sSql
	Next 
End If 

' Handle changes to the fees
' Loop through the listed fees
'response.write "MaxFees = " & request("maxfees") & "<br />"
'For x = 1 To CLng(request("maxfees"))
	' If the row exists
'	If request("permitfeeid" & x) <> "" Then
'		iPermitFeeId = CLng(request("permitfeeid" & x))
		'response.write "iPermitFeeId = " & iPermitFeeId & "<br />"

		' Handle marked/unmarked for inclusion here
		'response.write "includefee value = " & request("includefee" & iPermitFeeId) & "<br />"
'		If request("includefee" & iPermitFeeId) = "on" Then
'			sSql = "UPDATE egov_permitfees SET includefee = 1 WHERE permitfeeid = " & iPermitFeeId
'		Else 
'			sSql = "UPDATE egov_permitfees SET includefee = 0 WHERE permitfeeid = " & iPermitFeeId
'		End If 
'		response.write sSql & "<br />"
		' Only change if the fee is not required since disabled checkboxes are not passed as "on"
'		If FeeIsNotRequired( iPermitFeeId ) Then 
'			RunSQL sSql
'		End If 
'	End If 
'Next 

' Recalculate any valuation based fees
'response.write "RecalcFees<br /><br />"
RecalcValuationFees iPermitId, CDbl(sJobValue)

' Recalculate the total sq ft construction type fees
RecalcConstTypeFee iPermitId, CDbl(sTotalSqFt), iConstructiontyperate, "isconstructiontypegross"

' Recalculate the finished sq ft construction type fees
RecalcConstTypeFee iPermitId, CDbl(sFinishedSqFt), iConstructiontyperate, "isconstructiontypefinished"

' Recalculate the unfinished sq ft construction type fees
RecalcConstTypeFee iPermitId, CDbl(sUnFinishedSqFt), iConstructiontyperate, "isconstructiontypeunfinished"

' Recalculate the other sq ft construction type fees
RecalcConstTypeFee iPermitId, CDbl(sOtherSqFt), iConstructiontyperate, "isconstructiontypeother"

' Recalculate the percentage fees
RecalcResidentialUnitFees iPermitId, sResidentialUnits

' Recalculate the hourly rate fees
ReCalcExamHours iPermitId, CDbl(sHours)

' Recalculate the finished SQ Ft Fees
RecalcSqFtFee iPermitId, CDbl(sFinishedSqFt), "isfinishedsqft"

' Recalculate the unfinished SQ Ft Fees
RecalcSqFtFee iPermitId, CDbl(sUnFinishedSqFt), "isunfinishedsqft"

' Recalculate the total SQ Ft Fees
RecalcSqFtFee iPermitId, CDbl(sTotalSqFt), "istotalsqft"

' Recalculate the other SQ Ft Fees
RecalcSqFtFee iPermitId, CDbl(sOtherSqFt), "isothersqft"

' Recalculate the Volume Fees (Cu Ft) - this uses the OtherSqFt field and the label is changed to Volume
RecalcCuFtFee iPermitId, CDbl(sOtherSqFt), "iscuft"

' Recalculate the percentage fees
RecalcPercentageFees iPermitId

' recalculate the fee total using the ones flagged for inclusion
sResponse = SetPermitFeeTotal( iPermitId ) ' In permitcommonfunctions.asp


' Handle review includes
'For x = 1 To CLng(request("maxreviews"))
'	' If the row exists
'	If request("permitreviewid" & x) <> "" Then
'		iPermitReviewId = CLng(request("permitreviewid" & x))
'		'response.write "iPermitReviewId = " & iPermitReviewId & "<br />"
'
'		' Handle marked/unmarked for inclusion here
'		If request("includereview" & x) = "on" Then
'			sSql = "UPDATE egov_permitreviews SET isincluded = 1 WHERE permitreviewid = " & iPermitReviewId
'		Else 
'			sSql = "UPDATE egov_permitreviews SET isincluded = 0 WHERE permitreviewid = " & iPermitReviewId
'		End If 
'		response.write sSql & "<br />"
'		' Only change if the fee is not required since disabled checkboxes are not passed as "on"
'		If ReviewIsNotRequired( iPermitReviewId ) Then 
'			RunSQL sSql
'		End If 
'	End If 
'Next 

sInternalNotes = dbsafe(Trim(request("internalcomment")))
sPublicNotes = dbsafe(Trim(request("externalcomment")))
iPermitStatusId = GetPermitStatusId( iPermitId )		' in permitcommonfunctions.asp

' Figure out if the user entered notes
If sInternalNotes <> "" Or sPublicNotes <> "" Then 

	If sInternalNotes = "" Then
		sInternalNotes = "NULL"
	Else
		sInternalNotes = "'" & sInternalNotes & "'"
	End If 

	If sPublicNotes = "" Then
		sPublicNotes = "NULL"
	Else
		sPublicNotes = "'" & sPublicNotes & "'"
	End If 

	'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
	MakeAPermitLogEntry iPermitId, "'Permit Notes Added'", "NULL", sInternalNotes, sPublicNotes, iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"   ' in permitcommonfunctions.asp

End If 

' Push out the expiration date if the permit is not completed
iPermitStatusId = GetPermitStatusId( iPermitId )		' in permitcommonfunctions.asp

blnOverrideExpiration = GetExpirationOverride( iPermitId )
If StatusAllowsSaveChanges( iPermitStatusId ) and not blnOverrideExpiration Then 
	PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp
End If 

' Return to the edit page
response.redirect "permitedit.asp?permitid=" & iPermitid & "&activetab=" & request("activetab") & "&success=" & sSuccessMsg


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' string GetConstructionRate( iConstructionTypeId, iOccupancyTypeId )
'-------------------------------------------------------------------------------------------------
Function GetConstructionRate( ByVal iConstructionTypeId, ByVal iOccupancyTypeId )
	Dim sSql, oRs

	sSql = "SELECT constructiontyperate, isnotpermitted FROM egov_constructionfactors "
	sSql = sSql & " WHERE constructiontypeid = " & iConstructionTypeId
	sSql = sSql & " AND occupancytypeid = " & iOccupancyTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("isnotpermitted") Then
			GetConstructionRate = "NULL"
		Else 
			GetConstructionRate = oRs("constructiontyperate")
		End If 
	Else
		GetConstructionRate = "NULL"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' void GetContactTypeInfo iContactTypeId 
'-------------------------------------------------------------------------------------------------
Sub GetContactTypeInfo( ByVal iContactTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitcontacttypeid, company, firstname, lastname, address, city, state, zip, "
	sSql = sSql & " email, phone, cell, fax, userid, ISNULL(contractortypeid,0) AS contractortypeid, "
	sSql = sSql & " businesstypeid, statelicense, employeecount, reference1, reference2, reference3, "
	sSql = sSql & " otherlicensedcity1, otherlicensedcity2, generalliabilityagent, generalliabilityphone, "
	sSql = sSql & " workerscompagent, workerscompphone, autoinsuranceagent, autoinsurancephone, "
	sSql = sSql & " bondagent, bondagentphone "
	sSql = sSql & " FROM egov_permitcontacttypes WHERE permitcontacttypeid = " & iContactTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

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
	Else 
		sCompany = "NULL"
		sFirstname = "NULL"
		sLastname = "NULL"
		sAddress = "NULL"
		sCity = "NULL"
		sState = "NULL"
		sZip = "NULL"
		sEmail = "NULL"
		sPhone = "NULL"
		sCell = "NULL"
		sFax = "NULL"
		iContactUserId = "NULL"
		iContractorTypeId = "NULL"
		iBusinessTypeId = "NULL"
		sStateLicense = "NULL"
		sEmployeeCount = "NULL"
		sReference1 = "NULL"
		sReference2 = "NULL"
		sReference3 = "NULL"
		sOtherLicensedCity1 = "NULL"
		sOtherLicensedCity2 = "NULL"
		sGeneralLiabilityAgent = "NULL"
		sGeneralLiabilityPhone = "NULL"
		sWorkersCompAgent = "NULL"
		sWorkersCompPhone = "NULL"
		sAutoInsuranceAgent = "NULL"
		sAutoInsurancePhone = "NULL"
		sBondAgent = "NULL"
		sBondAgentPhone = "NULL"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'-------------------------------------------------------------------------------------------------
' void UpdateContactInfo iPermitContactid 
'-------------------------------------------------------------------------------------------------
Sub UpdateContactInfo( ByVal iPermitContactid, ByVal sContactType, ByVal iContactUserId )
	Dim sSql

	' put the contact info into the contact table
	If CLng(iPermitContactid) > CLng(0) Then
		If iContactTypeId = "NULL" Then
			' Delete the current record
			sSql = "DELETE FROM egov_permitcontacts "
		Else 
			' Update the existing record with the new contact
			sSql = "UPDATE egov_permitcontacts SET permitcontacttypeid = " & iContactTypeId
			sSql = sSql & ", company = " & sCompany
			sSql = sSql & ", firstname = " & sFirstname
			sSql = sSql & ", lastname = " & sLastname
			sSql = sSql & ", address = " & sAddress
			sSql = sSql & ", city = " & sCity
			sSql = sSql & ", state = " & sState
			sSql = sSql & ", zip = " & sZip
			sSql = sSql & ", email = " & sEmail
			sSql = sSql & ", phone = " & sPhone
			sSql = sSql & ", cell = " & sCell
			sSql = sSql & ", fax = " & sFax
			sSql = sSql & ", userid = " & iContactUserId
			sSql = sSql & ", contractortypeid = " & iContractorTypeId
			sSql = sSql & ", businesstypeid = " & iBusinessTypeId
			sSql = sSql & ", statelicense = " & sStateLicense
			sSql = sSql & ", employeecount = " & sEmployeeCount
			sSql = sSql & ", reference1 = " & sReference1
			sSql = sSql & ", reference2 = " & sReference2
			sSql = sSql & ", reference3 = " & sReference3
			sSql = sSql & ", otherlicensedcity1 = " & sOtherLicensedCity1
			sSql = sSql & ", otherlicensedcity2 = " & sOtherLicensedCity2
			sSql = sSql & ", generalliabilityagent = " & sGeneralLiabilityAgent
			sSql = sSql & ", generalliabilityphone = " & sGeneralLiabilityPhone
			sSql = sSql & ", workerscompagent = " & sWorkersCompAgent
			sSql = sSql & ", workerscompphone = " & sWorkersCompPhone
			sSql = sSql & ", autoinsuranceagent = " & sAutoInsuranceAgent
			sSql = sSql & ", autoinsurancephone = " & sAutoInsurancePhone
			sSql = sSql & ", bondagent = " & sBondAgent
			sSql = sSql & ", bondagentphone = " & sBondAgentPhone
			sSql = sSql & ", " & sContactType & " = 1 "
		End If 
		sSql = sSql & " WHERE permitcontactid = " & CLng(iPermitContactid)
		RunSQL sSql
	Else
		' Create a new record
		CreateNewContactRecord sContactType
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' void UpdatePrimaryContactInfo iPermitContactid 
'-------------------------------------------------------------------------------------------------
Sub UpdatePrimaryContactInfo( ByVal iPermitContactid )
	Dim sSql

	' put the contact info into the contact table
	If CLng(iPermitContactid) > CLng(0) Then
		If iPrimaryContactUserId = "NULL" Then
			' Delete the current record
			sSql = "DELETE FROM egov_permitcontacts "
		Else 
			' Update the existing record with the new contact
			sSql = "UPDATE egov_permitcontacts SET "
			sSql = sSql & " company = " & sCompany
			sSql = sSql & ", firstname = " & sFirstname
			sSql = sSql & ", lastname = " & sLastname
			sSql = sSql & ", address = " & sAddress
			sSql = sSql & ", city = " & sCity
			sSql = sSql & ", state = " & sState
			sSql = sSql & ", zip = " & sZip
			sSql = sSql & ", email = " & sEmail
			sSql = sSql & ", phone = " & sPhone
			sSql = sSql & ", cell = " & sCell
			sSql = sSql & ", fax = " & sFax
			sSql = sSql & ", userid = " & iPrimaryContactUserId
			sSql = sSql & ", isprimarycontact = 1 "
		End If 
		sSql = sSql & " WHERE permitcontactid = " & CLng(iPermitContactid)
		RunSQL sSql
	Else
		' Create a new record
		CreateNewPrimaryContactRecord 
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' integer CreateNewContactRecord( iContactTypeId, sContactType, iPermitId )
'-------------------------------------------------------------------------------------------------
Function CreateNewContactRecord( ByVal iContactTypeId, ByVal sContactType, ByVal iPermitId, ByVal iContactUserId )
	Dim sSql

	sSql = "INSERT INTO egov_permitcontacts ( permitid, permitcontacttypeid, orgid, company, firstname, "
	sSql = sSql & " lastname, address, city, state, zip, email, phone, cell, fax, "
	sSql = sSql & sContactType & ", userid, userpassword, userworkphone, emergencycontact, emergencyphone, "
	sSql = sSql & " neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable, "
	sSql = sSql & " residencyverified, contractortypeid, businesstypeid, statelicense, employeecount, reference1, "
	sSql = sSql & " reference2, reference3, otherlicensedcity1, otherlicensedcity2, generalliabilityagent, "
	sSql = sSql & " generalliabilityphone, workerscompagent, workerscompphone, autoinsuranceagent, "
	sSql = sSql & " autoinsurancephone, bondagent, bondagentphone )"
	sSql = sSql & " VALUES ( " & iPermitId & ", " & iContactTypeId & ", " & session("orgid") & ", "
	sSql = sSql & sCompany & ", " & sFirstname & ", " & sLastname & ", " & sAddress & ", "
	sSql = sSql & sCity & ", " & sState & ", " & sZip & ", " & sEmail & ", " & sPhone
	sSql = sSql & ", " & sCell & ", " & sFax & ", 1, " & iContactUserId
	sSql = sSql & ", " & sPassword & ", " & sWorkPhone & ", " & sEmergencyContact & ", "
	sSql = sSql & sEmergencyPhone & ", " & iNeighborhoodId & ", " & sResidentType & ", "
	sSql = sSql & sBusinessAddress & ", " & sUserUnit & ", " & sEmailnotavailable & ", "
	sSql = sSql & sResidencyVerified & ", " & iContractorTypeId & ", " & iBusinessTypeId & ", " & sStateLicense & ", " & sEmployeeCount & ", "
	sSql = sSql & sReference1 & ", " & sReference2 & ", " & sReference3 & ", " & sOtherLicensedCity1 & ", "
	sSql = sSql & sOtherLicensedCity2 & ", " & sGeneralLiabilityAgent & ", " & sGeneralLiabilityPhone & ", "
	sSql = sSql & sWorkersCompAgent & ", " & sWorkersCompPhone & ", " & sAutoInsuranceAgent & ", "
	sSql = sSql & sAutoInsurancePhone & ", " & sBondAgent & ", " & sBondAgentPhone & " )"
	'response.write sSql & "<br />"

	CreateNewContactRecord = RunIdentityInsert( sSql )
End Function 


'-------------------------------------------------------------------------------------------------
' void CreateContactInfo iContactTypeId, sContactType, iPermitId 
'-------------------------------------------------------------------------------------------------
Sub CreateContactInfo( ByVal iContactTypeId, ByVal sContactType, ByVal iPermitId )
	Dim iContactId

	GetContactTypeInfo iContactTypeId
	iContactId = CreateNewContactRecord( iContactTypeId, sContactType, iPermitId, iContactUserId )
	CreateNewLicenseRecords iContactTypeId, iContactId, iPermitId

End Sub 


'-------------------------------------------------------------------------------------------------
' void RemoveContact iPermitContactId 
'-------------------------------------------------------------------------------------------------
Sub RemoveContact( ByVal iPermitContactId )
	Dim sSql

	sSql = "UPDATE egov_permitcontacts SET ispriorcontact = 1 WHERE permitcontactid = " & iPermitContactId
	RunSQL sSql

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreateNewPrimaryContactRecord
'-------------------------------------------------------------------------------------------------
Sub CreateNewPrimaryContactRecord( )
	Dim sSql

	sSql = "INSERT INTO egov_permitcontacts ( permitid, permitcontacttypeid, orgid, company, firstname, "
	sSql = sSql & " lastname, address, city, state, zip, email, phone, cell, fax, "
	sSql = sSql & " isprimarycontact, userid ) VALUES ( " & iPermitId & ", NULL, " & session("orgid") & ", "
	sSql = sSql & sCompany & ", " & sFirstname & ", " & sLastname & ", " & sAddress & ", "
	sSql = sSql & sCity & ", " & sState & ", " & sZip & ", " & sEmail & ", " & sPhone
	sSql = sSql & ", " & sCell & ", " & sFax & ", 1, " & iPrimaryContactUserId & " )"

	RunSQL sSql

End Sub 


'-------------------------------------------------------------------------------------------------
' void GetPrimaryContactInfo
'-------------------------------------------------------------------------------------------------
Sub GetPrimaryContactInfo( ByVal iPrimaryContactUserId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_users WHERE userid = " & iPrimaryContactUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("userbusinessname")) Then
			sCompany = "NULL"
		Else
			sCompany = "'" & dbsafe(oRs("userbusinessname")) & "'"
		End If 
		If IsNull(oRs("userfname")) Then
			sFirstname = "NULL"
		Else
			sFirstname = "'" & dbsafe(oRs("userfname")) & "'"
		End If 
		If IsNull(oRs("userlname")) Then
			sLastname = "NULL"
		Else
			sLastname = "'" & dbsafe(oRs("userlname")) & "'"
		End If 
		If IsNull(oRs("useraddress")) Then
			sAddress = "NULL"
		Else
			sAddress = "'" & dbsafe(oRs("useraddress")) & "'"
		End If 
		If IsNull(oRs("usercity")) Then
			sCity = "NULL"
		Else
			sCity = "'" & dbsafe(oRs("usercity")) & "'"
		End If
		If IsNull(oRs("userstate")) Then
			sState = "NULL"
		Else
			sState = "'" & dbsafe(oRs("userstate")) & "'"
		End If
		If IsNull(oRs("userzip")) Then
			sZip = "NULL"
		Else
			sZip = "'" & dbsafe(oRs("userzip")) & "'"
		End If
		If IsNull(oRs("useremail")) Then
			sEmail = "NULL"
		Else
			sEmail = "'" & dbsafe(oRs("useremail")) & "'"
		End If
		If IsNull(oRs("userhomephone")) Then
			sPhone = "NULL"
		Else
			sPhone = "'" & oRs("userhomephone") & "'"
		End If
		If IsNull(oRs("usercell")) Then
			sCell = "NULL"
		Else
			sCell = "'" & oRs("usercell") & "'"
		End If
		If IsNull(oRs("userfax")) Then
			sFax = "NULL"
		Else
			sFax = "'" & oRs("userfax") & "'"
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
			sResidentType = "R"
		Else 
			sResidentType = "'" & oRs("residenttype") & "'"
		End If 
		sBusinessAddress = "'" & oRs("userbusinessaddress") & "'"
		sUserUnit = "'" & oRs("userunit") & "'"
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
		iContractorTypeId = "NULL"
	Else 
		sCompany = "NULL"
		sFirstname = "NULL"
		sLastname = "NULL"
		sAddress = "NULL"
		sCity = "NULL"
		sState = "NULL"
		sZip = "NULL"
		sEmail = "NULL"
		sPhone = "NULL"
		sCell = "NULL"
		sFax = "NULL"
		sPassword = "NULL"
		sWorkPhone = "NULL"
		sEmergencyContact = "NULL"
		sEmergencyPhone = "NULL"
		iNeighborhoodid = "NULL"
		sResidentType = "NULL"
		sBusinessAddress = "NULL"
		sUserUnit = "NULL"
		sEmailnotavailable = 0
		sResidencyVerified = 0
		iContractorTypeId = "NULL"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' boolean FeeIsNotRequired( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function FeeIsNotRequired( ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT isrequired FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isrequired") Then
			FeeIsNotRequired = False 
		Else
			FeeIsNotRequired = True 
		End If 
	Else
		FeeIsNotRequired = True  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' boolean ReviewIsNotRequired( iPermitReviewId )
'-------------------------------------------------------------------------------------------------
Function ReviewIsNotRequired( ByVal iPermitReviewId )
	Dim sSql, oRs

	sSql = "SELECT isrequired FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isrequired") Then
			ReviewIsNotRequired = False 
		Else
			ReviewIsNotRequired = True 
		End If 
	Else
		ReviewIsNotRequired = True  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function

Function GetExpirationOverride( intPermitID )
	retVal = false
	sSQL = "SELECT overrideexpiration FROM egov_permits WHERE permitid = '" & intPermitID & "'"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	If not oRs.EOF then retVal = oRs("overrideexpiration")
	oRs.Close
	Set oRs = Nothing

	GetExpirationOverride = retVal

End Function


%>
