<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcontactupdate.asp
' AUTHOR: Steve Loar
' CREATED: 03/19/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit contact types from the permit. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   03/19/2008   Steve Loar - INITIAL VERSION
' 1.1	06/05/2008	Steve Loar - License Date dropped, Licensee addded
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactTypeid, sSql, sAddress, sCity, sState, sZip, sEmail, sPhone, sFax, sCell, sUserId
Dim iMaxLicenseRows, sLicenseType, sLicenseNumber, sLicensee, sLicenseEndDate, sFirstName, iPermitContactId
Dim sLastName, sCompany, bUpdate, sReturn, iPermitId, iPermitStatusId, bChangesPropagate, iContractorTypeId
Dim sLicenseClass, iLicenseTypeId, iMaxUsers, iIsOrganization, iBusinessTypeId, sStateLicense, sEmployeeCount
Dim sReference1, sReference2, sReference3, sOtherLicensedCity1, sOtherLicensedCity2, sGeneralLiabilityAgent
Dim sGeneralLiabilityPhone, sWorkersCompAgent, sWorkersCompPhone, sAutoInsuranceAgent, sAutoInsurancePhone
Dim sBondAgent, sBondAgentPhone

iPermitContactTypeid = CLng(request("permitcontacttypeid"))
iPermitId = CLng(request("permitid"))
iPermitStatusId = CLng(request("permitstatusid"))
iContractorTypeId = CLng(request("contractortypeid"))
iPermitContactId = CLng(request("permitcontactid"))
iIsOrganization = clng(request("isorganization"))

bChangesPropagate = StatusAllowsChangesToPropagate( iPermitStatusId )
bUpdate = False 
sReturn = ""

If request("firstname") = "" Then
	sFirstName = "NULL"
Else
	sFirstName = "'" & dbsafe(request("firstname")) & "'"
End If 

If request("lastname") = "" Then
	sLastName = "NULL"
Else
	sLastName = "'" & dbsafe(request("lastname")) & "'"
End If 

If request("company") = "" Then
	sCompany = "NULL"
Else
	sCompany = "'" & dbsafe(request("company")) & "'"
End If 

If request("address") = "" Then
	sAddress = "NULL"
Else
	sAddress = "'" & dbsafe(request("address")) & "'"
End If 

If request("city") = "" Then
	sCity = "NULL"
Else
	sCity = "'" & dbsafe(request("city")) & "'"
End If 

If request("state") = "" Then
	sState = "NULL"
Else
	sState = "'" & dbsafe(request("state")) & "'"
End If 

If request("zip") = "" Then
	sZip = "NULL"
Else
	sZip = "'" & dbsafe(request("zip")) & "'"
End If 

If request("email") = "" Then
	sEmail = "NULL"
Else
	sEmail = "'" & dbsafe(request("email")) & "'"
End If 

If request("phone") = "" Then
	sPhone = "NULL"
Else
	sPhone = "'" & request("phone") & "'"
End If 

If request("fax") = "" Then
	sFax = "NULL"
Else
	sFax = "'" & request("fax") & "'"
End If 

If request("cell") = "" Then
	sCell = "NULL"
Else
	sCell = "'" & request("cell") & "'"
End If 

sUserId = "NULL"

If request("businesstypeid") <> "" Then 
	If CLng(request("businesstypeid")) > CLng(0) Then 
		iBusinessTypeId = CLng(request("businesstypeid"))
	Else 
		iBusinessTypeId = "NULL"
	End If 
Else
	iBusinessTypeId = "NULL"
End If 

If request("statelicense") = "" Then
	sStateLicense = "NULL"
Else
	sStateLicense = "'" & dbsafe(request("statelicense")) & "'"
End If 

If request("employeecount") = "" Then
	sEmployeeCount = "NULL"
Else
	sEmployeeCount = "'" & dbsafe(request("employeecount")) & "'"
End If 

If request("reference1") = "" Then
	sReference1 = "NULL"
Else
	sReference1 = "'" & dbsafe(request("reference1")) & "'"
End If 

If request("reference2") = "" Then
	sReference2 = "NULL"
Else
	sReference2 = "'" & dbsafe(request("reference2")) & "'"
End If 
If request("reference3") = "" Then
	sReference3 = "NULL"
Else
	sReference3 = "'" & dbsafe(request("reference3")) & "'"
End If 

If request("otherlicensedcity1") = "" Then
	sOtherLicensedCity1 = "NULL"
Else
	sOtherLicensedCity1 = "'" & dbsafe(request("otherlicensedcity1")) & "'"
End If 

If request("otherlicensedcity2") = "" Then
	sOtherLicensedCity2 = "NULL"
Else
	sOtherLicensedCity2 = "'" & dbsafe(request("otherlicensedcity2")) & "'"
End If 

If request("generalliabilityagent") = "" Then
	sGeneralLiabilityAgent = "NULL"
Else
	sGeneralLiabilityAgent = "'" & dbsafe(request("generalliabilityagent")) & "'"
End If 

If request("generalliabilityphone") = "" Then
	sGeneralLiabilityPhone = "NULL"
Else
	sGeneralLiabilityPhone = "'" & dbsafe(request("generalliabilityphone")) & "'"
End If 

If request("workerscompagent") = "" Then
	sWorkersCompAgent = "NULL"
Else
	sWorkersCompAgent = "'" & dbsafe(request("workerscompagent")) & "'"
End If 

If request("workerscompphone") = "" Then
	sWorkersCompPhone = "NULL"
Else
	sWorkersCompPhone = "'" & dbsafe(request("workerscompphone")) & "'"
End If 

If request("autoinsuranceagent") = "" Then
	sAutoInsuranceAgent = "NULL"
Else
	sAutoInsuranceAgent = "'" & dbsafe(request("autoinsuranceagent")) & "'"
End If 

If request("autoinsurancephone") = "" Then
	sAutoInsurancePhone = "NULL"
Else
	sAutoInsurancePhone = "'" & dbsafe(request("autoinsurancephone")) & "'"
End If 

If request("bondagent") = "" Then
	sBondAgent = "NULL"
Else
	sBondAgent = "'" & dbsafe(request("bondagent")) & "'"
End If 

If request("bondagentphone") = "" Then
	sBondAgentPhone = "NULL"
Else
	sBondAgentPhone = "'" & dbsafe(request("bondagentphone")) & "'"
End If 


iMaxUsers = CLng(request("maxusers"))

If iPermitContactTypeid = CLng(0) Then 
	bChangesPropagate = True 
	sSql = "INSERT INTO egov_permitcontacttypes ( orgid, firstname, lastname, company,"
	sSql = sSql & " address, city, state, zip, email, phone, fax, cell, userid, contractortypeid, "
	sSql = sSql & " isorganization, businesstypeid, statelicense, employeecount, reference1, "
	sSql = sSql & " reference2, reference3, otherlicensedcity1, otherlicensedcity2, generalliabilityagent, "
	sSql = sSql & " generalliabilityphone, workerscompagent, workerscompphone, autoinsuranceagent, "
	sSql = sSql & " autoinsurancephone, bondagent, bondagentphone ) "
	sSql = sSql & " VALUES ( " & session("orgid") & ", " & sFirstName & ", " & sLastName & ", " & sCompany & ", " 
	sSql = sSql & sAddress & ", " & sCity & ", " & sState & ", " & sZip & ", " & sEmail
	sSql = sSql & ", " & sPhone & ", " & sFax & ", " & sCell & ", " & sUserId & ", " & iContractorTypeId
	sSql = sSql & ", " & iIsOrganization & ", " & iBusinessTypeId & ", " & sStateLicense & ", " & sEmployeeCount & ", "
	sSql = sSql & sReference1 & ", " & sReference2 & ", " & sReference3 & ", " & sOtherLicensedCity1 & ", "
	sSql = sSql & sOtherLicensedCity2 & ", " & sGeneralLiabilityAgent & ", " & sGeneralLiabilityPhone & ", "
	sSql = sSql & sWorkersCompAgent & ", " & sWorkersCompPhone & ", " & sAutoInsuranceAgent & ", "
	sSql = sSql & sAutoInsurancePhone & ", " & sBondAgent & ", " & sBondAgentPhone & " )"
	iPermitContactTypeid = RunIdentityInsert( sSql ) 

	' Need to add any associated users
	For x = 1 To iMaxUsers
		If request("userid" & x) <> "" Then 
			sSql = "INSERT INTO egov_permitcontacttypes_to_users ( permitcontacttypeid, userid ) VALUES ( " & iPermitContactTypeid & ", " & request("userid" & x) & " )"
			RunSQL sSql
		End If 
	Next 
Else 
	If bChangesPropagate Then
		bUpdate = True 
		sSql = "UPDATE egov_permitcontacttypes SET firstname = " & sFirstName
		sSql = sSql & ", lastname = " & sLastName
		sSql = sSql & ", company = " & sCompany
		sSql = sSql & ", address = " & sAddress
		sSql = sSql & ", city = " & sCity
		sSql = sSql & ", state = " & sState
		sSql = sSql & ", zip = " & sZip
		sSql = sSql & ", email = " & sEmail
		sSql = sSql & ", phone = " & sPhone
		sSql = sSql & ", fax = " & sFax
		sSql = sSql & ", cell = " & sCell
		sSql = sSql & ", userid = " & sUserId
		sSql = sSql & ", contractortypeid = " & iContractorTypeId
		sSql = sSql & ", isorganization = " & iIsOrganization
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
		sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitcontacttypeid = " & iPermitContactTypeid
		RunSQL sSql 

		' Update contact information on any active permits
		sSql = "SELECT P.permitid, C.permitcontactid FROM egov_permits P, egov_permitcontacts C, egov_permitstatuses S "
		sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitstatusid = S.permitstatusid "
		sSql = sSql & " AND S.iscompletedstatus = 0 AND C.permitcontacttypeid = " & iPermitContactTypeid

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 3, 1

		Do While Not oRs.EOF
			sSql = "UPDATE egov_permitcontacts SET firstname = " & sFirstName
			sSql = sSql & ", lastname = " & sLastName
			sSql = sSql & ", company = " & sCompany
			sSql = sSql & ", address = " & sAddress
			sSql = sSql & ", city = " & sCity
			sSql = sSql & ", state = " & sState
			sSql = sSql & ", zip = " & sZip
			sSql = sSql & ", email = " & sEmail
			sSql = sSql & ", phone = " & sPhone
			sSql = sSql & ", fax = " & sFax
			sSql = sSql & ", cell = " & sCell
			sSql = sSql & ", userid = " & sUserId
			sSql = sSql & ", contractortypeid = " & iContractorTypeId
			sSql = sSql & ", isorganization = " & iIsOrganization
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
			sSql = sSql & " WHERE permitid = " & oRs("permitid") & " AND permitcontacttypeid = " & iPermitContactTypeid
			RunSQL sSql
			' Update their licenses
			sSql = "DELETE FROM egov_permitcontacts_licenses WHERE permitcontactid = " & oRs("permitcontactid")
			RunSQL sSql
			NewPermitContactLicenses oRs("permitid"), oRs("permitcontactid")
			oRs.MoveNext
		Loop
		oRs.Close
		Set oRs = Nothing 

		' Delete any licenses
		sSql = "DELETE FROM egov_permitcontacttype_licenses WHERE permitcontacttypeid = " & iPermitContactTypeid
		RunSQL sSql
	Else
		' update the contacts for the one permit, since it does not propagate
		sSql = "UPDATE egov_permitcontacts SET firstname = " & sFirstName
		sSql = sSql & ", lastname = " & sLastName
		sSql = sSql & ", company = " & sCompany
		sSql = sSql & ", address = " & sAddress
		sSql = sSql & ", city = " & sCity
		sSql = sSql & ", state = " & sState
		sSql = sSql & ", zip = " & sZip
		sSql = sSql & ", email = " & sEmail
		sSql = sSql & ", phone = " & sPhone
		sSql = sSql & ", fax = " & sFax
		sSql = sSql & ", cell = " & sCell
		sSql = sSql & ", userid = " & sUserId
		sSql = sSql & ", contractortypeid = " & iContractorTypeId
		sSql = sSql & ", isorganization = " & iIsOrganization
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
		sSql = sSql & " WHERE permitid = " & iPermitId & " AND permitcontacttypeid = " & iPermitContactTypeid
		RunSQL sSql

		' Update their licenses
		sSql = "SELECT C.permitcontactid FROM egov_permits P, egov_permitcontacts C "
		sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitid = " & iPermitId 
		sSql = sSql & " AND C.permitcontacttypeid = " & iPermitContactTypeid

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 3, 1
		
		Do While Not oRs.EOF
			sSql = "DELETE FROM egov_permitcontacts_licenses WHERE permitcontactid = " & oRs("permitcontactid")
			RunSQL sSql
			NewPermitContactLicenses iPermitId, oRs("permitcontactid")
			oRs.MoveNext
		Loop
		oRs.Close
		Set oRs = Nothing 
	End If 

End If 

If bChangesPropagate Then

	If clng(iIsOrganization) = clng(0) Then ' Only Contractors
		' Add any licenses to the contact type 
		iMaxLicenseRows = CLng(request("maxlicenserows"))

		For x = 0 To iMaxLicenseRows
			' See if the license type data exists
			If request("licenseenddate" & x) <> "" Then 
				If CLng(request("licensetypeid" & x)) = CLng(0) Then
					iLicenseTypeId = "NULL"
				Else
					iLicenseTypeId = request("licensetypeid" & x)
				End If  
				If request("licensenumber" & x) = "" Then
					sLicenseNumber = "NULL"
				Else
					sLicenseNumber = "'" & dbsafe(request("licensenumber" & x)) & "'"
				End If 
				If request("licenseclass" & x) = "" Then
					sLicenseClass = "NULL"
				Else
					sLicenseClass = "'" & dbsafe(request("licenseclass" & x)) & "'"
				End If
				If request("licensee" & x) = "" Then
					sLicensee = "NULL"
				Else
					sLicensee = "'" & dbsafe(request("licensee" & x)) & "'"
				End If
				If request("licenseenddate" & x) = "" Then
					sLicenseEndDate = "NULL"
				Else
					sLicenseEndDate = "'" & dbsafe(request("licenseenddate" & x)) & "'"
				End If
				sSql = "INSERT INTO egov_permitcontacttype_licenses (permitcontacttypeid, licensetypeid, licensenumber, licensee, licenseenddate, licenseclass ) VALUES ( "
				sSql = sSql & iPermitContactTypeid & ", " & iLicenseTypeId & ", " & sLicenseNumber & ", " & sLicensee & ", " & sLicenseEndDate & ", " & sLicenseClass & " )"
				'response.write sSql
				RunSQL sSql
			End If 
		Next 
	End If 
End If 


' Update the associated registered users for can add others and Primary Contact
' first set all the flags to false
sSql = "UPDATE egov_permitcontacttypes_to_users SET canaddothers = 0, isprimarycontact = 0 WHERE permitcontacttypeid = " & iPermitContactTypeid
RunSQL sSql

' Set any Can Add Others flags
For x = 1 To iMaxUsers
	If request("canaddothers" & x) <> "" Then 
		sSql = "UPDATE egov_permitcontacttypes_to_users SET canaddothers = 1 WHERE permitcontacttypeid = " & iPermitContactTypeid 
		sSql = sSql & " AND userid = " & request("canaddothers" & x)
		RunSQL sSql
	End If 
Next 

' Set the Primary Contact flag if any
If request("isprimarycontact") <> "" Then
	sSql = "UPDATE egov_permitcontacttypes_to_users SET isprimarycontact = 1 WHERE permitcontacttypeid = " & iPermitContactTypeid 
	sSql = sSql & " AND userid = " & request("isprimarycontact")
	RunSQL sSql
End If 


' Update the contractor type on this permit and contactid only
'response.write "contractortypeid = " & iContractorTypeId & "<br />"
'sSql = "UPDATE egov_permitcontacts SET contractortypeid = " & iContractorTypeId & " WHERE permitcontactid = " & iPermitContactId
'response.write sSql 
'RunSQL sSql


sReturn = iPermitContactTypeid
If request("firstname") <> "" Then 
	sReturn = sReturn & "$" & request("firstname") & " " & request("lastname")
	If request("firstname") <> "" Then 
		sReturn = sReturn & " (" & request("company") & ")"
	End If 
Else
	sReturn = sReturn & "$" & request("company")
End If 
response.write sReturn

'response.write "Success"


'-------------------------------------------------------------------------------------------------
' Sub NewPermitContactLicenses( iPermitid, iPermitcontactid )
'-------------------------------------------------------------------------------------------------
Sub NewPermitContactLicenses( iPermitid, iPermitcontactid )
	Dim x, iMaxLicenseRows, sLicenseType, sLicenseNumber, sLicensee, sLicenseEndDate, sLicenseClass

	iMaxLicenseRows = CLng(request("maxlicenserows"))

	For x = 0 To iMaxLicenseRows
		' See if the license type data exists
		If request("licensetype" & x) <> "0" And request("licenseenddate" & x) <> "" Then 
			If CLng(request("licensetypeid" & x)) = CLng(0) Then
				iLicenseTypeId = "NULL"
			Else
				iLicenseTypeId = request("licensetypeid" & x)
			End If 
			If request("licensenumber" & x) = "" Then
				sLicenseNumber = "NULL"
			Else
				sLicenseNumber = "'" & dbsafe(request("licensenumber" & x)) & "'"
			End If 
			If request("licenseclass" & x) = "" Then
				sLicenseClass = "NULL"
			Else
				sLicenseClass = "'" & dbsafe(request("licenseclass" & x)) & "'"
			End If
			If request("licensee" & x) = "" Then
				sLicensee = "NULL"
			Else
				sLicensee = "'" & dbsafe(request("licensee" & x)) & "'"
			End If
			If request("licenseenddate" & x) = "" Then
				sLicenseEndDate = "NULL"
			Else
				sLicenseEndDate = "'" & dbsafe(request("licenseenddate" & x)) & "'"
			End If
			sSql = "INSERT INTO egov_permitcontacts_licenses ( permitid, permitcontactid, licensetypeid, "
			sSql = sSql & " licensenumber, licensee, licenseenddate, licenseclass ) VALUES ( "
			sSql = sSql & iPermitid & ", " & iPermitcontactid & ", " & iLicenseTypeId & ", " & sLicenseNumber & ", "
			sSql = sSql  & sLicensee & ", "& sLicenseEndDate & ", " & sLicenseClass & " )"
			'response.write sSql & "<br />"
			RunSQL sSql
		End If 
	Next 

End Sub 



%>
