<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: facility_name_save.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
''
' DESCRIPTION:  Delete a facility.
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGE - INITIAL VERSION
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

subSaveChanges dbsafe(request("sFacilityName")), request("iFacilityId"), request("seltemplate"),request("chkisviewable"),request("chkisreservable"), request("categoryid")

'--------------------------------------------------------------------------------------------------
' SUB SUBSAVECHANGES(SFACILITYNAME, IFACILITYID, ISELTEMPLATE)
'--------------------------------------------------------------------------------------------------
Sub subSaveChanges( ByVal sFacilityName, ByVal iFacilityId, ByVal iSelTemplate, ByVal blnviewable, ByVal blnreservable, ByVal iCategoryId )
	Dim sSql

	If blnviewable = "on" Then
		blnviewable = 1
	Else
		blnviewable = 0
	End If 

	If blnreservable = "on" Then 
		blnreservable = 1
	Else
		blnreservable = 0
	End If 
	
	' NEW RECORD
	If iFacilityId = "0" Then
		sSql = "INSERT INTO egov_facility ( facilityname, facilitytemplateid, orgid, pricetypegroupid ) Values ( " 
		sSql = sSql & "'" & sFacilityName & "', " &  iSelTemplate & ", " & session("orgid") & ", 1 )"

		iFacilityId = RunInsertStatement( sSql )
	Else
		' UPDATE EXISTING RECORDS
		sSqL = "UPDATE egov_facility SET facilityname = '" & sFacilityName & "', facilitytemplateid = " & iSelTemplate & ", isviewable = " & blnviewable & ", isreservable = " & blnreservable 
		sSql = sSql & " WHERE facilityid = " & iFacilityId 

		RunSQLStatement sSql 
	End If

	' Set the Facility Category
	sSql = "Delete from egov_recreation_category_to_item Where itemid = " & iFacilityId
	RunSQLStatement sSql 

	sSql = "Insert Into egov_recreation_category_to_item (categoryid, itemid) Values ( " & iCategoryId & ", " & iFacilityId & " )"
	RunSQLStatement sSql 

	' REDIRECT TO facility edit page
	response.redirect( "facility_edit.asp?facilityid=" & iFacilityId )

End Sub


%>