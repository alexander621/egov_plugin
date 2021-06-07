<%
' END USED TO PREVENT ACCIDENTALLY RUNNING OF SCRIPT THRU ERRANT BROWSE OF THIS PAGE
' response.write "SECURITY GROUPS WERE NOT CREATED.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
' response.end

' -------------------------------------------------------------------------------------------------
' BEGIN SCRIPT INFORMATION
'--------------------------------------------------------------------------------------------------
' AUTHOR:		JOHN STULLENBERGER
' DATE:			11/28/2005
' REVISION:		1.0
' DESCRIPTION:	CREATES DEFAULT SECURITY ROLES AND GROUPS.
' TO RUN BROWER TO HTTP://WWW.EGOVLINK.COM/ECLINK/ADMIN/ADMIN/CREATEDEFAULTSECURITY.ASP
' 
' MODIFICATION HISTORY
' 1.0   02/15/06   Steve Loar - Added the group description array, changed grouprole to insert numbers,
'								and Added Role Permission Inserts, also deleted insert and update triggers on groups table
' -------------------------------------------------------------------------------------------------
' END SCRIPT INFORMATION
'--------------------------------------------------------------------------------------------------

' INITIALIZE VALUES AND OBJECTS
iOrgID = "63" ' CHANGE TO MATCH ORG
sDSN = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
arrRoles = Array("Manage Action Alerts","Manage Form Design","Manage All Requests","Manage Own Requests","Manage Calendar","Manage Dept Requests","Manage Documents","Manage Payments","Manage Security")
arrGroups = Array("Manage Action Alerts","Manage Form Design","Manage All Requests","Manage Own Requests","Manage Calendar","Manage Dept Requests","Manage Documents","Manage Payments","Manage Security")
arrGrpDesc = Array("Allows user to manage notifications and escalations as well as department and category.","Allows users to access Action form design tool.","Allows users to manage all action line items no matter what department or who is assigned.","Allows users to mange action requests assigned to them only.","This allows users to manage the calendar","Manage Action Line Requests for own department only.","This allows users to manage documents","This allows users to manage payments","This allows users to manage security")


response.write "<div style=""background-color:#e0e0e0;border: solid 1px #000000;padding:10px;FONT-FAMILY: Verdana,Tahoma,Arial;font-size:10px;"">"
response.write "<p><b>Adding security roles and groups for orgid(" & iorgid & ")...</b></p>"
response.write Now() & "<br /><br />"


' CREATE ROLES,GROUPS, AND LINK
For r = 0 to UBOUND(arrRoles)
	
	response.write "<h2>" & arrRoles(r) & "</h2>"
	' CREATE ROLE
	response.write "<P>INSERT INTO ROLES (RoleName,RoleDescription,OrgID) VALUES ('" & arrRoles(r) & "','System Role.','" & iOrgID & "')</p>"
	iRoleID = RunSQL("INSERT INTO ROLES (RoleName,RoleDescription,OrgID) VALUES ('" & arrRoles(r) & "','System Role.','" & iOrgID & "')")
	response.write "iRoleId = " & iRoleId
	
	' CREATE GROUP
	response.write "<P>INSERT INTO GROUPS (GroupName,GroupDescription,OrgID,GroupType) VALUES ('" & arrGroups(r) & "','" & arrGrpDesc(r) & "'," & iOrgID & ",1)</p>"
	iGroupID = RunSQL("INSERT INTO GROUPS (GroupName,GroupDescription,OrgID,GroupType) VALUES ('" & arrGroups(r) & "','" & arrGrpDesc(r) & "'," & iOrgID & ",1)")
	response.write "iGroupID = " & iGroupID

	' CREATE GROUP TO ROLE LINK
	response.write "<P>INSERT INTO GROUPSROLES (GroupID,RoleID) VALUES (" & iGroupID & "," & iRoleID & ")</p>"
	iGroupRoleID = RunSQL("INSERT INTO GROUPSROLES (GroupID,RoleID) VALUES (" & iGroupID & "," & iRoleID & ")")
	response.write "iGroupRoleID = " & iGroupRoleID

	' Create the Role Permissions
	response.write "<br />Inserting into rolepermission table"
	Select Case r
		Case 0 ' Manage Action Alerts
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 429)" ) ' CanManageActionRequests
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 704)" ) ' CanManageActionAlerts
			response.write "<br />iRolPerm = " & iRolPerm
		Case 1  ' Manage Form Design
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 429)" ) ' CanManageActionRequests
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 701)" ) ' CanEditActionForms
			response.write "<br />iRolPerm = " & iRolPerm
		Case 2  ' Manage All Requests
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 429)" ) ' CanManageActionRequests
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 700)" ) ' CanViewAllActionItems
			response.write "<br />iRolPerm = " & iRolPerm
		Case 3  ' Manage Own Requests
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 429)" ) ' CanManageActionRequests
			response.write "<br />iRolPerm = " & iRolPerm 
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 698)" ) ' CanViewOwnActionItems
			response.write "<br />iRolPerm = " & iRolPerm
		Case 4  ' Manage Calendar
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 2)" ) ' CanEditEvents
			response.write "<br />iRolPerm = " & iRolPerm
		Case 5  ' Manage Dept Requests
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 429)" ) ' CanManageActionRequests
			response.write "<br />iRolPerm = " & iRolPerm 
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 699)" ) ' CanViewDeptActionItems
			response.write "<br />iRolPerm = " & iRolPerm
		Case 6  ' Manage Documents
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 893)" ) ' CanEditDocuments
			response.write "<br />iRolPerm = " & iRolPerm
		Case 7  ' Manage Payments
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 897)" ) ' CandEditPayments
			response.write "<br />iRolPerm = " & iRolPerm
		Case 8  ' Manage Security
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 33)" ) ' CanEditCommittee
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 30)" ) ' CanRegisterCommittee
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 32)" ) ' CanRegisterContact
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 41)" ) ' CanRegisterRole
			response.write "<br />iRolPerm = " & iRolPerm
			iRolPerm = RunSQL("Insert Into rolespermissions ( roleid, permissionid) values (" & iRoleId & ", 31)" ) ' CanRegisterUser
			response.write "<br />iRolPerm = " & iRolPerm
	End Select 

Next

response.write "<p><b>Done Adding security roles and groups.</b></p>"
response.write Now()
response.write "</div>"



'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'-------------------------------------------------------------------------------------------------
Function RunSQL(sInsertStatement)
	Dim sSQL
	RunSQL = 0

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	'response.write sSQL
	oInsert.Open sSQL, sDSN, 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.close
	Set oInsert = Nothing

	RunSQL = iReturnValue

End Function


%>
