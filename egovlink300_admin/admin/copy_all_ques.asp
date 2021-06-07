<%
' END USED TO PREVENT ACCIDENTALLY RUNNING OF SCRIPT THRU ERRANT BROWSE OF THIS PAGE
' response.write "FORMS WERE NOT ADDED.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
' response.end


response.write "<br />Started: " & Now() & "<br />"
' THIS SCRIPT COPIES ALL FORMS TO A NEW ORGINIZATION

' INITIALIZE VARIABLES
iNewOrg = "66"		' The New Orgid						  -- Change this 
iAdminId = "2345"	' The id of the initial admin person  -- Change this
iDeptId = "2515"	' The City Employees Department id	  -- Change this

' GET ALL FORMS FOR ADMINISTRATIVE ORGINIZATION (0) 
'					except the contact E-GOV form - Steve Loar 4/10/2006
sSQL = "SELECT * FROM egov_action_request_forms WHERE orgid = 47 " 

Set oAllForms = Server.CreateObject("ADODB.Recordset")
oAllForms.Open sSQL, Application("DSN"), 3, 2
	
If NOT oAllForms.EOF Then
	Do While NOT oAllForms.EOF 
		subCopyForm oAllForms("action_form_id"), iNewOrg, iAdminId, iDeptId 
		oAllForms.MoveNext
	Loop 
End If

oAllForms.close
Set oAllForms = Nothing

response.write "<br />Finished: " & Now() & "<br />"


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBCOPYFORM(IFORMID,iNewOrg, iAdminId, iDeptId)
'--------------------------------------------------------------------------------------------------
Sub subCopyForm( iFormID,iNewOrg, iAdminId, iDeptId )
	
	' COPY the FORMS
	iNewFormID = fnCopySQLDataRow( "action_form_id", iFormID, "egov_action_request_forms", iNewOrg, iAdminId, iDeptId )

	' COPY FORM QUESTIONS
	SubCopyFormQuestions iFormID, iNewOrg, iNewFormID 

	' GET CATEGORY ID 
	iNewCategoryID = GetCategoryID( iFormID )

	'CODE TO ASSIGN FORM TO CATEGORY
	subAssignFormtoCategory iNewFormID, iNewCategoryID, iNewOrg 

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNCOPYSQLDATAROW(SPRIMARYKEY,IPRIMARYKEYID,STABLENAME)
'--------------------------------------------------------------------------------------------------
Function fnCopySQLDataRow(sPrimaryKey,iPrimaryKeyID,sTableName, iNewOrg, iAdminId, iDeptId)

	iReturnValue = 0 

	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * FROM " & sTableName & " WHERE " & sPrimaryKey & "='" & iPrimaryKeyID & "'"
	oSchema.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oSchema.EOF Then

		' BUILD COLUMN NAME LIST
		For Each fldLoop in oSchema.Fields
			iCount = iCount + 1 
			If iCount <> 1 Then
				sFieldList = sFieldList & fldLoop.Name 
				If iCount <> oSchema.Fields.Count Then
					sFieldList = sFieldList & ","
				End If
			End If
		Next

		' WRITE DATA ROWS
		Do while NOT  oSchema.EOF 
			iCount = 0
			sValueList = ""
			If NOT oSchema.EOF Then
				For Each fldLoop in oSchema.Fields
					iCount = iCount + 1 
					If iCount <> 1 Then ' SKIP FIRST FIELD (AUTO IDENTITY)
						
						' CUSTOM FIELD HANDLING
						Select Case LCase(fldLoop.Name)

							'Case "action_form_name"
								'sValueList = sValueList & "'NEW ACTION FORM'"

							Case "orgid" ' Try to use the new orgid
								sValueList = sValueList & "'" & iNewOrg & "'"
							
							Case "assigned_userid"
								sValueList = sValueList & "'" & iAdminId & "'"

							Case "assigned_userid2"
								sValueList = sValueList & "'" & 0 & "'"
							
							Case "assigned_userid3"
								sValueList = sValueList & "'" & 0 & "'"

							Case "deptid"
								sValueList = sValueList & "'" & iDeptId & "'"

							Case Else

								' ADD DATA TO STRING
								If fldLoop.Type = 11 Then
									sValueList = sValueList & "'" & fnBitConvert(fldLoop.Value) & "'"
								Else
									sValueList = sValueList & "'" & DBSafe(fldLoop.Value) & "'"
								End If

						End Select 

						' ADD TRAILING COMMA IF NECESSARY
						If iCount <> oSchema.Fields.Count Then
							sValueList = sValueList & ","
						End If
					End If
				Next

				sInsertStatement = "INSERT INTO " & sTableName & " (" & sFieldList & ") VALUES (" & sValueList & ")"
				'DEBUG DATA: 
				response.write sInsertStatement & "<br /><br />" & vbcrlf
				response.flush
				
				' INSERT NEW ROW INTO DATABASE AND GET ROWID
				sSQL = "SET NOCOUNT ON;" &_
				sInsertStatement &_
				"SELECT @@IDENTITY AS ROWID;"
				Set oInsert = Server.CreateObject("ADODB.Recordset")
				oInsert.Open sSQL, Application("DSN"), 3, 1
				iReturnValue = oInsert("ROWID")
				'DEBUG DATA: response.write "(" & iReturnValue & ")"
				
				oSchema.MoveNext
			End If
		Loop

	End If

	oSchema.close
	Set oSchema = Nothing 

	fnCopySQLDataRow = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' SUB SUBCOPYFORMQUESTIONS(IFORMID,IORGID,INEWFORMID)
'--------------------------------------------------------------------------------------------------
Sub SubCopyFormQuestions(iFormID,iOrgID,iNewFormID)

	sSQL = "Select * FROM egov_action_form_questions WHERE formid='" & iFormID & "' ORDER BY SEQUENCE"
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 2
	
	If NOT oQuestions.EOF Then
		Do While NOT oQuestions.EOF 
			iTempReturnID = fnCopySQLQuestionRow("questionid",oQuestions("questionid"),"egov_action_form_questions",iNewFormID,iOrgID)
			oQuestions.MoveNext
		Loop 
	End If

	oQuestions.close
	Set oQuestions = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION  FNBITCONVERT(BLNVALUE)
'--------------------------------------------------------------------------------------------------
Function  fnBitConvert(blnValue)

	iReturnValue = 0

	If blnValue Then
		iReturnValue = 1
	End If	

	fnBitConvert = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION FNCOPYSQLQUESTIONROW(SPRIMARYKEY,IPRIMARYKEYID,STABLENAME,IFORMID,IORGID)
'--------------------------------------------------------------------------------------------------
Function fnCopySQLQuestionRow(sPrimaryKey,iPrimaryKeyID,sTableName,iFormID,iOrgID)

	iReturnValue = 0 

	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * FROM " & sTableName & " WHERE " & sPrimaryKey & "='" & iPrimaryKeyID & "'"
	oSchema.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oSchema.EOF Then

		' BUILD COLUMN NAME LIST
		For Each fldLoop in oSchema.Fields
			iCount = iCount + 1 
			If iCount <> 1 Then
				sFieldList = sFieldList & fldLoop.Name 
				If iCount <> oSchema.Fields.Count Then
					sFieldList = sFieldList & ","
				End If
			End If
		Next

		' WRITE DATA ROWS
		Do while NOT  oSchema.EOF 
			iCount = 0
			sValueList = ""
			If NOT oSchema.EOF Then
				For Each fldLoop in oSchema.Fields
					iCount = iCount + 1 
					If iCount <> 1 Then ' SKIP FIRST FIELD (AUTO IDENTITY)
						
						sValue = DbSafe(fldLoop.Value)

						' CHANGE FORMID AND ORGID VALUES
						If fldLoop.Name = "formid" Then
							sValue = iFormID
						End If

						If fldLoop.Name = "orgid" Then
							sValue = iOrgID
						End If

						' ADD DATA TO STRING
						If fldLoop.Type = 11 Then
							sValueList = sValueList & "'" & fnBitConvert(sValue) & "'"
						Else
							sValueList = sValueList & "'" & sValue & "'"
						End If

						' ADD TRAILING COMMA IF NECESSARY
						If iCount <> oSchema.Fields.Count Then
							sValueList = sValueList & ","
						End If
					End If
				Next

				sInsertStatement = "INSERT INTO " & sTableName & " (" & sFieldList & ") VALUES (" & sValueList & ")"
				'DEBUG DATA: response.write sInsertStatement & "<BR>"
				
				' INSERT NEW ROW INTO DATABASE AND GET ROWID
				sSQL = "SET NOCOUNT ON;" &_
				sInsertStatement &_
				"SELECT @@IDENTITY AS ROWID;"
				Set oInsert = Server.CreateObject("ADODB.Recordset")
				oInsert.Open sSQL, Application("DSN"), 3, 1
				iReturnValue = oInsert("ROWID")
				response.write "(" & iReturnValue & ")"
				
				oSchema.MoveNext
			End If
		Loop

	End If

	oSchema.close
	Set oSchema = Nothing 

	fnCopySQLQuestionRow = iReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------------------------------------
' SUB SUBASSIGNFORMTOCATEGORY(IFORMID,iCategoryID)
'------------------------------------------------------------------------------------------------------------
Sub subAssignFormtoCategory(iFormID,iCategoryID, iOrgId)
	' INSERT NEW 
	sSQL = "INSERT INTO egov_forms_to_categories (form_category_id,action_form_id, orgid) VALUES ('" & iCategoryID & "','" & iFormID & "','" & iOrgId & "')"
	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 1

	' INSERT NEW
	sSQL = "INSERT INTO egov_organizations_to_forms (orgid,action_form_id,action_form_enabled) VALUES ('" & iOrgId & "','" & iFormID & "','1')"
	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 1
End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION GETCATEGORYID(IFORMID)
'------------------------------------------------------------------------------------------------------------
Function GetCategoryID(iFormID)

	iReturnValue = 0

	sSQL = "Select * FROM egov_forms_to_categories WHERE action_form_id='" & iFormID & "'"
	Set oCategoryID = Server.CreateObject("ADODB.Recordset")
	oCategoryID.Open sSQL, Application("DSN"), 3, 2
	
	If NOT oCategoryID.EOF Then
		iReturnValue = oCategoryID("form_category_id")
	End If

	oCategoryID.close
	Set oCategoryID = Nothing 

	GetCategoryID = iReturnValue

End Function
%>