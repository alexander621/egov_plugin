<%
Call subCopyForm(request("iformid"),request("iorgid"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBCOPYFORM(IFORMID,IORGID)
'--------------------------------------------------------------------------------------------------
Sub subCopyForm(iFormID,iOrgID)
	
	' COPY MAIN FORM DATA ROW 
	iNewFormID = fnCopySQLDataRow("action_form_id",iFormID,"egov_action_request_forms")

	' COPY FORM QUESTION DATA ROWS
	Call SubCopyFormQuestions(iFormID,iOrgID,iNewFormID)

	' GET CATEGORY VALUE FOR "OTHER" PER ORGANIZATION
	iCategoryID = fnGetCategoryID(iOrgID)

	'CODE TO ASSIGN DEFAULT CATEGORY
	If request("task")="NEW" OR UCASE(request("task")) = "COPYME" Then
		'The "Other" Category value is hard-coded to the original value.
		Call subAssignFormtoCategory(iNewFormID,iCategoryID,iorgid)'iCategoryID,iorgid)
	End If

	' REDIRECT TO MANANGE FORM PAGE
	response.redirect("manage_form.asp?iformid=" & iNewFormID)

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNCOPYSQLDATAROW(SPRIMARYKEY,IPRIMARYKEYID,STABLENAME)
'--------------------------------------------------------------------------------------------------
Function fnCopySQLDataRow(sPrimaryKey,iPrimaryKeyID,sTableName)

	iReturnValue = 0 

	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * FROM " & sTableName & " WHERE " & sPrimaryKey & "='" & iPrimaryKeyID & "'"
	oSchema.Open sSQL, Application("DSN") , 3, 1
	
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
						Select Case fldLoop.Name

							Case "action_form_name"
								
								if request("type") = "market" then
									If fldLoop.Type = 11 Then
										sValueList = sValueList & "'" & fnBitConvert(fldLoop.Value) & "'"
									Else
										sValueList = sValueList & "'" & DBSafe(fldLoop.Value) & "'"
									End If
								else
									sValueList = sValueList & "'NEW ACTION FORM'"
								end if

							Case "orgid"
								
								sValueList = sValueList & "'" & request("iorgid") & "'"
							
							Case "action_form_enabled"
								
								' DEFAULT TO OFF
								sValueList = sValueList & "'0'"

							Case Else

								' ADD DATA TO STRING
								If fldLoop.Type = 11 Then
									sValueList = sValueList & "'" & fnBitConvert(fldLoop.Value) & "'"
								Else
									sValueList = sValueList & "'" & DBSafe(trim(fldLoop.Value)) & "'"
								End If

						End Select 

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
				oInsert.Open sSQL, Application("DSN") , 3, 1
				iReturnValue = oInsert("ROWID")
				response.write "(" & iReturnValue & ")"
				Set oInsert = Nothing
				
				oSchema.MoveNext
			End If
		Loop

	End If

	fnCopySQLDataRow = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' SUB SUBCOPYFORMQUESTIONS(IFORMID,IORGID,INEWFORMID)
'--------------------------------------------------------------------------------------------------
Sub SubCopyFormQuestions(iFormID,iOrgID,iNewFormID)

	sSQL = "Select * FROM egov_action_form_questions WHERE formid='" & iFormID & "' ORDER BY SEQUENCE"
	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN") , 3, 2
	
	If NOT oQuestions.EOF Then
		Do While NOT oQuestions.EOF 
			iTempReturnID = fnCopySQLQuestionRow("questionid",oQuestions("questionid"),"egov_action_form_questions",iNewFormID,iOrgID)
			oQuestions.MoveNext
		Loop 
	End If

	oQuestions.close
	Set oQuestions = nothing

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
	oSchema.Open sSQL, Application("DSN") , 3, 1
	
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
				oInsert.Open sSQL, Application("DSN") , 3, 1
				iReturnValue = oInsert("ROWID")
				response.write "(" & iReturnValue & ")"
				Set oInsert = Nothing
				
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
Sub subAssignFormtoCategory(iFormID,iCategoryID,iorgid)
	' INSERT NEW 
	sSQL = "INSERT INTO egov_forms_to_categories (form_category_id,action_form_id,orgid) VALUES ('" & iCategoryID & "','" & iFormID & "','" & iorgid & "')"
	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN") , 3, 1
	Set oInsert = Nothing

	' INSERT NEW
	sSQL = "INSERT INTO egov_organizations_to_forms (orgid,action_form_id,action_form_enabled) VALUES ('" & iorgid & "','" & iFormID & "','1')"
	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN") , 3, 1
	Set oInsert = Nothing
End Sub

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNGETCATEGORYID(IFORMID,IORGID)
' 9/7/2005 - Vincent Evans
'------------------------------------------------------------------------------------------------------------
Function fnGetCategoryID(iorgid)

	'Select category id to update form category on copy with corresponding 'other' category.
	sSQL = "SELECT form_category_id FROM egov_form_categories Where orgid=" & iorgid & " and form_category_name = 'Other'"

	Set oSelect = Server.CreateObject("ADODB.Recordset")
	oSelect.Open sSQL, Application("DSN"), 3, 1
	
	If Not oSelect.EOF Then
		iTemp = clng(oSelect("form_category_id"))
		fnGetCategoryID = iTemp
	End If

	oSelect.close
	Set oSelect = Nothing

End Function
%>
