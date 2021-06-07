<%
Dim iOrgid, iCitizenId, sDSN, iOtherId, iRequestId, iRepairId, iNuisanceId, iAnimal, iLicenseId, iFormCategoryId

iOrgid = "66"   ' Change to match new orgid

'response.write Now() & "<br /><br />"
response.write "<div style=""background-color:#e0e0e0;border: solid 1px #000000;padding:10px;FONT-FAMILY: Verdana,Tahoma,Arial;font-size:10px;"">"
response.write "<p><b>Adding the Seven Categories for orgid(" & iorgid & ")...</b></p>"
response.write Now() & "<br /><br />"

sDSN = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;"

iCitizenId = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Citizen Comments and Concerns',1," & iOrgid & ")")
response.write "Citizen Comments and Concerns &ndash; " & iCitizenId & "<br/>"
response.flush

iRequestId = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Requests for Information',2," & iOrgid & ")")
response.write "Requests for Information &ndash; " & iRequestId & "<br/>"
response.flush

iRepairId = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Repairs and Requests for Service',3," & iOrgid & ")")
response.write "Repairs and Requests for Service &ndash; " & iRepairId & "<br/>"
response.flush

iNuisanceId = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Nuisance/Code Violations',4," & iOrgid & ")")
response.write "Nuisance/Code Violations &ndash; " & iNuisanceId & "<br/>"
response.flush

iAnimal = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Animal Control',5," & iOrgid & ")")
response.write "Animal Control &ndash; " & iAnimal & "<br/>"
response.flush

iLicenseId = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Licenses',6," & iOrgid & ")")
response.write "Licenses &ndash; " & iLicenseId & "<br/>"
response.flush

iOtherId = RunSQL("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Other',7," & iOrgid & ")")
response.write "Other &ndash; " & iOtherId & "<br/>"
response.flush

response.write "<p><b>Done Adding the Seven Categories.</b></p>"
response.write Now()
response.flush

response.write "<p><b>Updating the Form Categories.</b></p>"
' Update the egov_form_categories table with the new catagories

sSQL = "Select egov_rowid, form_category_id, form_category_name "
sSQL = sSQL & "from egov_forms_categories_view where orgid = " & iorgid

Set oData = Server.CreateObject("ADODB.Recordset")
oData.Open sSQL, Application("DSN"), 0, 1

If Not oData.EOF Then 
	Do While Not oData.EOF
		Select Case oData("form_category_name")
			Case "Citizen Comments and Concerns"
				iFormCategoryId = iCitizenId
			Case "Requests for Information"
				iFormCategoryId = iRequestId
			Case "Repairs and Requests for Service"
				iFormCategoryId = iRepairId
			Case "Nuisance/Code Violations"
				iFormCategoryId = iNuisanceId
			Case "Animal Control"
				iFormCategoryId = iAnimal
			Case "Licenses"
				iFormCategoryId = iLicenseId
			Case "Other"
				iFormCategoryId = iOtherId
			Case Else 
				iFormCategoryId = iOtherId
		End Select 

		response.write "(" & oData("egov_rowid") & ") " & oData("form_category_name") & " &ndash; " & iFormCategoryId & "<br />"
		UpdateCategoryId oData("egov_rowid"), iFormCategoryId  
		oData.movenext
	Loop 
Else
	response.write "<h2>Error &ndash; No forms were found to place in the categories.</h2>Clear out the categories, edit and run copy_all_ques.asp, then run this script again.<br /><br />"
End If 

oData.close
Set oData = Nothing

response.write "<p><b>Done Updating the Form Categories.</b></p>"
response.write Now()
response.write "</div>"

'-------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'-------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' RunSQL(sInsertStatement)
'-------------------------------------------------------------------------------------------------
Function RunSQL(sInsertStatement)
	Dim sSQL
	RunSQL = 0

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	
	oInsert.Open sSQL, sDSN, 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.close
	Set oInsert = Nothing

	response.write sSQL & "<br /><br />"

	RunSQL = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' UpdateCategoryId(iEgovRowId, iFormCategoryId )
'-------------------------------------------------------------------------------------------------
Sub  UpdateCategoryId(iEgovRowId, iFormCategoryId )
	Dim sSql, oCmd

	' update the formcategoryid
	sSQL = "UPDATE  egov_forms_to_categories SET form_category_id = " & iFormCategoryId
	sSQL = sSQL & " WHERE egov_rowid =" & iEgovRowId  

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub 

%>
