<%
if request("type") = "feetypes" then ShowFeeTypePicks
if request("type") = "feemultipliers" then ShowUnusedMultipliers( request("value") )
if request("type") = "fixtures" then ShowUnusedFixtures( )
if request("type") = "reviewtypes" then ShowReviewTypePicks
if request("type") = "inspectiontypes" then ShowInspectionTypePicks
if request("type") = "customfieldtypes" then ShowCustomFieldTypePicks

Sub ShowFeeTypePicks()
	Dim sSql, oRs

	sSql = "SELECT F.permitfeetypeid, F.permitfee, F.isupfrontfee, F.isreinspectionfee, M.permitfeemethod "
	sSql = sSql & " FROM egov_permitfeetypes F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.orgid = " & SESSION("orgid") & " ORDER BY F.permitfee"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write "<option value=""0"">Select a Fee Type</option>"
		Do While NOT oRs.EOF 
			response.write "<option value=""" & oRs("permitfeetypeid") & """>" 
				response.write oRs("permitfee") & " (" & Left(oRs("permitfeemethod"),23) & ")"
			response.write "</option>"
			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing

End Sub 
Sub ShowUnusedMultipliers( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT feemultiplier, feemultipliertypeid FROM egov_feemultipliertypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND feemultipliertypeid NOT IN (SELECT feemultipliertypeid "
	sSql = sSql & " FROM egov_permitfeetypes_to_feemultipliertypes WHERE permitfeetypeid = '" & replace(iPermitFeeTypeid,"'","''") & "'"
	sSql = sSql & " ) ORDER BY feemultiplier"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("feemultipliertypeid") & """>" & oRs("feemultiplier") & "</option>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 
End Sub 
Sub ShowUnusedFixtures( )
	Dim sSql, oRs

	sSql = "SELECT F.permitfixture, F.permitfixturetypeid FROM egov_permitfixturetypes F "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY F.displayorder, F.permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitfixturetypeid") & """>" & oRs("permitfixture") & "</option>"
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 
End Sub 
Sub ShowReviewTypePicks()
	Dim sSql, oRs

	sSQL = "SELECT permitreviewtypeid, permitreviewtype "
	sSql = sSql & " FROM egov_permitreviewtypes WHERE orgid = " & SESSION("orgid") & " ORDER BY permitreviewtype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrfl & "<option value=""0"">Select a Review Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitreviewtypeid") & """"  
			response.write ">" & oRs("permitreviewtype")
			response.write "</option>"
			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  
Sub ShowInspectionTypePicks()
	Dim sSql, oRs

	sSql = "SELECT permitinspectiontypeid, permitinspectiontype "
	sSql = sSql & " FROM egov_permitinspectiontypes WHERE orgid = " & SESSION("orgid") & " ORDER BY permitinspectiontype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrfl & "<option value=""0"">Select an Inspection Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("permitinspectiontypeid") & """"  
			response.write ">" & oRs("permitinspectiontype")
			response.write "</option>"
			oRs.MoveNext
		Loop

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  
Sub ShowCustomFieldTypePicks( )
	Dim sSql, oRs

	sSQL = "SELECT customfieldtypeid, fieldname "
	sSql = sSql & " FROM egov_permitcustomfieldtypes WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " AND isactive = 1 ORDER BY fieldname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		response.write vbcrfl & "<option value=""0"">Select a Custom Field</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("customfieldtypeid") & """"  
			response.write ">" & oRs("fieldname")
			response.write "</option>"
			oRs.MoveNext
		Loop


	End If

	oRs.Close
	Set oRs = Nothing

End Sub 
%>
