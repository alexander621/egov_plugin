<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
iContactTypeId = request("iContactTypeId")

	sSql = "SELECT permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSQl & "ISNULL(company,'') AS company, ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, "
	sSql = sSql & "ISNULL(contractortypeid,0) AS contractortypeid "
	sSql = sSQl & "FROM egov_permitcontacttypes "
	sSql = sSql & "WHERE isorganization = 0 AND orgid = " & session("orgid") & " "
	if request("query") <> "" then
		sSql = sSql & " AND ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') LIKE '%" & DBSafe(request("query")) & "%' "
	end if
	if iContactTypeId <> "" then
		sSql = sSQL & " AND permitcontacttypeid = '" & DBSafe(iContactTypeId) & "'"
	end if
	sSql = sSql & "ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	if oRs.EOF and request("query") <> ""  then response.write "<option value=""0"">No match for your search</option>"
	if oRs.EOF and request("query") = ""  then response.write "<option value=""0"">Search for a contractor below.</option>"
	if not oRs.EOF and request("query") <> ""  then response.write "<option value=""0"">Select a contractor here.</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitcontacttypeid") & """"
		If CLng(iContactTypeId) = CLng(oRs("permitcontacttypeid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">"
		If oRs("company") <> "" Then
			response.write oRs("company")
			If oRs("firstname") <> "" Then 
				response.write " &mdash; "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			response.write oRs("lastname") & ", " & oRs("firstname")
		End If 

		' Contractor type
		sContractorType = GetContractorType( oRs("contractortypeid") ) 
		If sContractorType <> "" Then 
			response.write " &ndash; (<strong>" & sContractorType & "</strong>)"
		End If 

		response.write "</option>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

%>
