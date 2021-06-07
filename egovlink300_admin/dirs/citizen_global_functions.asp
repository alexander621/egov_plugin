<%
'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub buildRequiredFieldLabel(iLabel)

  lcl_label = ""
  lcl_label = lcl_label & "<span class=""cot-text-emphasized"" title=""This field is required"">"
  lcl_label = lcl_label & "<span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>"
  lcl_label = lcl_label & "&nbsp;" & iLabel
  lcl_label = lcl_label & "</span>"

  response.write lcl_label & vbcrlf

end sub

'------------------------------------------------------------------------------
function getCommitteeEmails(iOrgID, iCommitteeID)
  lcl_return    = ""
  lcl_emaillist = ""

  if iCommitteeID <> "" then

    	sSQL = "SELECT useremail "
     sSQL = sSQL & " FROM egov_users u "
     sSQL = sSQL &      " INNER JOIN vwCitizenGroups ug ON u.userid = ug.citizenid "
     sSQL = sSQL & " WHERE ug.groupid=" & iCommitteeID
     sSQL = sSQL & " AND u.orgid=" & iOrgID

    	set oGetCitizenEmails = Server.CreateObject("ADODB.Recordset")
    	oGetCitizenEmails.Open sSQL, Application("DSN"), 3, 1

     if not oGetCitizenEmails.eof then
        do while not oGetCitizenEmails.eof
      	   	if trim(oGetCitizenEmails("useremail")) <> "" then
              lcl_emaillist = lcl_emaillist & oGetCitizenEmails("useremail") & ";"
           end if

           oGetCitizenEmails.movenext
        loop
     end if

     oGetCitizenEmails.close
     set oGetCitizenEmails = nothing

     if lcl_emaillist <> "" then
        lcl_emaillist = "mailto:" & lcl_emaillist
        lcl_return    = lcl_emaillist
     end if

  end if

 	getCommitteeEmails = lcl_return

end function

'------------------------------------------------------------------------------
function getTotalUsersInGroup(iOrgID, iGroupID, iView, iOrgHasFeature_HasFamily)
  lcl_return       = 0
  lcl_where_clause = ""

  if iOrgID <> "" AND iGroupID <> "" then
     if iView = "" then
        if not iOrgHasFeature_HasFamily then
           lcl_where_clause = " AND headofhousehold = 1 "
        end if
     elseif clng(iView) = clng(2) then
        lcl_where_clause = lcl_where_clause & " AND headofhousehold = 1 "
        lcl_where_clause = lcl_where_clause & " AND userlname IS NOT NULL "
     elseif clng(iView) = clng(3) then
       	lcl_where_clause = lcl_where_clause & " AND u.residencyverified = 0 "
        lcl_where_clause = lcl_where_clause & " AND residenttype = 'R' "
        lcl_where_clause = lcl_where_clause & " AND headofhousehold = 1 "
        lcl_where_clause = lcl_where_clause & " AND userlname IS NOT NULL "
     elseif clng(iView) = clng(4) then
       	lcl_where_clause = lcl_where_clause & " AND u.registrationblocked = 1 "
        lcl_where_clause = lcl_where_clause & " AND headofhousehold = 1 "
        lcl_where_clause = lcl_where_clause & " AND userlname IS NOT NULL "
     elseif clng(iView) = clng(5) then
       	lcl_where_clause = lcl_where_clause & " AND userfname IS NULL "
        lcl_where_clause = lcl_where_clause & " AND userlname IS NULL "
        lcl_where_clause = lcl_where_clause & " AND headofhousehold = 1 "
     end if


     sSQL = "SELECT count(u.userid) as total_users "
     sSQL = sSQL & " FROM egov_users u "
     sSQL = sSQL &      " INNER JOIN vwcitizengroups ug ON u.userid=ug.citizenid "
     sSQL = sSQL & " WHERE u.orgid = " & iOrgID
     sSQL = sSQL & lcl_where_clause
     sSQL = sSQL & " AND isdeleted = 0 "
     sSQL = sSQL & " AND userregistered = '1' "
     sSQL = sSQL & " AND ug.groupid = " & iGroupID

    	set oCountUsers = Server.CreateObject("ADODB.Recordset")
    	oCountUsers.Open sSQL, Application("DSN"), 3, 1

     if not oCountUsers.eof then
        lcl_return = oCountUsers("total_users")
     end if

     oCountUsers.close
     set oCountUsers = nothing

  end if

  getTotalUsersInGroup = lcl_return

end Function


'------------------------------------------------------------------------------
' void DisplayGenderPicks sElement, sGender 
'------------------------------------------------------------------------------
Sub DisplayGenderPicks( ByVal sElement, ByVal sGenderMatch )
	Dim sSql, oRs

	sSql = "SELECT gender, genderdescription FROM egov_user_genders ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""" & sElement & """ name=""" & sElement & """>"
		response.write vbcrlf & "<option value=""N"">Select a gender...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("gender") & """"
			If sGenderMatch = oRs("gender") Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("genderdescription") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		' this should never happen
		response.write vbcrlf & "<input type=""hidden"" id=""" & sElement & """ name=""" & sElement & """ value=""N"" />"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' string GetHeadOfHouseholdName( iFamilyId )
'------------------------------------------------------------------------------
Function GetHeadOfHouseholdName( ByVal iFamilyId )
	Dim sSql, oRs

	sSql = "SELECT userfname + ' ' + userlname AS familyname FROM egov_users WHERE headofhousehold = 1 AND familyid = " & iFamilyId 
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetHeadOfHouseholdName = oRs("familyname")
		'response.write oRs("familyname") & "<br /><br />"
	Else
		GetHeadOfHouseholdName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>
