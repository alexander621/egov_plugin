<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/JSON_2.0.2.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getcitizendetails.asp
' AUTHOR: Steve Loar
' CREATED: 12/23/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Pulls the details for citizens for the citizenmerge feature. Called via AJAX.
'
' MODIFICATION HISTORY
' 1.0   12/23/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sResponse, sSql, oRs, iFamilyId, sType, sAddress, bFound, sFamily, iRecCount

iFamilyId = CLng(request("familyid"))
sType = request("sType")
sAddress = ""
iRecCount = 0

' Create the JSON object to pass data back to the calling page
Set sResponse = jsObject()

sSql = "SELECT userid, useremail, ISNULL(useraddress,'') AS useraddress, ISNULL(usercity,'') AS usercity, "
sSql = sSql & " ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, userhomephone "
sSql = sSql & " FROM egov_users WHERE headofhousehold = 1 AND isdeleted = 0 AND familyid = " & iFamilyId
'response.write sSql & "<br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	' Format the JSON return
	sResponse("flag") = "success"
	sResponse("type") = sType
	'iFamilyId = oRs("familyid")
	sResponse("familyid") = "Familyid: " & iFamilyId & "<br /><br />"
	sResponse("userid") = "<input type=""hidden"" name=""" & sType & "userid1"" value=""" & oRs("userid") & """ />"
	sResponse("useremail") = "Email: " & oRs("useremail") & "<br /><br />"
	If oRs("useraddress") <> "" Then
		sAddress = sAddress & oRs("useraddress") & "<br />"
	End If 
	If oRs("usercity") <> "" Then 
		sAddress = sAddress & oRs("usercity") 
	End If 
	If oRs("usercity") <> "" And oRs("userstate") <> "" Then 
		sAddress = sAddress & ", "
	End If 
	If oRs("userstate") <> "" Then 
		sAddress = sAddress & oRs("userstate") & " " & oRs("userzip")
	End If 
	sResponse("useraddress") = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""addresstable""><tr><td valign=""top"" class=""addresslabel"">Address: </td><td valign=""top"">" & sAddress & "</td></tr></table>"
	sResponse("userhomephone") = "Phone: " & FormatPhoneNumber(oRs("userhomephone")) & "<br />"
	bFound = True 
Else
	' Format the JSON return
	sResponse("flag") = "failed"
	sResponse("type") = sType
	bFound = False 
End If

oRs.Close
Set oRs = Nothing 

If bFound Then 
	' Get some family information to show
	sSql = "SELECT U.userid, ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname, U.headofhousehold, ISNULL(F.relationship,'') AS relationship, U.birthdate "
	sSql = sSql & " FROM egov_users U, egov_familymembers F WHERE U.familyid = " & iFamilyId
	sSql = sSql & " AND U.isdeleted = 0 AND U.userid = F.userid AND F.isdeleted = 0 "
	sSql = sSql & " ORDER BY U.headofhousehold DESC, U.userfname, U.userlname"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sFamily = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""familytable""><tr><th colspan=""3"">Family Members</th></tr>"
		Do While Not oRs.EOF
			iRecCount = iRecCount + 1
			sFamily = sFamily & "<tr"
			If iRecCount Mod 2 = 0 Then
				sFamily = sFamily & " class=""altrow"" "
			End If 
			sFamily = sFamily & "><td align=""center"">" & oRs("userfname") & " " & oRs("userlname") & "</td><td align=""center"">"
			If oRs("headofhousehold") Then
				sFamily = sFamily & "Head of Household"
			Else
				sFamily = sFamily & oRs("relationship")
			End If 
			sFamily = sFamily & "</td><td align=""center"">"
			If LCase(oRs("relationship")) = "child" And Not IsNull(oRs("birthdate")) Then
				sFamily = sFamily & FormatNumber(GetAgeOnDate( oRs("birthdate"), Now ),1) & " yrs"
			Else 
				sFamily = sFamily & "&nbsp;"
			End If 
			sFamily = sFamily & "</td></tr>"
			oRs.MoveNext
		Loop 
		sFamily = sFamily & "</table>"
		'response.write "<br />Family: " & sFamily & "<br /><br />"
		sResponse("family") = sFamily
		sResponse("mergecount") = "<input type=""hidden"" name=""max" & sType & """ value=""" & iRecCount & """ />"
	Else 
		sResponse("family") = "<p>No family members found.</p>"
		sResponse("mergecount") = "<input type=""hidden"" name=""max" & sType & """ value=""0"" />"
	End If 

	oRs.Close
	Set oRs = Nothing 
End If 

sResponse.Flush



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetAgeOnDate( dBirthDate, dCompareDate )
'--------------------------------------------------------------------------------------------------
Function GetAgeOnDate( dBirthDate, dCompareDate )
	Dim iMonths, iAge

	iAge = (DateValue(dCompareDate) - DateValue(dBirthDate)) / 365.25
	GetAgeOnDate = iAge

End Function 

%>