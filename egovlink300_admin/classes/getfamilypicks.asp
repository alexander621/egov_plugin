<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getfamilypicks.asp
' AUTHOR: Steve Loar
' CREATED: 03/25/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the family members drop down using a userid of the head of household. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   03/25/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sSql, oRs, sResults, sMember, iCount, iMonths, iAge, iMembershipId, lcl_member_selected

iUserId = CLng(0)
iMembershipId = CLng(0)

If request("egovuserid") <> "" Then 
	If IsNumeric( request("egovuserid") ) Then 
		iUserid = CLng("0" & request("egovuserid"))
	End If 
Else
	If session("eGovUserId") <> "" Then 
		If IsNumeric( Session("eGovUserId") ) Then 
			iUserid = CLng( "0" & Session("eGovUserId"))
		End If 
	End If 
End If 

If IsNull(iUserid) Or iUserid = CLng(0) Then
	iUserid = CLng(-1)
End If 

If request("membershipid") <> "" Then 
	iMembershipId = CLng(request("membershipid"))
End If 

iCount = 0
iMemberCount = 0 

sSql = "SELECT firstname, lastname, familymemberid, relationship, birthdate, userid"
sSql = sSql & " FROM egov_familymembers "
sSql = sSql & " WHERE isdeleted = 0 AND belongstouserid = " & iUserid
sSql = sSql & " ORDER BY birthdate ASC"
session("getfamilypicks") = sSql

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	sResults = "<select name=""familymemberid"" id=""familymemberid"">" 

	Do While Not oRs.EOF
		If CLng(iMembershipId) > CLng(0) Then 
			sMember = DetermineMembership(oRs("familymemberid"), iUserid, iMembershipId)   ' In class_global_functions.asp
		End If 

		If iCount = 0 Then 
			lcl_member_selected = " selected=""selected"""
		Else 
			lcl_member_selected = ""
		End If 

		sResults = sResults & "<option value=""" & oRs("familymemberid") & """" & lcl_member_selected & ">" & oRs("firstname") & " " & oRs("lastname") & " &ndash; " 

		If CLng(oRs("userid")) = CLng(iUserid) Then
			sResults = sResults & "Head of Household"
		Else 
			sResults = sResults & oRs("relationship") 
		End If 

		If UCase(oRs("relationship")) = "CHILD" Then 
			iAge = GetChildAge(oRs("birthdate"))
			sResults = sResults & " &ndash; Age: " & iAge & " yrs"
		Else
			If UCase(oRs("relationship")) <> "SITTER" Then
				sResults = sResults & " &ndash; Adult"
			End If 
		End If 

		If CLng(iMembershipId) > CLng(0) Then 
			If sMember = "M" Then
				sResults = sResults & " &ndash; Member" 
				iMemberCount = iMemberCount + 1
			Else 
				sResults = sResults & " &ndash; NonMember"
			End If 
		End If 

		sResults = sResults & "</option>" 
		iCount = iCount + 1

		oRs.MoveNext
	Loop 
	sResults = sResults & "</select>" 
Else
	sResults = "<input type='hidden' name='familymemberid' id='familymemberid' value='0' /><span class=""nomatch"">No One Found for this User.</span>"
End If 

oRs.Close
Set oRs = Nothing

response.write sResults

%>