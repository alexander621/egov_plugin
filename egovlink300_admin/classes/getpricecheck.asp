<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getpricecheck.asp
' AUTHOR: Steve Loar
' CREATED: 04/12/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the pricetypeid that should be checked for a class purchase. 
'				It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   04/12/2011   Steve Loar - INITIAL VERSION
' 1.1	10/25/2012	Steve Loar - Modified the setting of the userid to help where the userid is passed as null
'								 and this would crash.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserid, sSql, oRs, sResults, iClassid, iMembershipid, sResidentType, sMemberType,	iMemberCount

sResults = ""
iUserId = CLng(0)

If request("userid") <> "" Then 
	If IsNumeric( request("userid") ) Then 
		iUserid = CLng("" & request("userid"))
	End If 
Else
	If session("eGovUserId") <> "" Then 
		If IsNumeric( Session("eGovUserId") ) Then 
			iUserid = Session("eGovUserId")
		End If 
	End If 
End If 

sResidentType = GetUserResidentType( iUserid )		' in class_global_functions.asp

iClassid = CLng(request("classid"))

iMembershipid = CLng(request("membershipid"))

If iMembershipid > CLng(0) Then 
	' Get the count of family member that are in the membership of the class
	iMemberCount = GetMemberCount( iMembershipid, iUserid )
Else
	iMemberCount = 0
End If 

' if at least one person in the family is a member, then set up for member pricing match
If clng(iMemberCount) > clng(0) Then 
	sMemberType = "M"
Else
	sMemberType = "O"
End If 

sSql = "SELECT P.pricetypeid, T.pricetypename, T.ismember, T.pricetype, "
sSql = sSql & "T.isfee, T.isbaseprice, T.checkmembership, ISNULL(P.membershipid,0) AS membershipid, T.isdropin "
sSql = sSql & " FROM egov_price_types T, egov_class_pricetype_price P "
sSql = sSql & " WHERE T.pricetypeid = P.pricetypeid "
sSql = sSql & " AND orgid = " & session("orgid") & " AND P.classid = " & iClassid & " ORDER BY P.pricetypeid"
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

Do While Not oRs.EOF
	If oRs("isfee") Then 
		'Always check a fee
		sResults = sResults & oRs("pricetypeid") & ","
	Else 
		If oRs("isbaseprice") Then 
			'always check a base price
			sResults = sResults & oRs("pricetypeid") & ","
		Else
			If oRs("pricetype") = sResidentType Then 
				'if the resident type requirement matches
				sResults = sResults & oRs("pricetypeid") & ","
			Else
				If oRs("checkmembership") Then
					' If there is membership, see it they are the kind of member or non-member of this 
					If oRs("pricetype") = sMemberType Then 
						sResults = sResults & oRs("pricetypeid") & ","
					End If 
				End If 
			End If 
		End If 
	End If 
	oRs.MoveNext
Loop

oRs.Close
Set oRs = Nothing 

If sResults <> "" Then 
	' strip off the last comma before sending back
	sResults = Left(sResults, (Len(sResults) -1))
	sResults = Trim(sResults)
Else
	sResults = "none"
End If 

response.write sResults



'------------------------------------------------------------------------------
' integer GetMemberCount( iMembershipid, iUserid )
'------------------------------------------------------------------------------
Function GetMemberCount( ByVal iMembershipid, ByVal iUserid )
	Dim sSql, oRs, sMembership

	GetMemberCount = 0

	sSql = "SELECT familymemberid"
	sSql = sSql & " FROM egov_familymembers "
	sSql = sSql & " WHERE isdeleted = 0 AND belongstouserid = " & iUserid
	sSql = sSql & " ORDER BY birthdate ASC"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		sMembership = DetermineMembership( oRs("familymemberid"), iUserid, iMembershipId )	' in class_global_functions.asp

		If sMembership = "M" Then
			GetMemberCount = 1
			Exit Do 
		End If 

		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Function 



%>
