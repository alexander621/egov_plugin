<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: mobilefeaturesetupsave.asp
' AUTHOR: Steve Loar
' CREATED: 04/15/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is where mobile feature setup options are saved
'
' MODIFICATION HISTORY
' 1.0   04/15/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sOrgId, iMaxFeatureCount

sOrgId = CLng(request("orgid"))

iMaxFeatureCount = CLng(request("maxfeaturecount"))

If iMaxFeatureCount > CLng(0) Then 
	
	For x = 1 To iMaxFeatureCount
		sSql = "UPDATE egov_organizations_to_features SET "
		' stuff here

		' mobileisactivated
		If request("mobileisactivated" & x) = "on" Then
			sSql = sSql & "mobileisactivated = 1, "
		Else 
			sSql = sSql & "mobileisactivated = 0, "
		End If 

		' mobilename
		If request("mobilename" & x) <> "" Then
			sSql = sSql & "mobilename = '" & DBsafe(request("mobilename" & x)) & "', "
		Else
			sSql = sSql & "mobilename = NULL, "
		End If 

		' mobiledisplayorder
		'response.write "request(""mobiledisplayorder" & x & """) = " & request("mobiledisplayorder" & x) & "<br />"
		If request("mobiledisplayorder" & x) <> "" Then
			sSql = sSql & "mobiledisplayorder = " & request("mobiledisplayorder" & x) & ", "
		Else
			sSql = sSql & "mobiledisplayorder = NULL, "
		End If 

		' mobileitemcount
		If request("mobileitemcount" & x) <> "" Then
			sSql = sSql & "mobileitemcount = " & request("mobileitemcount" & x) & ", "
		Else
			sSql = sSql & "mobileitemcount = NULL, "
		End If 

		' mobilelistcount
		If request("mobilelistcount" & x) <> "" Then
			sSql = sSql & "mobilelistcount = " & request("mobilelistcount" & x) & " "
		Else
			sSql = sSql & "mobilelistcount = NULL "
		End If 

		sSql = sSql & " WHERE orgid = " & sOrgId & " AND featureid = " & CLng(request("featureid" & x))
		'response.write sSql & "<br /><br />"

		RunSQLStatement sSql
	Next 

End If 


' Take them back To the mobile feature setup page
response.redirect "mobilefeaturesetup.asp?s=upd&orgid=" & sOrgId

%>