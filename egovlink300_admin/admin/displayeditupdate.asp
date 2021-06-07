<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: displayeditupdate.asp
' AUTHOR: Steve Loar
' CREATED: 05/18/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates Displays. Called from displayedit.asp
'
' MODIFICATION HISTORY
' 1.0   05/18/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iDisplayId, sDisplayName, sDisplay, sDisplayDescription, iFeatureId, sisOnPublicSide
Dim sisOnAdminSide, sAdminCanEdit, sSql, sMsg, sUsesDisplayName

If request("displayid") <> "" Then
	iDisplayId = CLng(request("displayid"))
Else
	response.redirect "displaylist.asp"
End If 

sDisplay = "'" & dbsafe(request("display")) & "'"
sDisplayName = "'" & dbsafewithHTML(request("displayname")) & "'"
sDisplayDescription = "'" & dbsafewithHTML(request("displaydescription")) & "'"

If CLng(request("featureid")) > CLng(0) Then
	iFeatureId = request("featureid")
Else
	iFeatureId = "NULL"
End If 

If LCase(request("admincanedit")) = "on" Then
	sAdminCanEdit = "1"
Else
	sAdminCanEdit = "0"
End If 

If LCase(request("isonpublicside")) = "on" Then
	sisOnPublicSide = "1"
Else
	sisOnPublicSide = "0"
End If 

If LCase(request("isonadminside")) = "on" Then
	sisOnAdminSide = "1"
Else
	sisOnAdminSide = "0"
End If 

If LCase(request("usesdisplayname")) = "on" Then
	sUsesDisplayName = "1"
Else
	sUsesDisplayName = "0"
End If 

If iDisplayId = CLng(0) Then 
	sSql = "INSERT INTO egov_organization_displays ( display, displayname, displaydescription, "
	sSql = sSql & "featureid, admincanedit, isonpublicside, isonadminside, usesdisplayname ) VALUES ( "
	sSql = sSql & sDisplay & ", " & sDisplayName & ", " & sDisplayDescription & ", "
	sSql = sSql & iFeatureId & ", " & sAdminCanEdit & ", " & sisOnPublicSide & ", "
	sSql = sSql & sisOnAdminSide & ", " & sUsesDisplayName & " )"
	'response.write sSql

	iDisplayId = RunInsertStatement( sSql )
	sMsg = "n"
Else
	sSql = "UPDATE egov_organization_displays SET "
	sSql = sSql & "display = " & sDisplay & ", "
	sSql = sSql & "displayname = " & sDisplayName & ", "
	sSql = sSql & "displaydescription = " & sDisplayDescription & ", "
	sSql = sSql & "featureid = " & iFeatureId & ", "
	sSql = sSql & "admincanedit = " & sAdminCanEdit & ", "
	sSql = sSql & "isonpublicside = " & sisOnPublicSide & ", "
	sSql = sSql & "isonadminside = " & sisOnAdminSide & ", "
	sSql = sSql & "usesdisplayname = " & sUsesDisplayName
	sSql = sSql & " WHERE displayid = " & iDisplayId

	RunSQLStatement sSql
	sMsg = "u"
End If 

response.redirect "displayedit.asp?displayid=" & iDisplayId & "&msg=" & sMsg

%>

