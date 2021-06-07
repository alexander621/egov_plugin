<!-- menu script itself. you should not modify this file -->
<script language="JavaScript" src="<%=sLevel%>menu/menu_scripts/menu.js"></script>
<!-- items structure. menu hierarchy and links are stored there -->

	<script>
		<% DisplayFeatureListNav sLevel %>
	</script>


<!-- files with geometry and styles structures -->
<script language="JavaScript" src="<%=sLevel%>menu/menu_scripts/menu_tpl.js"></script>
<script language="JavaScript">
	<!--//
	// Note where menu initialization block is located in HTML document.
	// Don't try to position menu locating menu initialization block in
	// some table cell or other HTML element. Always put it before </body>

	// each menu gets two parameters (see demo files)
	// 1. items structure
	// 2. geometry structure

	new menu (MENU_ITEMS, MENU_POS);
	// make sure files containing definitions for these variables are linked to the document
	// if you got some javascript error like "MENU_POS is not defined", then you've made syntax
	// error in menu_tpl.js file or that file isn't linked properly.
	
	// also take a look at stylesheets loaded in header in order to set styles
	//-->
</script>

<%


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYFEATURELIST
'--------------------------------------------------------------------------------------------------
Sub DisplayFeatureListNav( sLevel )
	Dim sSQL, oFeatures
	
	' Get the top level features that the organization has
	'sSQL = "SELECT F.* FROM egov_organization_features F where F.parentfeatureid = 0 and F.featuretype = 'N'" ' and orgid = " & session("orgid")
	sSQL = "Select F.featureid, F.featurename, F.haspermissions, F.haspermissionlevels, F.adminurl "
	sSQL = sSQL & " from egov_organization_features F, egov_organizations_to_features FO "
	sSQL = sSQL & " where F.parentfeatureid = 0 and F.featuretype = 'N' "
	sSQL = sSQL & " and FO.featureid = F.featureid and FO.orgid = " & Session("orgid")
	sSql = sSql & " Order By F.admindisplayorder, F.featurename"

	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 3, 1

	If Not oFeatures.EOF Then
	
		response.write "var MENU_ITEMS = [['Navigation', null, null," & vbcrlf 
		Do While NOT oFeatures.EOF
			If oFeatures("haspermissions") Then 
				bUserCanNav = UserHasNavRights( oFeatures("Featureid"), Session("userid") )
			Else
				bUserCanNav = True 
			End If 
			If bUserCanNav Then 
				' IF NULL NO LINK TO PAGE
				If IsNull(oFeatures("adminurl")) Then
					sAdminURL = "null"
				Else
					' No first level menu has actual nav capability except home and log off
					If oFeatures("FeatureName") = "Home" Or oFeatures("FeatureName") = "Log Off" Then
						sAdminURL = "'" & sLevel & oFeatures("adminurl") & "'"
					Else 
						'sAdminURL = "'" & sLevel & oFeatures("adminurl") & "'"
						sAdminURL = "null"
					End If 
				End If 
				
				response.write vbcrlf & "['" & oFeatures("FeatureName") & "'," & sAdminURL 
				GetChildrenFeatureListNav oFeatures("Featureid"), sLevel 
			End If 
			oFeatures.MoveNext
		Loop

		' Add logoff
'		response.write vbcrlf & "['Log Off'," & "'" & sLevel & "signoff.asp']"

		response.write vbcrlf & "],];"
		 
	End If
	oFeatures.close
	Set oFeatures = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB GETCHILDRENFEATURELIST(IPARENTID)
'--------------------------------------------------------------------------------------------------
Sub GetChildrenFeatureListNav( iParentID, sLevel )
	Dim bUserCanNav, sSQL, oFeatures
	
	' Get all the features for the parent filtered by what the org has then filter by what the user has permissions to.
	sSQL = "SELECT F.* FROM egov_organization_features F, egov_organizations_to_features FO where F.parentfeatureid = " & iParentID
	sSql = sSql & " and FO.featureid = F.featureid and FO.orgid = " & Session("orgid")
	sSql = sSql & " and F.featuretype = 'N' Order By F.securitydisplayorder, F.featurename"
	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 3, 1

	If NOT oFeatures.EOF Then
		response.write ",null," & vbcrlf
		Do While NOT oFeatures.EOF
			If oFeatures("haspermissions") Then
				bUserCanNav = UserHasNavRights( oFeatures("Featureid"), Session("userid") )
			Else
				bUserCanNav = True 
			End If 
			If bUserCanNav Then 
				' IF NULL NO LINK TO PAGE
				If IsNull(oFeatures("adminurl")) Then
					sAdminURL = "null"
				Else
					' Handle external URLs
					If UCase(Left(oFeatures("adminurl"), 4)) = "HTTP" Then 
						sAdminURL = "'" & oFeatures("adminurl") & "'"
					Else 
						sAdminURL = "'" & sLevel & oFeatures("adminurl") & "'"
					End If 
				End If 
				response.write vbcrlf & vbTab & "['" & oFeatures("FeatureName") & "'," & sAdminURL & "]"
			End If
			oFeatures.MoveNext
			If NOT oFeatures.EOF And bUserCanNav Then
				response.write "," & vbcrlf
			End If 
		Loop
		response.write "]," 
	Else
		response.write "]," & vbcrlf
	End If
	oFeatures.close
	Set oFeatures = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' Function UserHasNavRights( iFeatureid, iUserId )
'--------------------------------------------------------------------------------------------------
Function UserHasNavRights( iFeatureid, iUserId )
	Dim sSql, oFeatures

	UserHasNavRights = False 
	sSql = "Select count(featureid) as hits from egov_users_to_features where featureid = " & iFeatureid & " and userid = " & iUserId

	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 3, 1

	If NOT oFeatures.EOF Then
		oFeatures.MoveFirst
		If clng(oFeatures("hits")) > 0 Then 
			UserHasNavRights = True 
		End If 
	End If 
	oFeatures.close
	Set oFeatures = Nothing

End Function 
%>
