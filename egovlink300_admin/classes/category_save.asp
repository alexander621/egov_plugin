<!-- #include file="../includes/common.asp" //-->
<%

Call subSaveCategory(request("iCategoryId"), request("sURL"), request("sTitle"), request("sSubtitle"), request("sDescription"), request("bRoot"), request("sAltTag"))


'--------------------------------------------------------------------------------------------------
' SUB subSaveCategory(iCategoryId, sURL, sTitle, sSubtitle, sDescription, bRoot)
' AUTHOR: Terry Foster
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subSaveCategory(iCategoryId, sURL, sTitle, sSubtitle, sDescription, bRoot, sAltTag)
	Dim iSequenceId, sSql, iMsgId

	sURL = DBsafe( sURL )
	sTitle = DBsafe( sTitle )
	sSubtitle = DBsafe( sSubtitle )
	sDescription = DBsafe( sDescription )
	sAltTag = DBsafe( sAltTag )

	if bRoot = "on" then 
		bRoot = 1
	else
		bRoot = 0
	end if

	iSequenceId = GetNextSequenceId( Session("orgid") )

	If iCategoryId = "0" Then
		' Insert new records
		sSql = "INSERT INTO egov_class_categories ( OrgID, imgurl, categorytitle, categorysubtitle, categorydescription, imgalttag, sequenceid ) "
		sSql = sSql & "VALUES ( " & Session("OrgID") & ",'" & sURL & "','" &  sTitle & "','" &  sSubtitle & "','" &  sDescription & "','"
		sSql = sSql & sAltTag & "', " & iSequenceId & " )"
		'response.write sSql & "<br /><br />"
		
		iCategoryId = RunInsertStatement( sSql )
		'response.write "iCategoryId: " & iCategoryId & "<br /><br />"
		
		' Create the subcategory row
		iRootId = GetRootCategory( Session("orgid") )
		sSql = "INSERT INTO egov_class_category_to_subcategory ( categoryid, subcategoryid ) VALUES (" & iRootId & ", " & iCategoryId & ")"
		'response.write sSql & "<br /><br />"
		
		RunSQLStatement sSql
		'CreateSubCategory Session("orgid"), iRootId 
		
		iMsgId = 1
	Else 
		' Update existing records
		sSql = "UPDATE egov_class_categories SET imgurl = '" & sURL & "', categorytitle = '" & sTitle & "', categorysubtitle = '" & sSubtitle & "', "
		sSql = sSql & "categorydescription = '" & sDescription & "', imgalttag = '" & sAltTag  & "' WHERE Categoryid = " & iCategoryId 
		'response.write sSql & "<br /><br />"
		
		RunSQLStatement sSql
		
		iMsgId = 2
	End If

	'response.write "Done.<br /><br />"
	' REDIRECT TO edit page
	response.redirect "category_edit.asp?categoryid=" & iCategoryId & "&msg=" & iMsgId

End Sub


'--------------------------------------------------------------------------------------------------
' Sub CreateSubCategory( iOrgId, iRootId )
'--------------------------------------------------------------------------------------------------
Sub CreateSubCategory( ByVal iOrgId, ByVal iRootId )
	Dim sSql, oRs

	sSQL = "SELECT categoryid FROM egov_class_categories WHERE orgid = " & iOrgId
	sSql = sSql & " AND categoryid <> " & iRootId & " AND categoryid NOT IN (SELECT subcategoryid "
	sSql = sSql & " FROM egov_class_category_to_subcategory WHERE categoryid = " & iRootId & " )"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		'InsertSubCategory iRootId, oRs("categoryid")
		sSql = "INSERT INTO egov_class_category_to_subcategory ( categoryid, subcategoryid ) VALUES ( " & iRootId & ", " & oRs("categoryid") & ")"
		RunSQLStatement sSql
		
		oRs.MoveNext
	Loop 

	oRs.close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub InsertSubCategory( iRootId, iSubCategoryid )
'--------------------------------------------------------------------------------------------------
Sub InsertSubCategory( iRootId, iSubCategoryid )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Insert into egov_class_category_to_subcategory (categoryid, subcategoryid) VALUES (" & iRootId & ", " & iSubCategoryid & ")"
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetRootCategory( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetRootCategory( ByVal iOrgId )
	Dim sSql, oRs

	GetRootCategory = 0

	sSql = "SELECT categoryid FROM egov_class_categories WHERE orgid = " & iOrgId & " AND isroot = 1"
	'response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		GetRootCategory = oRs("categoryid") 
	End If

	' CLEAN UP OBJECTS
	oRs.close
	Set oRs = Nothing
	
End Function


'--------------------------------------------------------------------------------------------------
' Function GetNextSequenceId( iOrgid )
'--------------------------------------------------------------------------------------------------
Function GetNextSequenceId( ByVal iOrgid )
	Dim sSql, oRs

	GetNextSequenceId = 1

	sSql = "SELECT ISNULL(MAX(sequenceid),0) AS sequenceid FROM egov_class_categories WHERE orgid = " & iOrgId 
	'response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetNextSequenceId = CLng(oRs("sequenceid")) + 1
	End If

	oRs.close
	Set oRs = Nothing

End Function 


%>
