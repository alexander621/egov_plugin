<!-- #include file="../includes/common.asp" //-->
<%


subDeleteCategory request("iCategoryId")



'--------------------------------------------------------------------------------------------------
' SUB subDeleteCategory(InstructorId)
' AUTHOR: TERRY FOSTER
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteCategory( ByVal iCategoryId )
	Dim sSql
	
	sSql = "DELETE FROM egov_class_Categories WHERE Categoryid = " & iCategoryId 
	'response.write sSql
	'response.end
	
	RunSQLStatement sSql
	
	sSql = "DELETE FROM egov_class_category_to_subcategory WHERE subcategoryid = " & iCategoryId 
	
	RunSQLStatement sSql

	' Renumber the sequenceid of the remaining categories
	CategoryReorder Session("orgid") 

	' REDIRECT TO facility waivers page
	response.redirect "Category_mgmt.asp?msg=3"

End Sub


'--------------------------------------------------------------------------------------------------
' Sub CategoryReorder( iOrgId )
'--------------------------------------------------------------------------------------------------
Sub CategoryReorder( ByVal iOrgId )
	Dim iNewOrder, oRs, sSql

	iNewOrder = 0
	
	sSql = "SELECT categoryid FROM egov_class_Categories WHERE orgid = " & iOrgId & "ORDER BY sequenceid"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iNewOrder = iNewOrder + 1
		
		sSql = "UPDATE egov_class_Categories SET sequenceid = " & iNewOrder & " WHERE categoryid = " & oRs("categoryid")
		RunSQLStatement sSql
		
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
