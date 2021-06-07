<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcategorydelete.asp
' AUTHOR: Steve Loar
' CREATED: 09/11/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Deletes Rentals. Called from rentaledit.asp
'
' MODIFICATION HISTORY
' 1.0   09/11/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRecreationCategoryId, sSql

iRecreationCategoryId = CLng(request("rc"))

' Remove the rentals to categories table
sSql = "DELETE FROM egov_rentals_to_categories WHERE recreationcategoryid = " & iRecreationCategoryId
'response.write sSql & "<br /><br />"

RunSQLStatement sSql

' Remove the categories table
sSql = "DELETE FROM egov_recreation_categories WHERE recreationcategoryid = " & iRecreationCategoryId
sSql = sSql & " AND orgid = " & session("orgid")
'response.write sSql & "<br /><br />"

RunSQLStatement sSql

' Back to the rentals list
response.redirect "rentalscategorieslist.asp?s=d"

%>