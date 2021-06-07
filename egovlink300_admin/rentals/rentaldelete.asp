<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaldelete.asp
' AUTHOR: Steve Loar
' CREATED: 08/27/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Deletes Rentals. Called from rentaledit.asp
'
' MODIFICATION HISTORY
' 1.0   08/27/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, sSql, oRs

iRentalId = CLng(request("rentalid"))

' Remove the day rates
sSql = "DELETE FROM egov_rentaldayrates WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Remove the days
sSql = "DELETE FROM egov_rentaldays WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Delete the Categories
sSql = "DELETE FROM egov_rentals_to_categories WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Delete the Images
sSql = "DELETE FROM egov_rentalimages WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Delete the Documents
sSql = "DELETE FROM egov_rentaldocuments WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Delete the Associated Rentals
sSql = "DELETE FROM egov_rentals_to_rentals WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Clear out the Rental Items
sSql = "DELETE FROM egov_rentalitems WHERE rentalid = " & iRentalId
RunSQLStatement sSql

' Clear out the rental fees
sSql = "DELETE FROM egov_rentalfees WHERE rentalid = " & iRentalId
RunSQLStatement sSql

' Delete the rental
sSql = "DELETE FROM egov_rentals WHERE rentalid = " & iRentalId
response.write sSql & "<br /><br />"
RunSQLStatement sSql


' Back to the rentals list
response.redirect "rentalslist.asp?s=d"

%>