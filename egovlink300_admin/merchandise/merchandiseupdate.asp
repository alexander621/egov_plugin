<!-- #include file="../includes/common.asp" //-->
<!-- #include file="merchandisecommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiseupdate.asp
' AUTHOR: Steve Loar
' CREATED: 04/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the merchandise
'
' MODIFICATION HISTORY
' 1.0   04/28/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMerchandiseId, sMerchandise, iShowToPublic, sPrice, iMaxOfferings, x, iInStock, sMerchandiseColor
Dim iDisplayOrder

iMerchandiseId = CLng(request("merchandiseid"))

sMerchandise = "'" & dbsafe(request("merchandise")) & "'"

If request("showpublic") = "on" Then
	iShowToPublic = 1
Else
	iShowToPublic = 0
End If 

sPrice = request("price")

iMaxOfferings = CLng(request("maxofferings"))

If iMerchandiseId = CLng(0) Then 
	sSql = "INSERT INTO egov_merchandise ( orgid, merchandise, price, showpublic ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & sMerchandise & ", " & sPrice & ", " & iShowToPublic & " )"

	iMerchandiseId = RunInsertStatement( sSql ) 
	sSuccessMsg = "Merchandise Item Created"
Else
	sSql = "UPDATE egov_merchandise SET merchandise = " & sMerchandise & ", price = " & sPrice & ", showpublic = " & iShowToPublic
	sSql = sSql & " WHERE merchandiseid = " & iMerchandiseId

	RunSQLStatement sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Remove the old offerings
sSql = "DELETE FROM egov_merchandisecatalog WHERE merchandiseid = " & iMerchandiseId
RunSQLStatement sSql 

For x = 1 To iMaxOfferings	
	'response.write "merchandisecatalogid" & x & " = " & request("merchandisecatalogid" & x) & "<br /><br />"
	' See if the offering exists
	If request("merchandisecatalogid" & x) <> "" Then 
		If request("instock" & x) = "on" Then 
			iInStock = 1
		Else
			iInStock = 0
		End If 
		If request("showpublic" & x) = "on" Then 
			iShowToPublic = 1
		Else
			iShowToPublic = 0
		End If 
		sMerchandiseColor = GetMerchandiseColor( request("merchandisecolorid" & x) )
		iDisplayOrder = GetMerchandiseSizeDisplayOrder( request("merchandisesizeid" & x) )

		sSql = "INSERT INTO egov_merchandisecatalog ( merchandiseid, merchandisecolorid, merchandisesizeid, orgid, instock, "
		sSql = sSql & " showpublic, merchandisecolor, displayorder ) VALUES ( "
		sSql = sSql & iMerchandiseId & ", " & request("merchandisecolorid" & x) & ", " & request("merchandisesizeid" & x) & ", "
		sSql = sSql & session("orgid") & ", " & iInStock & ", " & iShowToPublic & ", '" & dbsafe(sMerchandiseColor) & "', "
		sSql = sSql & iDisplayOrder & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	End If 
Next 

' Take them to the edit page for this merchandise
response.redirect "merchandiseedit.asp?merchandiseid=" & iMerchandiseId & "&success=" & sSuccessMsg

%>