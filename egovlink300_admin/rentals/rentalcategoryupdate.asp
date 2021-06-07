<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcategoryupdate.asp
' AUTHOR: Steve Loar
' CREATED: 09/11/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates Rental Categories. Called from rentalcategoryedit.asp
'
' MODIFICATION HISTORY
' 1.0   09/11/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRecreationCategoryId, sSql, sCategoryTitle, sMessageFlag, sImgUrl, sHasRestrictedPeriod
Dim iRestrictedPeriodId, sCategoryDescription, sHideFromPublic

iRecreationCategoryId = CLng(request("recreationcategoryid"))

sCategoryTitle = "'" & dbsafe(request("categorytitle")) & "'"

If request("imgurl") <> "" Then 
	sImgUrl = "'" & dbsafe(request("imgurl")) & "'" 
Else
	sImgUrl = "NULL"
End If 

If request("hasrestrictedperiod") = "on" Then 
	sHasRestrictedPeriod = "1"
Else
	sHasRestrictedPeriod = "0 "
End If 

If request("hidefrompublic") = "on" Then 
	sHideFromPublic = "1"
Else
	sHideFromPublic = "0 "
End If 

iRestrictedPeriodId = request("restrictedperiodid")

If request("categorydescription") = "" Then 
	sCategoryDescription = "NULL"
Else
	sCategoryDescription = "'" & DBsafeWithHTML(request("categorydescription")) & "'"
End If 

If CLng(iRecreationCategoryId) > CLng(0) Then
	sMessageFlag = "u"

	' Update existing category
	sSql = "UPDATE egov_recreation_categories SET categorytitle = " & sCategoryTitle
	sSql = sSql & ", imgurl = " & sImgUrl
	sSql = sSql & ", categorydescription = " & sCategoryDescription
	sSql = sSql & ", hasrestrictedperiod = " & sHasRestrictedPeriod
	sSql = sSql & ", restrictedperiodid = " & iRestrictedPeriodId
	sSql = sSql & ", hidefrompublic = " & sHideFromPublic
	sSql = sSql & " WHERE recreationcategoryid = " & iRecreationCategoryId 
	sSql = sSql & " AND orgid = " & session("orgid")
	response.write sSql & "<br />"

	RunSQLStatement sSql

Else
	sMessageFlag = "n"

	' Create a new category 
	sSql = "INSERT INTO egov_recreation_categories ( categorytitle, isroot, orgid, imgurl, hasrestrictedperiod, "
	sSql = sSql & " restrictedperiodid, categorydescription, hidefrompublic, isforrentals ) VALUES ( "
	sSql = sSql & sCategoryTitle & ", 0, " & session("orgid") & ", " & sImgUrl & ", " & sHasRestrictedPeriod & ", "
	sSql = sSql & iRestrictedPeriodId & ", " & sCategoryDescription & ", " & sHideFromPublic & ", 1 )"
	'response.write sSql & "<br />"

	iRecreationCategoryId = RunInsertStatement( sSql )
End If 

' Take them to the edit page for this rental category
response.redirect "rentalcategoryedit.asp?rc=" & iRecreationCategoryId & "&s=" & sMessageFlag

%>