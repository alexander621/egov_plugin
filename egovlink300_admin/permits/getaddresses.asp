<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getaddresses.asp
' AUTHOR: Steve Loar
' CREATED: 4/1/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the addresses drop down using a street name search. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   4/1/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sSql, oRs, sResults, sWhere, strFieldName

sWhere = ""

strFieldName = "residentaddressid"
if request.querystring("fieldname") <> "" then strFieldName = request.querystring("fieldname")

If request("searchstreet") <> "" Then 
	sWhere = sWhere & " AND ( residentstreetname LIKE '%" & dbsafe(request("searchstreet")) & "%' "
	sWhere = sWhere & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(request("searchstreet")) & "' "
	sWhere = sWhere & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(request("searchstreet")) & "' "
	sWhere = sWhere & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(request("searchstreet")) & "' )"
End If 

If request("searchnumber") <> "" Then
	sWhere = sWhere & " AND residentstreetnumber = '" & request("searchnumber") & "' "
End If 

If request("searchowner") <> "" Then 
	sWhere = sWhere & " AND listedowner LIKE '%" & dbsafe(request("searchowner")) & "%' "
End If

sSql = "SELECT residentaddressid, residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, ISNULL(listedowner,'') AS listedowner, "
sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, ISNULL(residentunit,'') AS residentunit " 
sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & session("orgid") 
sSql = sSql & sWhere 
sSql = sSql & " ORDER BY sortstreetname, CAST(residentstreetnumber AS INT), residentunit"

'response.write sSql & "<br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRS.EOF Then
	sResults = "<select name='" & strFieldName & "' id='" & strFieldName & "'>"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("residentaddressid") & "'>"
		sResults = sResults & oRs("residentstreetnumber")
		If oRs("residentstreetprefix") <> "" Then 
			sResults = sResults & " " & oRs("residentstreetprefix") 
		End If 
		sResults = sResults & " " & oRs("residentstreetname")
		If oRs("streetsuffix") <> "" Then 
			sResults = sResults & " " & oRs("streetsuffix")
		End If 
		If oRs("streetdirection") <> "" Then 
			sResults = sResults & " " & oRs("streetdirection")
		End If 
		If oRs("residentunit") <> "" Then 
			sResults = sResults & ", " & oRs("residentunit")
		End If 
		If oRs("listedowner") <> "" Then
			sResults = sResults & " (" & Left(oRs("listedowner"),50) & ")"
		End If 
		sResults = sResults & "</option>"
		oRs.MoveNext
	Loop 
	sResults = sResults & "</select>"
Else
	sResults = "<input type='hidden' name='" & strFieldName & "' id='" & strFieldName & "' value='0' />No Match Found"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults



%>
