<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: catalogexport.asp
' AUTHOR: SteveLoar
' CREATED: 03/05/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   03/05/2010	Steve Loar - INITIAL VERSION
' 1.1	10/13/2011	Steve Loar - Added prices for Maricopa
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iStatusid, iClassTypeId, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange
Dim sSearchName, sSearchActivity, sWhere, sFrom, bShowPricesInCatalog

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Class_Catalog.xls"

sWhere = ""
sFrom = ""

bShowPricesInCatalog = orgHasFeature( "show prices in catalog" )


' Class Season Id
If request("classseasonid") = "" Then 
   iClassSeasonId = GetRosterSeasonId()
Else
   iClassSeasonId = CLng(request("classseasonid"))
   bFilter = True
End If 

If CLng(iClassSeasonId) > CLng(0) Then 
	sWhere = sWhere & " AND C.classseasonid = " & iClassSeasonId
End If 


' Class Status
If request("statusid") = "" Or CLng(request("statusid")) = 0 Then 
   iStatusid = CLng(0)
Else
   iStatusid = request("statusid")
   bFilter = True
End If 

If CLng(iStatusid) > CLng(0) Then 
	sWhere = sWhere & " AND C.statusid = " & iStatusid
End If 


' Class Type 
If request("classtypeid") = "" Or CLng(request("classtypeid")) = CLng(0) Then 
	iClasstypeid = CLng(0)
Else
	iClasstypeid = CLng(request("classtypeid"))
	bFilter = True
End If 

If CLng(iClassTypeId) > CLng(0) Then
	sWhere = sWhere & " AND C.classtypeid = " & iClassTypeId
End If 


If request("categoryid") = "" Or CLng(request("categoryid")) = CLng(0) Then 
	iCategoryid = CLng(0)
Else
	iCategoryid = CLng(request("categoryid"))
	bFilter = True
End If 

If CLng(iCategoryid) > CLng(0) Then
	sWhere = sWhere & " AND CL.categoryid = " & iCategoryid
End If 


' Date Filters and selected start and end dates
If request("datefilter") = "" Then 
	iDatefilter = ""
Else
	iDatefilter = request("datefilter")
	bFilter = True 
End If 

If request("startdate") = "" Then 
	sStartdate = ""
Else
	sStartdate = request("startdate")
	bFilter = True
End If
If request("enddate") = "" Then 
	sEnddate = ""
Else
	sEnddate = request("enddate")
	bFilter = True
End If
If iDatefilter = "alldates" Then
	sStartdate = ""
	sEnddate = ""
'	bFilter = False 
End If

' if all date choices are blank, give them the current published classes and events
If sShowDates = 1 And bFilter = False Then
	sDefaultRange = " AND (('' + CONVERT(CHAR(8),GETDATE(),112) + '' >= publishstartdate AND '' + CONVERT(CHAR(8),GETDATE(),112) + '' <= publishenddate) OR publishstartdate > '' + CONVERT(CHAR(8),GETDATE(),112) + '' OR publishstartdate IS NULL) "
Else 
	sDefaultRange = ""
End If 

If bFilter Then
	sDefaultRange = ""
End If 

If iDatefilter <> "alldates" And (sStartdate <> "" Or sEnddate <> "") Then
	If sStartdate <> "" Then 
		sWhere = sWhere & " AND C." & iDatefilter & " >= '" & sStartdate & "' " 
	End If
	If sEnddate <> "" Then 
		sWhere = sWhere & " AND C." & iDatefilter & " <= '" & sEnddate & "' "
	End If 
	sWhere = sWhere & " AND C." & iDatefilter & " IS NOT NULL "
Else 
	' add in the default range of classes and events to get
	sWhere = sWhere & sDefaultRange
End If 


' Class Name Like
sSearchName = request("searchname")
If sSearchName <> "" Then 
	sSearchName = dbsafe(sSearchName)
	sWhere = sWhere & " AND LOWER(C.classname) LIKE LOWER('%" & sSearchName & "%') "
End If 


' Activity Number Like
sSearchActivity = request("searchactivity")
If sSearchActivity <> "" Then 
	sSearchActivity = dbsafe(sSearchActivity)
	sWhere = sWhere & " AND LOWER(T.activityno) = LOWER('" & sSearchActivity & "') "
End If 


sSql = "SELECT CA.categorytitle, C.classid, C.classname, C.classdescription, L.name AS locationname, "
sSql = sSql & " ISNULL(C.minage,0.0) AS minage, ISNULL(C.maxage,99.9) AS maxage, "
sSql = sSql & " CONVERT(VARCHAR(2),MONTH(C.startdate)) + '/' + CONVERT(VARCHAR(2),DAY(C.startdate)) AS startdate, "
sSql = sSql & " CONVERT(VARCHAR(2),MONTH(C.enddate)) + '/' + CONVERT(VARCHAR(2),DAY(C.enddate)) AS enddate, "
sSql = sSql & " ISNULL(T.max,999) AS maxsize, ISNULL(T.min,999) AS minsize, D.starttime, D.endtime, "
sSql = sSql & " CASE WHEN sunday = 1 THEN 'Su' ELSE '' END + CASE WHEN monday = 1 THEN 'M' ELSE '' END + "
sSql = sSql & " CASE WHEN tuesday = 1 THEN 'T' ELSE '' END + CASE WHEN wednesday = 1 THEN 'W' ELSE '' END + "
sSql = sSql & " CASE WHEN thursday = 1 THEN 'Th' ELSE '' END + CASE WHEN friday = 1 THEN 'F' ELSE '' END + "
sSql = sSql & " CASE WHEN saturday = 1 THEN 'S' ELSE '' END AS weekdays "
sSql = sSql & " FROM egov_class C, egov_class_location L, egov_class_category_to_class CL, "
sSql = sSql & " egov_class_categories CA, egov_class_time T, egov_class_time_days D " 
sSql = sSql & " WHERE C.orgid = " & session("orgid") & " AND C.locationid = L.locationid AND CA.categoryid = CL.categoryid "
sSql = sSql & " AND CL.classid = C.classid AND T.classid = C.classid AND T.timeid = D.timeid " & sWhere
sSql = sSql & " ORDER BY CA.categorytitle, C.classname, L.name"
'response.write sSql & "<br /><br />"
'response.End 

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

response.write "<html>"

response.write vbcrlf & "<style>  "
response.write " .moneystyle "
response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
response.write vbcrlf & "</style>"

response.write "<body><table border=""1"">"
response.flush

If Not oRs.EOF Then
	response.write "<tr><th>Category</th><th>Class Name</th><th>Age Suit.</th><th>Description</th>"
	response.write "<th>Min Age</th><th>Max Age</th><th>Session Dates</th><th>Weekdays</th>"
	response.write "<th>Time Range</th><th>Location</th><th>Min Size</th><th>Max Size</th>"
	If bShowPricesInCatalog Then
		response.write "<th>Price</th>"
	End If 
	response.write "<th>New ?</th></tr>"

	Do While Not oRs.EOF
		response.write "<tr>"
		response.write "<td>" & oRs("categorytitle") & "</td>"
		response.write "<td>" & oRs("classname") & "</td>"
		response.write "<td>&nbsp;"
		If clng(oRs("minage")) > clng(0) Then 
			response.write clng(oRs("minage"))
			If CDbl(oRs("maxage")) < CDbl(99.9) Then 
				If clng(oRs("minage")) <> clng(oRs("maxage")) Then 
					response.write "-" & clng(oRs("maxage"))
				Else
					response.write "&nbsp;"
				End If 
			Else
				response.write "+"
			End If
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td>" & oRs("classdescription") & "</td>"
		response.write "<td>"
		If clng(oRs("minage")) > clng(0) Then 
			response.write clng(oRs("minage")) 
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td>" 
		If CDbl(oRs("maxage")) < CDbl(99.9) Then
			response.write clng(oRs("maxage"))
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td>&nbsp;" 
		response.write oRs("startdate")
		If oRs("startdate") <> oRs("enddate") Then 
			response.write "-" & oRs("enddate") 
		End If 
		response.write "</td>"
		response.write "<td>" 
		If oRs("weekdays") <> "" Then
			response.write oRs("weekdays")
		Else 
			response.write "&nbsp;"
		End If  
		response.write "</td>"

		response.write "<td>&nbsp;"
		If oRs("starttime") <> "" Then 
			If Right(oRs("starttime"),2) = Right(oRs("endtime"),2) Then
				response.write Left( oRs("starttime"), Len(oRs("starttime")) - 2 )
			Else
				response.write oRs("starttime")
			End If 
			If oRs("endtime") <> "" Then 
				response.write "-" & oRs("endtime")
			End If 
		Else
			response.write "&nbsp;"
		End If
		response.write "</td>" 

		response.write "<td>" & oRs("locationname") & "</td>"

		response.write "<td>" 
		If clng(oRs("minsize")) < clng(999) And clng(oRs("minsize")) > clng(0) Then 
			response.write clng(oRs("minsize"))
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		' Max Size
		response.write "<td>" 
		If clng(oRs("maxsize")) < clng(999) Then 
			response.write clng(oRs("maxsize"))
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		' The class price
		If bShowPricesInCatalog Then
			' This gets one price only
			response.write "<td class=""moneystyle"">"
			response.write GetClassPrice( oRs("classid") )
			response.write "</td>"
		End If 

		' The new column
		response.write "<td>&nbsp;</td>"

		response.write "</tr>"
		response.flush
		oRs.MoveNext 
	Loop 

End If 

response.write "</table></body></html>"

oRs.Close
Set oRs = Nothing 


'------------------------------------------------------------------------------
' string GetClassPrice( iClassId )
'------------------------------------------------------------------------------
Function GetClassPrice( ByVal iClassId )
	Dim sSql, oRs

	' this gets the first price found for a class. so there should be only one for this to work
	sSql = "SELECT ISNULL(amount,0.00) AS amount FROM egov_class_pricetype_price WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetClassPrice = FormatNumber(oRs("amount"),2)
	Else
		GetClassPrice = FormatNumber(0.00,2)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
