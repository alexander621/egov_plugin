<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'' FILENAME: refund_notes_export.asp
'' AUTHOR: SteveLoar
'' AUTHOR: SteveLoar
'' CREATED: 10/22/2013
'' COPYRIGHT: Copyright 2013 eclink, inc.
''			 All Rights Reserved.
''
'' Description:  This pulls together the refund notes report into an Excel spreadsheet. 
'' 				 Part of a Menlo Park Project.
''
'' MODIFICATION HISTORY
'' 1.0   10/22/2013	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
Dim iClassSeasonId, iCategoryid, fromDate, toDate, sWhereClause, sFrom, sSearchName, sClassName
Dim sActivityNo, iPaymentId, toDateDisplay, sDate, sRptTitle, sSql, oRs, iOrderBy, sOrderBy

sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
sWhereClause = ""
sFrom = ""
iOrderBy = 0
sRptTitle = vbcrlf & "<tr><th></th><th>Refund Notes Report</th><th></th><th></th><th></th><th></th><th></th></tr>"

' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "refund notes rpt" ) Then
	response.redirect sLevel & "../permissiondenied.asp"
End If 

If request("classseasonid") = "" Then 
	iClassSeasonId = CLng(0)
Else
	iClassSeasonId = clng(request("classseasonid"))
End If 

If request("categoryid") = "" Or CLng(request("categoryid")) = CLng(0) Then 
	iCategoryid = CLng(0)
Else
	iCategoryid = CLng(request("categoryid"))
End If 

If request("searchname") <> "" Then 
	sSearchName = dbsafe(request("searchname"))
End If 

If request("classname") <> "" Then 
	sClassName = dbsafe(request("classname"))
End If 

If request("activityno") <> "" Then 
	sActivityNo = dbsafe(request("activityno"))
End If

If request("paymentid") = "" Then 
	iPaymentId = ""
Else
	iPaymentId = CLng(request("paymentid"))
End If 

fromDate = Request("fromDate")
toDate = Request("toDate")

If request("orderby") <> "" Then 
	iOrderBy = clng(request("orderby"))
End If 

If clng(iOrderBy) = clng(0) Then
	sOrderBy = "paymentdate"
Else
	sOrderBy = "U.userlname, U.userfname, paymentdate"
End If


' BUILD SQL WHERE CLAUSE
If iClassSeasonId > CLng(0) Then
	sWhereClause = sWhereClause & " AND C.classseasonid = " & iClassSeasonId
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Season: " & GetSeasonName( iClassSeasonId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If iCategoryid > CLng(0) Then
	sFrom = ", egov_class_category_to_class G "
	sWhereClause = sWhereClause & " AND C.classid = G.classid AND G.categoryid = " & iCategoryid
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Category: " & GetCategoryName( iCategoryid )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If sClassName <> "" Then
	sWhereClause = sWhereClause & " AND C.classname LIKE '%" & sClassName & "%' "
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Class Name Like: " & sClassName  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If sActivityNo <> "" Then
	sWhereClause = sWhereClause & " AND T.activityno = '" & sActivityNo & "' "
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Activity No: " & sActivityNo  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If iPaymentId <> "" Then
	sWhereClause = sWhereClause & " AND P.paymentid = " & iPaymentId & " "
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Receipt No: " & iPaymentId  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If sSearchName <> "" Then
	sWhereClause = sWhereClause & " AND ( U.userfname LIKE '%" & sSearchName & "%' OR U.userlname LIKE '%" & sSearchName & "%' )"
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Name Like: " & sSearchName  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If fromDate <> "" Then 
	sWhereClause = sWhereClause & " AND P.paymentdate >= '" & fromDate & "' "
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Drop Date From: " & fromDate  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If toDate <> "" Then 
	toDateDisplay = toDate
	toDate = DateAdd( "d", 1, toDate )
	sWhereClause = sWhereClause & " AND P.paymentdate < '" & toDate & "' "
	sRptTitle = sRptTitle & vbcrlf & "<tr><th>Drop Date To: " & toDateDisplay  & "</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Refund_Notes_" & sDate & ".xls"

sSql = "SELECT P.paymentid, P.paymentdate, P.userid, ISNULL(U.userfname,'' ) AS userfname, ISNULL(U.userlname,'') AS userlname, "
sSql = sSql & "ISNULL(U.useraddress,'') AS useraddress, U.usercity, U.userstate, U.userzip, ISNULL(U.userhomephone,'') AS userhomephone, "
sSql = sSql & "C.classname, C.classseasonid, ISNULL(T.activityno,'') AS activityno, D.dropreason, ISNULL(P.notes,'') AS notes "
sSql = sSql & "FROM egov_class_payment P, egov_users U, egov_class_list L, egov_class C, egov_class_time T, egov_class_dropreasons D " & sFrom
sSql = sSql & "WHERE P.userid = U.userid AND P.paymentid = L.paymentid	AND L.classid = C.classid "
sSql = sSql & "AND L.classtimeid = T.timeid AND P.dropreasonid = D.dropreasonid AND P.journalentrytypeid = 2 AND P.orgid = " & Session("orgid") & sWhereClause
sSql = sSql & " ORDER BY " & sOrderBy

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	response.Write vbcrlf & "<html>"
	response.write vbcrlf & "<body><table border=""1"">"
	response.write sRptTitle
	response.write vbcrlf & "<tr><th>First Name</th><th>Last Name</th><th>Address</th><th>City</th><th>State</th><th>Zip</th><th>Phone</th><th>Class Name</th><th>Class #</th><th>Receipt #</th><th>Date Dropped</th><th>Reason</th><th>Notes</th></tr>"
	response.flush

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td>" & oRs("userfname") & "</td>"
		response.write "<td>" & oRs("userlname") & "</td>"
		response.write "<td>" & oRs("useraddress") & "</td>"
		response.write "<td>" & oRs("usercity") & "</td>"
		response.write "<td>" & oRs("userstate") & "</td>"
		response.write "<td>" & oRs("userzip") & "</td>"
		response.write "<td>" & FormatPhoneNumber(oRs("userhomephone")) & "</td>"
		response.write "<td>" & oRs("classname") & "</td>"
		response.write "<td>" & oRs("activityno") & "</td>"
		response.write "<td>" & oRs("paymentid") & "</td>"
		response.write "<td>" & oRs("paymentdate") & "</td>"
		response.write "<td>" & oRs("dropreason") & "</td>"
		response.write "<td>" & oRs("notes") & "</td>"
		response.write "</tr>"
		response.flush

		oRs.MoveNext
	Loop
	response.write vbcrlf & "</table></body></html>"
	response.flush
End If

oRs.Close
Set oRs = Nothing 




'--------------------------------------------------------------------------------------------------
' string GetCategoryName( iCategoryid )
'--------------------------------------------------------------------------------------------------
Function GetCategoryName( ByVal iCategoryid )
	Dim sSql, oRs

	sSql = "SELECT categorytitle FROM egov_class_categories WHERE categoryid = " & iCategoryid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetCategoryName = oRs("categorytitle") 
	Else
		GetCategoryName = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 



%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
