<%
	Dim sSql, oRequests, oSchema, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal
	Dim iLocationId, toDate, fromDate, sDateRange, iPaymentLocationId, iReportType, sAdminlocation
	Dim sFile, sRptTitle, iAdminUserId

'SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
 if request("reporttype") = "" then
    iReportType = "2"
 else
    iReportType = request("reporttype")
 end if

	sRptTitle = "<tr>" & vbcrlf

 sRptTitle = sRptTitle & "    <th>Purchase Distribution</th>" & vbcrlf

 if iReportType = "2" then
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
 end if

 sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
 sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
 sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
'   displayTHColumns sRptTitle, iReportType
 sRptTitle = sRptTitle & "</tr>" & vbcrlf

	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Purchase_Distribution_" & sDate & ".xls"

'PROCESS REPORT FILTER VALUES
'PROCESS DATE VALUES
	fromDate = Request("fromDate")
	toDate   = Request("toDate")
	today    = Date()

'IF EMPTY DEFAULT TO CURRENT TO DATE
	if toDate = "" or IsNull(toDate) then
    toDate = today 
	end if

	if fromDate = "" or IsNull(fromDate) then
    fromDate = today
	end if

 if request("locationid") = "0" then
  		iLocationId = 0
	else
  		iLocationId = CLng(request("locationid"))
	end if

	if request("adminuserid") = "" then
  		iAdminUserId = 0
	else
  		iAdminUserId = CLng(request("adminuserid"))
	end if

	if request("paymentlocationid") = "" then
  		iPaymentLocationId = 0
	else
  		iPaymentLocationId = CLng(request("paymentlocationid"))
	end if

if request("classseasonid") <> "" then
   iClassSeasonID = request("classseasonid")
else
   iClassSeasonID = ""
end if

'BUILD SQL WHERE CLAUSE
	varWhereClause = " AND (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
 sRptTitle = sRptTitle & "    <th>Payment Date >= "     & fromDate              & "</th>" & vbcrlf
 sRptTitle = sRptTitle & "    <th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th>" & vbcrlf

 displayTHColumns sRptTitle, iReportType

 sRptTitle = sRptTitle & "</tr>" & vbcrlf
	varWhereClause = varWhereClause & " AND P.orgid = " & session("orgid") 

	if iLocationId > 0 then
  		varWhereClause = varWhereClause & " AND adminlocationid = " & iLocationId
		  sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Admin Location: " & GetLocationName( iLocationId ) & "</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

    displayTHColumns sRptTitle, iReportType

    sRptTitle = sRptTitle & "</tr>" & vbcrlf
	else
		  sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Admin Location: All Locations</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

    displayTHColumns sRptTitle, iReportType

    sRptTitle = sRptTitle & "</tr>" & vbcrlf
	end if

	if CLng(iAdminUserId) > CLng(0) then
  		varWhereClause = varWhereClause & " AND adminuserid = " & iAdminUserId
		  sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Admin: " & GetAdminName( iAdminUserId ) & "</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

    displayTHColumns sRptTitle, iReportType

    sRptTitle = sRptTitle & "</tr>" & vbcrlf
	else
		  sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Admin: All Admins</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

    displayTHColumns sRptTitle, iReportType

    sRptTitle = sRptTitle & "</tr>" & vbcrlf
	end if

	if iPaymentLocationId > 0 then
  		if iPaymentLocationId = CLng(2) then
		    	varWhereClause = varWhereClause & " AND P.paymentlocationid = 3 " 
    			sRptTitle = sRptTitle & "<tr>" & vbcrlf
       sRptTitle = sRptTitle & "    <th>Payment Location: Web Site</th>" & vbcrlf
       sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

       displayTHColumns sRptTitle, iReportType

       sRptTitle = sRptTitle & "</tr>" & vbcrlf
  		else
		    	varWhereClause = varWhereClause & " AND P.paymentlocationid < 3 " 
    			sRptTitle = sRptTitle & "<tr>" & vbcrlf
       sRptTitle = sRptTitle & "    <th>Payment Location: Office</th>" & vbcrlf
       sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

       displayTHColumns sRptTitle, iReportType

       sRptTitle = sRptTitle & "</tr>" & vbcrlf
  		end if
	else
		  sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Payment Location: All Locations</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf

    displayTHColumns sRptTitle, iReportType

    sRptTitle = sRptTitle & "</tr>" & vbcrlf
	end if

'Determine which season has been selected
 if iClassSeasonID <> "" then
    varWhereClause = varWhereClause & " AND C.classseasonid = " & iClassSeasonID
 end if 

 if iReportType = "1" then
   	DisplaySummary varWhereClause, sRptTitle
 else
   	DisplayDetails varWhereClause, sRptTitle
 end if

'--------------------------------------------------------------------------------------------------
sub DisplaySummary( varWhereClause, sRptTitle )

	dGrandTotal = CDbl(0.00)

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
 oSchema.fields.append "season", adVarChar, 50, adFldUpdatable
'	oSchema.fields.append "receiptno", adInteger, , adFldUpdatable
'	oSchema.fields.append "activityno", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "amount", adCurrency, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

'	sSQL = "SELECT A.accountname, A.accountnumber, SUM(L.amount) AS sub_total "
'	sSQL = sSQL & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_class_list CL, egov_class_time T "

	sSQL = "SELECT A.accountname, A.accountnumber, SUM(L.amount) AS sub_total, C.ClassSeasonID "
	sSQL = sSQL & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_class_list CL, "
 sSQL = sSQL & " egov_class_time T, egov_class C "
	sSQL = sSQL & " WHERE ispaymentaccount = 0 "
 sSQL = sSQL & " AND L.paymentid = P.paymentid "
 sSQL = sSQL & " AND A.accountid = L.accountid "
 sSQL = sSQL & " AND CL.classlistid = L.itemid "
 sSQL = sSQL & " AND CL.classtimeid = T.timeid "
 sSQL = sSQL & " AND L.entrytype = 'credit' "
 sSQL = sSQL & " AND C.classid = CL.classid "
 sSQL = sSQL & varWhereClause 
 sSQL = sSQL & " GROUP BY A.accountname, A.accountnumber, C.ClassSeasonID "
	sSQL = sSQL & " ORDER BY A.accountname, A.accountnumber, C.ClassSeasonID "
'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	if NOT oRequests.eof then

  	'Loop through and build the display recordset.
		  do while NOT oRequests.eof
  		  	oSchema.addnew 
    			oSchema("accountname")   = oRequests("accountname")
		    	oSchema("accountnumber") = oRequests("accountnumber")
       oSchema("season")        = getSeasonName(oRequests("classseasonid"))
  		  	oSchema("amount")        = CDbl(oRequests("sub_total"))
  		  	dGrandTotal = dGrandTotal + CDbl(oRequests("sub_total"))
    			oSchema.Update
		    	oRequests.MoveNext
  		loop
 else
  	'A blank row
  		oSchema.addnew 
  		oSchema("accountname")   = " "
		  oSchema("accountnumber") = " "
    oSchema("season")        = " "
  		oSchema("amount")        = 0.00
		  oSchema.Update
 end if

	' Sort them 
	'oSchema.Sort = "accountname ASC, accountnumber ASC, receiptno ASC"

'Total Row
	sTotalRow = "<tr>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td>Totals</td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td>" & FormatNumber(dGrandTotal,2) & "</td>" & vbcrlf
 sTotalRow = sTotalRow & "</tr>" & vbcrlf

	oSchema.MoveFirst

	CreateExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

	oRequests.Close
	Set oRequests = Nothing

end sub

'--------------------------------------------------------------------------------------------------
' Sub DisplayDetails( varWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
sub DisplayDetails( varWhereClause, sRptTitle )
	dGrandTotal = CDbl(0.00)

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
 oSchema.fields.append "season", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "receiptno", adInteger, , adFldUpdatable
 oSchema.fields.append "classname", adVarChar, 255, adFldUpdatable
	oSchema.fields.append "activityno", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "amount", adCurrency, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

'	sSQL = "SELECT A.accountname, A.accountnumber, L.paymentid, T.activityno, L.accountid, L.amount "
'	sSQL = sSQL & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_class_list CL, egov_class_time T "
	sSQL = "SELECT A.accountname, A.accountnumber, L.paymentid, C.classname, T.activityno, L.accountid, L.amount, C.ClassSeasonID "
	sSQL = sSQL & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSQL = sSQL & " egov_class_list CL, egov_class_time T, egov_class C "
	sSQL = sSQL & " WHERE ispaymentaccount = 0 "
 sSQL = sSQL & " AND L.paymentid = P.paymentid "
 sSQL = sSQL & " AND A.accountid = L.accountid "
 sSQL = sSQL & " AND CL.classlistid = L.itemid "
 sSQL = sSQL & " AND CL.classtimeid = T.timeid "
 sSQL = sSQL & " AND L.entrytype = 'credit' "
 sSQL = sSQL & " AND C.classid = CL.classid "
 sSQL = sSQL & varWhereClause
	sSQL = sSQL & " ORDER BY A.accountname, A.accountnumber, L.accountid, L.paymentid, T.activityno, C.ClassSeasonID "
'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	if NOT oRequests.eof then

  	'Loop through and build the display recordset.
		  do while NOT oRequests.eof
  		  	oSchema.addnew 
  		  	oSchema("accountid")     = oRequests("accountid")
    			oSchema("accountname")   = oRequests("accountname")
		    	oSchema("accountnumber") = oRequests("accountnumber")
       oSchema("season")        = getSeasonName(oRequests("classseasonid"))
    			oSchema("receiptno")     = oRequests("paymentid")
       oSchema("classname")     = oRequests("classname")
		    	oSchema("activityno")    = oRequests("activityno")
  		  	oSchema("amount")        = CDbl(oRequests("amount"))
  		  	dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
    			oSchema.Update
		    	oRequests.MoveNext
  		loop
 else
  	'A blank row
  		oSchema.addnew 
		  oSchema("accountid")     = 0
  		oSchema("accountname")   = " "
		  oSchema("accountnumber") = " "
    oSchema("season")        = " "
  		oSchema("receiptno")     = 0
    oSchema("classname")     = " "
		  oSchema("activityno")    = " "
  		oSchema("amount")        = 0.00
		  oSchema.Update
 end if

	' Sort them 
	'oSchema.Sort = "accountname ASC, accountnumber ASC, receiptno ASC"

'Total Row
	sTotalRow = "<tr>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td>Totals</td>" & vbcrlf
 sTotalRow = sTotalRow & "    <td>" & FormatNumber(dGrandTotal,2) & "</td>" & vbcrlf
 sTotalRow = sTotalRow & "</tr>" & vbcrlf

	oSchema.MoveFirst

	CreateExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

	oRequests.Close
	Set oRequests = Nothing

End Sub 

'--------------------------------------------------------------------------------------------------
' Function GetAdminName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminName( iUserId )
	Dim sSql, oName

	sSql = "SELECT firstname + ' ' + lastname as username FROM users Where userid = " & iUserId 

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1

	If Not oName.EOF Then
		GetAdminName = oName("username")
	Else
		GetAdminName = ""
	End If 

	oName.close
	Set oName = Nothing 
End Function 

'--------------------------------------------------------------------------------------------------
' Function GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( iLocationid )
	Dim sSql, oLocation

	sSql = "Select name from egov_class_location where locationid = " & iLocationId

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 3, 1
	
	If Not oLocation.EOF Then 
		GetLocationName = oLocation("name")
	Else
		GetLocationName = ""
	End If 

	oLocation.Close 
	Set oLocation = Nothing

End Function

'--------------------------------------------------------------------------------------------------
' Sub CreateExcelDownload( sRtpTitle, sTotalRow )
'--------------------------------------------------------------------------------------------------
Sub CreateExcelDownload( sRtpTitle, sTotalRow )
 'Pulled this in to make sub-totals

	iOldAccountId = CLng(0) 
	dSubTotal     = CDbl(0.00)

 if NOT oSchema.EOF then
		  response.write "<html>" & vbcrlf
    response.write "<body>" & vbcrlf
    response.write "<table border=""1"">" & vbcrlf

 		'Write the title
  		if sRtpTitle <> "" then
		    	response.write sRtpTitle
  		end if

  		response.write "<tr>" & vbcrlf

 		'WRITE COLUMN HEADINGS
  		for each fldLoop in oSchema.Fields
		     	if fldLoop.Name <> "accountid" then
       				response.write  "    <th>" & fldLoop.Name & "</th>" & vbcrlf
     			end if
  		next
		  response.write "</tr>" & vbcrlf

 		'WRITE DATA
  		do while NOT oSchema.eof
       if iReportType = "2" then
   		    	if oSchema("accountid") <> iOldAccountId then
        				'SubTotal row
         				if iOldAccountId <> CLng(0) then
       		   			'Sub Total Row
           					response.write "<tr>" & vbcrlf
                response.write "    <td></td>" & vbcrlf
                response.write "    <td></td>" & vbcrlf
                response.write "    <td></td>" & vbcrlf
                response.write "    <td></td>" & vbcrlf
                response.write "    <td></td>" & vbcrlf
                response.write "    <td>Sub-Total:</td>" & vbcrlf
                response.write "    <td>" & FormatNumber(dSubTotal, 2) & "</td>" & vbcrlf
            				response.write "</tr>" & vbcrlf
         				end if

         				dSubTotal     = CDbl(0.00)
          			iOldAccountId = oSchema("accountid")
          end if
			    end if

  		 	'Normal Row
     		response.write "<tr>" & vbcrlf
    			for each fldLoop in oSchema.Fields
       				sFieldValue = trim(fldLoop.Value)

      				'REMOVE LINE BREAKS
       				if NOT ISNULL(sFieldValue) then
         					sFieldValue = replace(sFieldValue,chr(10),"")
         					sFieldValue = replace(sFieldValue,chr(13),"")
        			end if

       				if fldLoop.Name = "amount" OR fldLoop.Name = "sub_total" then
         					dSubTotal = dSubTotal + CDbl(sFieldValue)
       				end if

        			if fldLoop.Name <> "accountid" then
         					response.write "<td>" & sFieldValue & "</td>" & vbcrlf
        			end if
    			next

    			response.write "</tr>" & vbcrlf

			    oSchema.MoveNext
  		loop
		
 		'Sub Total Row
		  response.write "<tr>" & vbcrlf

    if iReportType = "2" then
       response.write "    <td></td>" & vbcrlf
       response.write "    <td></td>" & vbcrlf
       response.write "    <td></td>" & vbcrlf
       response.write "    <td></td>" & vbcrlf
       response.write "    <td></td>" & vbcrlf
    else
       response.write "    <td></td>" & vbcrlf
       response.write "    <td></td>" & vbcrlf
    end if

    response.write "    <td>Sub-Total:</td>" & vbcrlf
    response.write "    <td>" & FormatNumber(dSubTotal, 2) & "</td>" & vbcrlf
    response.write "</tr>" & vbcrlf

		 'Total Row
		  if sTotalRow <> "" then
    			response.write sTotalRow
		  end if

  		response.write "</table>" & vbcrlf
    response.write "</body>"  & vbcrlf
    response.write "</html>"  & vbcrlf
 else
 		'NO DATA
 end if

end sub

'------------------------------------------------------------
sub displayTHColumns (sRptTitle, p_report_type)

 if iReportType = "2" then
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
 else
    sRptTitle = sRptTitle & "    <td></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
 end if

end sub

'---------------------------------------------------------
function getSeasonName(p_classseasonid)
  lcl_return = ""

  if p_classseasonid <> "" then
     sSQLs = "SELECT seasonname FROM egov_class_seasons WHERE classseasonid = " & p_classseasonid

    	Set oSeasonName = Server.CreateObject("ADODB.Recordset")
    	oSeasonName.Open sSQLs, Application("DSN"), 3, 1

     if not oSeasonName.eof then
        lcl_return = oSeasonName("seasonname")
     end if

  		 oSeasonName.close
   		set oSeasonName = nothing 

  end if

  getSeasonName = lcl_return

end function
%>
<!-- #include file="../includes/adovbs.inc" -->
