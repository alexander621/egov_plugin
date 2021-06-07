<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->


<%
' PROCESS DATE VALUES
fromDate = Request("Fromdate")
toDate = Request("Todate")
today = Date()

If orderBy = "" or IsNull(orderBy) Then orderBy = "date" End If
If toDate = "" or IsNull(toDate) Then toDate = today End If
If fromDate = "" or IsNull(fromDate) Then fromDate = cdate(Month(today)& "/1/" & Year(today)) End If

%>

<html>
<head>
  <title><%=langBSPayments%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">

   <script language="Javascript">
  <!--
    function doCalendar(ToFrom) {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }
  //-->
  </script>

	<style>
		/*EXCEL TABLE STYLES*/
		table.excel {}
		td.excelheader {font-weight:bold;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldata {background-color:white;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
		td.excelheaderleft {font-weight:bold;border-left: solid #000000 1px;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldataleft {background-color:#eeeeee;border-left: solid #c0c0c0 1px;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;font-weight:bold;text-align:center;}
		TD {font-family: font-family: arial,tahoma;; font-size: 10px; color: #000000;}
		FONT {font-family: font-family: arial,tahoma;; font-size: 10px; color: #000000;}
		input.excelexport {height:18px;background-repeat:no-repeat;font-family: verdana,sans-serif; font-size: 10px;font-weight:bold;}
	</style>

</head>


<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabPayments,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
	    <td><font size="+1"><b>Class\Event Payment Report</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back()"><%=langBackToStart%></a></td>
    </tr>
	<tr>
    <td>
				 <!--BEGIN: SEARCH OPTIONS-->
				  <fieldset>
				  <legend><b>Search/Sorting Option(s)</b></legend>
				  <form action="class_payment_list.asp" method=post name=frmFilter >
				  <table border=0>
				  <tr>
				  <td valign=top>
					  <b>From: 
					  <input type=text name="fromDate" value="<%=fromDate%>">
					  <a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border=0></a>		 
				  </td>
				  <td>&nbsp;</td>
				   <td valign=top>
					<b>To:</b> 
					  <input type=text name="toDate" value="<%=toDate%>">
					  <a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border=0></a>&nbsp;<input type=submit value="Refresh">
				   </td>
				  </tr>
				  </table>
				  
					</form>
					</fieldset>
					<!--END: SEARCH OPTIONS-->
    </td>
  </tr>
	<tr>
 
      <td colspan="3" valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
      
		<% 
		' BUILD OPTION LIST
		sOptions = "&fromdate=" & fromdate & "&todate=" & todate

		' BUILD SQL STATEMENT
		varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate < '" & toDate & "') "
		varWhereClause = varWhereClause & " AND orgid='" & session("orgid") & "'"
		
		sSQLSUM = "Select sum(amount) as total from egov_class_to_user_payment_report " & varWhereClause
		sSQL = "Select userlname + ', ' + userfname as Payee,lastname + ', ' + firstname as Participant,classname as [Class\Event],paymentdate as [Payment Date],paymenttypename as [Payment Type],paymentlocationname as [Payment Location],isnull(amount,0) as Amount,isnull(refundamount,0) as [Refund Amount],paymentreferenceid as [Transaction ID],(" & sSQLSUM  & ") as [Grand Total] from egov_class_to_user_payment_report " & varWhereClause
			
		
		' DISPLAY RESULTS
		Display_Results sSQL,sOptions
		
		%>
	  
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>
</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' FUNCTION DISPLAY_RESULTS(SSQL,SOPTIONS)
'------------------------------------------------------------------------------------------------------------
Function Display_Results(sSQL,sOptions)

	' INITIALIZE VALUES
	iPageSize = 100
	iCacheSize = 100
	iCursorLocation = 3
	sScriptName = request.servervariables("SCRIPT_NAME")
	session("DISPLAYQUERY") = sSQL
	
	
	' BUILD RECORDSET
	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.PageSize = iPageSize
	oRequests.CacheSize = iCacheSize
	oRequests.CursorLocation = iCursorLocation

	' OPEN RECORDSET
	oRequests.Open sSQL, Application("DSN"), 3, 1
	
	' NAVIGATION VARIABLES
	abspage = oRequests.AbsolutePage
	pagecnt = oRequests.PageCount
	
	' CHECK FOR EMPTY RECORDSET
	If oRequests.EOF then
		' EMPTY
		response.write "<p><b>No records found</p>"
	Else
		' FOUND RECORDS, DISPLAY TO USER

		 ' SET PAGE TO VIEW
		If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
			' DEFAULT FIRST PAGE VIEW
			oRequests.AbsolutePage = 1
		Else
			' DISPLAY CURRENTLY SELECTED PAGE
			If clng(Request("pagenum")) <= clng(oRequests.PageCount) Then
				oRequests.AbsolutePage = request("pagenum")
			Else
				' DEFAULT TO FIRST PAGE VIEW
				oRequests.AbsolutePage = 1
			End If
	End If

	' NAVIGATION BUTTONS
	Response.write "<table style=""margin-bottom: 5px;"">"
	response.write "<tr>"
	response.write "<td valign=bottom><a href=""" & sScriptName & "?pagenum=1" & sOptions & """><img border=0 src=""../images/nav_first.gif""></a><a href=""" & sScriptName & "?pagenum="&oRequests.AbsolutePage-1 & sOptions & """><img border=0 src=""../images/nav_back.gif""></a></td>"
	Response.Write "<TD valign=bottom><B>Number of pages: <font style=""color:blue;""> " & oRequests.PageCount & "</font> | " & vbcrlf
	Response.Write "Current page: <font  style=""color:blue;"">" & oRequests.AbsolutePage & "</font> | " & vbcrlf
	Response.Write "Number of Records: <font  style=""color:blue;"">" & oRequests.RecordCount
	Response.write "</B></TD>"
	response.write "<td valign=bottom><a href=""" & sScriptName & "?pagenum="&oRequests.AbsolutePage+1 & sOptions & """><img border=0 src=""../images/nav_forward.gif"" valign=bottom></a><a href=""" & sScriptName & "?pagenum="&oRequests.PageCount & sOptions & """><img border=0 src=""../images/nav_last.gif"" valign=bottom></a></td>"
	response.write "<td align=right valign=bottom><input type=button class=excelexport value=""Download as CSV"" onClick=""location.href='csv_export.asp'""></td>"
	response.write "</tr>"
	response.write "</table>"
		
				
	' DISPLAY DATA	
	Response.Write "<table cellspacing=0 cellpadding=2 class=excel width=""100%"">"

	' WRITE COLUMN HEADINGS
	response.write "<tr class=excel><td class=excelheaderleft>&nbsp;</td>"
	For Each fldLoop in oRequests.Fields
		response.write "<td class=excelheader>" & fldLoop.Name & "</td>"
	Next
	response.write "</tr >"
	
	' SET BASE RECORD COUNT
	iRecordNumber = (oRequests.AbsolutePage * iPageSize) - iPageSize
		 			
	' LOOP AND DISPLAY THE RECORDS
	For irows = 1 to oRequests.PageSize
		 
		     If NOT oRequests.EOF Then
				bgcolor = "#eeeeee"
				iRecordNumber = iRecordNumber + 1

				Response.Write "<tr>"
				response.write "<td class=exceldataleft>" & iRecordNumber & "</td>"
				
				For Each fldLoop in oRequests.Fields
					iType = fldLoop.type
					response.write "<td class=exceldata>" & FormatData(iType,trim(fldLoop.Value)) & "&nbsp;</td>"
				Next
				
				response.write "</tr>"
				
				oRequests.MoveNext 

			End If

	Next			

	End If
 

End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION FORMATDATA(STYPE,SDATA)
'------------------------------------------------------------------------------------------------------------
Function FormatData(iType,sData)
	
	' DEFAULT RETURN UNALTERED DATA
	sReturnValue = sData

	' FORMAT ACCORDING TO DATA TYPE
	Select Case iType
		Case 6
			' FORMAT DISPLAY AS CURRENCY
			If sData = "" OR IsNull(sData) Then
				' IF NULL OR EMPTY SET TO 0
				sData = 0
			End If
			sReturnValue = FormatCurrency(sData,2)
		Case Else
			' DO NOTHING
		End Select

		FormatData = sReturnValue

End Function
%>
