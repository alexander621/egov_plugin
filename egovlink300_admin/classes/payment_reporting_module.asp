<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->


<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payment_reporting_module.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/10/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   1/10/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then toDate = today End If
If fromDate = "" or IsNull(fromDate) Then fromDate = cdate(Month(today)& "/1/" & Year(today)) End If

' BUILD SQL WHERE CLAUSE
varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
varWhereClause = varWhereClause & " AND orgid='" & session("orgid") & "'"
%>



<html>
<head>
  <title>E-Gov Advanced Payment Reporting</title>

 
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="Javascript" src="tablesort.js"></script>

	<script language="Javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}
	  //-->
	</script>

	<script language="Javascript" src="dates.js"></script>

	<style>
		/*EXCEL TABLE STYLES*/
		table.excel {}
		th.excelheader {font-weight:bold;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;cursor: pointer;}
		td.exceltotalrow {text-align:right;font-weight:bold;background-color:eeeeee;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
		td.excelgrandtotalrow {text-align:right;font-weight:bold;background-color:dddddd;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
		td.exceldata {text-align:right;background-color:white;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
		th.excelheaderleft {cursor: pointer;font-weight:bold;border-left: solid #000000 1px;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldataleft {background-color:#eeeeee;border-left: solid #c0c0c0 1px;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;font-weight:bold;text-align:center;}
		input.excelexport {height:18px;background-repeat:no-repeat;font-family: verdana,sans-serif; font-size: 10px;font-weight:bold;}
		TD {font-family: font-family: verdana,sans-serif; font-size: 12px; color: #000000;}
		FONT {font-family: font-family: verdana,sans-serif; font-size: 12px; color: #000000;}
	</style>

</head>


<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">


<% ShowHeader sLevel %>


<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->


<form action="payment_reporting_module.asp" method=post name=frmPFilter >

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>E-Gov Advanced Payment Reporting</b></font></td>
		</tr>
		<tr>
			<td>
				<fieldset >
					
					<legend ><b>Date Filters:</b></legend>
				
				<!--BEGIN: FILTERS-->
				<!--BEGIN: DATE FILTERS-->
				<P>
				<table>
					<tr>
						<td  align=right> <b>Payment Date: </td>
						<td>
							<input type=text name="fromDate" value="<%=fromDate%>">
							<a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border=0></a>		 
						</td>
						<td>&nbsp;</td>
						<td >
							<b>To:</b> 
						</td>
						<td>
							<input type=text name="toDate" value="<%=toDate%>">
							<a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border=0></a>
						</td>
						<td>&nbsp;</td>
						<td><%DrawDateChoices("Dates")%></td>
					</tr>
				</table>
				</p>
				<!--END: DATE FILTERS-->


				<!--BEGIN: OTHER FILTERS-->
				<!--<P>
				
				<%
					' BASE VIEW DATA
					' sSQLBase = "SELECT paymenttypename,category,item,paymentlocationname,account FROM egov_glreport_combined WHERE orgid='" & session("orgid") & "'"
					' GetFieldChoices(sSQLBase)
				%>
				</p>-->
				<!--END: OTHER FILTERS-->
	

				</fieldset>
				<!--END: FILTERS-->



				 <!--BEGIN: PREDEFINED REPORTS-->
				  <fieldset>

					<legend><b>Predefined Reports:</b></legend>

					<P>
					  <Select name="ireport">
						<option value=1 <%if request("ireport") = 1 Then response.write " SELECTED " End If %> > Daily Receipt - List
						<option Value=3 <%if request("ireport") = 3 Then response.write " SELECTED " End If %>> Daily Receipt - Detail
						<option value=2 <%if request("ireport") = 2 Then response.write " SELECTED " End If %>> Daily Receipt - Summary
						<option value=4 <%if request("ireport") = 4 Then response.write " SELECTED " End If %>> Monthly Revenue by Category - List
						<option value=6 <%if request("ireport") = 6 Then response.write " SELECTED " End If %>> Monthly Revenue by Category - Detail
						<option value=5 <%if request("ireport") = 5 Then response.write " SELECTED " End If %>> Monthly Revenue by Category - Summary
						<option value=7 <%if request("ireport") = 7 Then response.write " SELECTED " End If %>> Monthly Revenue by Source - List
						<option value=9 <%if request("ireport") = 9 Then response.write " SELECTED " End If %>> Monthly Revenue by Source - Detail
						<option value=8 <%if request("ireport") = 8 Then response.write " SELECTED " End If %>> Monthly Revenue by Source - Summary
					  </select>
					  <input class=excelexport type=submit value="View Report">
					</P>
				 

					</fieldset>
				 <!--END: PREDEFINED REPORTS-->

				
    </td>
  </tr>
	<tr>
 
      <td colspan="3" valign="top">
	  
	  
		<!--BEGIN: DISPLAY RESULTS-->
		<!-- #include file="payment_queries.asp" //-->

		<%
		
		' DISPLAY RESULTS
		Display_Results sSQL,sOptions
		
		%>
		<!-- END: DISPLAY RESULTS -->
      

	  </td>
       
    </tr>
  </table>


  </form>
  

<!--END: PAGE CONTENT-->



<!--#Include file="../admin_footer.asp"-->  


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
	iPageSize = 1000
	iCacheSize = 1000
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
	Response.Write "<TD valign=bottom ><B>Number of pages: <font style=""color:blue;"" > " & oRequests.PageCount & "</font> | " & vbcrlf
	Response.Write "Current page: <font  style=""color:blue;"">" & oRequests.AbsolutePage & "</font> | " & vbcrlf
	Response.Write "Number of Rows: <font  style=""color:blue;"">" & oRequests.RecordCount
	Response.write "</B></TD>"
	response.write "<td valign=bottom><a href=""" & sScriptName & "?pagenum="&oRequests.AbsolutePage+1 & sOptions & """><img border=0 src=""../images/nav_forward.gif"" valign=bottom></a><a href=""" & sScriptName & "?pagenum="&oRequests.PageCount & sOptions & """><img border=0 src=""../images/nav_last.gif"" valign=bottom></a></td>"
	response.write "<td align=right valign=bottom><input type=button class=excelexport value=""Download as CSV"" onClick=""location.href='csv_export.asp'""></td>"
	response.write "</tr>"
	response.write "</table>"
		
				
	' DISPLAY DATA	
	Response.Write "<table cellspacing=0 cellpadding=2  class=""excel style-alternate sortable-onload-1"" width=""100%"">"


	' WRITE COLUMN HEADINGS
	response.write "<tr class=excel><th class=""excelheaderleft sortable"">&nbsp;</th>"
	For Each fldLoop in oRequests.Fields 
		If LEFT(fldLoop.Name,3) <> "ecG" Then
			response.write "<th class=""excelheader sortable"">" & fldLoop.Name & "</th>"
		End If
	Next
	response.write "</tr >"
	

	' SET BASE RECORD COUNT
	iRecordNumber = (oRequests.AbsolutePage * iPageSize) - iPageSize
		 			

	' LOOP AND DISPLAY THE RECORDS
	For irows = 1 to oRequests.PageSize


		     If NOT oRequests.EOF Then
				bgcolor = "#eeeeee"
				iRecordNumber = iRecordNumber + 1
				blnSkipCheck = False
				
				response.write "<TR>"

				response.write "<td class=exceldataleft>" & iRecordNumber & "</td>"
				
				For Each fldLoop in oRequests.Fields 

				If LEFT(fldLoop.Name,3) <> "ecG" AND blnSkipCheck <> True Then
					sClass = "exceldata"
				Else
					
					
					If LEFT(fldLoop.Name,3) = "ecG" AND LEFT(fldLoop.Name,4) <> "ecGT"  Then
						If Trim(fldLoop.Value) > 0 Then
							If sClass <> "excelgrandtotalrow" Then
								sClass = "exceltotalrow"
								blnSkipCheck = True
							End If
						End If
					End If


					If LEFT(fldLoop.Name,4) = "ecGT" Then
						If Trim(fldLoop.Value) > 0 Then
							sClass = "excelgrandtotalrow"
							blnSkipCheck = True
						End If
					End If
		
				End If 

					If LEFT(fldLoop.Name,3) <> "ecG" Then
						iType = fldLoop.type
						response.write "<td class=" & sClass &">" & FormatData(iType,trim(fldLoop.Value)) & "&nbsp;</td>"
					End If
				Next
				
				response.write "</tr>"
				
				blnSkipCheck = False

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
		
		Case 129
			' FORMAT DATE AS MONTH SHORT NAME AND YEAR
			If NOT isnull(sData) Then
				sReturnValue = Left(MonthName(RIGHT(sData,2)),3) & " " & LEFT(sData,4)
			End If

		Case Else
			' DO NOTHING
		End Select

		FormatData = sReturnValue 

End Function


'------------------------------------------------------------------------------------------------------------
' SUB SUBDRAWGROUPBYOPTIONS(SSQL)
'------------------------------------------------------------------------------------------------------------
Sub subDrawGroupByOptions(sSQL)

	Set oSelectList = Server.CreateObject("ADODB.Recordset")
	oSelectList.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oSelectList.EOF Then
	
		' WRITE COLUMN HEADINGS
		response.write "<select name=""groupbyfield"" multiple>"
		response.write "<option value=""-1"">Do Not Summarize</option>"
		
		i = 0

		For Each fldLoop in oSelectList.Fields
			response.write "<option value=""" & fldLoop.Name & """>" & UCASE(fldLoop.Name) & "</option>"
			i = i + 1
		Next

		response.write "</select>"


	End If

	Set oSelectList = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION DRAWDATECHOICES(SNAME)
'------------------------------------------------------------------------------------------------------------
Function DrawDateChoices(sName)

	response.write "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=calendarinput Name=" & sName & ">"
	response.write "<option value=0>Or Select Common Date Range Below...</option>"
	response.write "<option value=11>This Week</option>"
	response.write "<option value=12>Last Week</option>"
	response.write "<option value=1>This Month</option>"
	response.write "<option value=2>Last Month</option>"
	response.write "<option value=3>This Quarter</option>"
	response.write "<option value=4>Last Quarter</option>"
	response.write "<option value=6>Year to Date</option>"
	response.write "<option value=5>Last Year</option>"
	response.write "<option value=7>All Dates to Date</option>"
	response.write "</select>"

End Function


'------------------------------------------------------------------------------------------------------------
' SUB GETFIELDCHOICES(SSQL)
'------------------------------------------------------------------------------------------------------------
Sub GetFieldChoices(sSQL)

	sSQLSub = Right(sSQL,Len(sSQL) - instr(sSQL," FROM "))

	Set oChoices = Server.CreateObject("ADODB.Recordset")
	oChoices.Open sSQL, Application("DSN"), 3, 1
	

	If NOT oChoices.EOF Then

			response.write "<table>"
			For Each oColumn in oChoices.Fields 
				response.write "<tr>"
				response.write "<TD align=right><b>" &  UCASE(oColumn.Name)  & ":</b></td>"
				response.write "<TD>"
				response.write "<select name=""filter_" &  UCASE(oColumn.Name) & """>"
				response.write "<option value = ""NO FILTER"">Do Not Filter</option>"
				Call ListFieldChoices("SELECT DISTINCT [" & oColumn.Name & "] " & sSQLSub)
				response.write "</select>"
				response.write "</td>"
				response.write "</tr>"
			Next
			response.write "</table>"
	End If


	Set oChoices = Nothing


End Sub


'------------------------------------------------------------------------------------------------------------
' SUB LISTFIELDCHOICES(SSQL)
'------------------------------------------------------------------------------------------------------------
Sub ListFieldChoices(sSQL)

	Set oChoices = Server.CreateObject("ADODB.Recordset")
	oChoices.Open sSQL, Application("DSN"), 3, 1
	

	If NOT oChoices.EOF Then
		
		Do While NOT oChoices.EOF
			response.write "<option value=""" & oChoices.Fields(0).Value & """>" & oChoices.Fields(0).Value & "</option>"
			oChoices.MoveNext
		Loop

	End If


	Set oChoices = Nothing

End Sub
%>
