<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
Dim fromDate, toDate, today, orderBy

sLevel = "../" ' Override of value from common.asp

'PROCESS DATE VALUES
fromDate = Request("Fromdate")
toDate   = Request("Todate")
today    = Date()

If orderBy  = "" Or IsNull(orderBy) Then
	orderBy  = "date" 
End If 

If toDate   = "" Or IsNull(toDate) Then
	toDate   = today  
End If 

If fromDate = "" Or IsNull(fromDate) Then
	fromDate = CDate(Month(today)& "/1/" & Year(today)) 
End If 


%>
<html>
<head>
	<title>E-GovLink Product Usage Stats</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

	<script language="javascript" src="scripts/modules.js"></script>

	<style>
		/*EXCEL TABLE STYLES*/
		table.excel        { }
		td.excelheader     { font-weight:bold;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldata       { background-color:white;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
		td.excelheaderleft { font-weight:bold;border-left: solid #000000 1px;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldataleft   { background-color:#eeeeee;border-left: solid #c0c0c0 1px;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;font-weight:bold;text-align:center;}
		TD                 { font-family: font-family: arial,tahoma;; font-size: 10px; color: #000000;}
		FONT               { font-family: font-family: arial,tahoma;; font-size: 10px; color: #000000;}
		input.excelexport  { height:18px;background-repeat:no-repeat;font-family: verdana,sans-serif; font-size: 10px;font-weight:bold;}
	</style>

	<script language="javascript" > 
	<!--

		//Set timezone in cookie to retrieve later
		var d=new Date();
		if (d.getTimezoneOffset)
		{
			var iMinutes = d.getTimezoneOffset();
			document.cookie = "tz=" + iMinutes;
		}

		function doPicker(sFormField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function openWin2(url, name) 
		{
			popupWin = window.open(url, name, "resizable,width=500,height=450");
		}

		function openCustomReports(p_report) 
		{
			w = 900;
			h = 500;
			t = (screen.availHeight/2)-(h/2);
			l = (screen.availWidth/2)-(w/2);
			eval('window.open("../customreports/customreports.asp?cr='+p_report+'", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
		}

	//-->
	</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<%
	'BEGIN: Page Content
	response.write vbcrlf & "<div id=""content"">"
	response.write vbcrlf & "<div id=""centercontent"">"
	response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" class=""start"" width=""100%"">"
	response.write vbcrlf & "<tr>"
	response.write "<td>"
	response.write "<font size=""+1""><strong>E-GovLink Product Usage Information as of "
	response.write "(<font style=""color:#0000ff;"">" & Now() & "</font>)</font></strong>"
	response.write "</td>"
	response.write "</tr>"

	'Custom Report - "Total Public Requests Per Org"
	response.write vbcrlf & "<tr>"
	response.write "<td align=""right"">"
	response.write "<input type=""button"" class=""button"" name=""sCustomReports_PublicRequestsPerOrg"" id=""sCustomReports_PublicRequestsPerOrg"" class=""excelexport"" value=""Custom Report - Total Public Requests per Org"" onClick=""openCustomReports('PUBLICREQUESTSBYORG')"" />"
	response.write "</td>"
	response.write "</tr>"

	'BEGIN: Search Options
	response.write vbcrlf & "<tr>"
	response.write "<td>"
	response.write vbcrlf & "<fieldset>"
	response.write vbcrlf & "<legend><strong>Information&nbsp;</strong></legend>"
	response.write vbcrlf & "All of the information displayed below is unfiltered.  "
	response.write "It reflects summary data for all transactions processed thru the E-GovLink system."
	response.write "</fieldset>"
	response.write "</td>"
	response.write "</tr>"
	'END: Search Options

	response.write vbcrlf & "<tr>"
	response.write "<td colspan=""3"" valign=""top"">"

	'Build Option List
	sOptions = "&fromdate=" & fromdate & "&todate=" & todate

	'Build SQL
	sSql = "SELECT * FROM egov_organization_stat_view"		

	' Display Results
	Display_Results sSql, sOptions

	response.write "</td>"
	response.write "</tr>"
	response.write vbcrlf & "</table>"
	response.write vbcrlf & "</div>"
	response.write vbcrlf & "</div>"
	'END: Page Content
%>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%

'-------------------------------------------------------------------------------------------------
' void Display_Results sSql, sOptions 
'-------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sSql, ByVal sOptions )
	Dim oRs

	'Initialize Values
	iPageSize       = 100
	iCacheSize      = 100
	iCursorLocation = 3
	sScriptName     = request.servervariables("SCRIPT_NAME")
	session("DISPLAYQUERY") = sSql

	'Build Recordset
	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.PageSize       = iPageSize
	oRs.CacheSize      = iCacheSize
	oRs.CursorLocation = iCursorLocation
	oRs.Open sSql, Application("DSN"), 3, 1

	'Navigation Variables
	abspage = oRs.AbsolutePage
	pagecnt = oRs.PageCount
	
	If oRs.EOF Then 
		response.write "<p style=""color:#ff0000;""><strong> -- No records found -- </strong></p>"
	Else 
		If Len(request("pagenum")) = 0 Or clng(Request("pagenum")) < 1 Then 
			'Default on page 1
			oRs.AbsolutePage = 1
		Else 
			'Display currently selected page
			If clng(Request("pagenum")) <= clng(oRs.PageCount) Then 
				oRs.AbsolutePage = request("pagenum")
			Else 
				'Default to first page
				oRs.AbsolutePage = 1
			End If 
		End If 

		'Navigation Buttons
		response.write vbcrlf & "<table style=""margin-bottom: 5px;"">"
		response.write vbcrlf & "<tr>"
		response.write "<td valign=""bottom"">"
		response.write "<img border=""0"" src=""../images/nav_first.gif"" style=""cursor:pointer"" onclick=""location.href='" & sScriptName & "?pagenum=1" & sOptions & "'"" />"
		response.write "<img border=""0"" src=""../images/nav_back.gif"" style=""cursor:pointer"" onclick=""location.href='" & sScriptName & "?pagenum="&oRs.AbsolutePage-1 & sOptions & "'"" />"
		response.write "</td>"
		response.write "<td valign=""bottom"">"
		response.write "<strong>"
		response.write "Number of pages: <font style=""color:#0000ff;""> " & oRs.PageCount & "</font> | "
		response.write "Current page: <font  style=""color:#0000ff;"">" & oRs.AbsolutePage & "</font> | "
		response.write "Number of Records: <font  style=""color:#0000ff;"">" & oRs.RecordCount
		response.write "</strong>"
		response.write "</td>"
		response.write "<td valign=""bottom"">"
		response.write "<img border=""0"" src=""../images/nav_forward.gif"" valign=""bottom"" style=""cursor:pointer"" onclick=""location.href='" & sScriptName & "?pagenum="&oRs.AbsolutePage+1 & sOptions & "'"" />"
		response.write "<img border=""0"" src=""../images/nav_last.gif"" valign=""bottom"" style=""cursor:pointer"" onclick=""location.href='" & sScriptName & "?pagenum="&oRs.PageCount & sOptions & "'"" />"
		response.write "</td>"
		response.write "<td align=""right"" valign=""bottom"">"
		response.write "<input type=""button"" class=""excelexport button"" value=""Download as CSV"" onClick=""location.href='csv_export.asp'"" />"
		response.write "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<table cellspacing=""0"" cellpadding=""2"" class=""excel"" width=""100%"">"
		response.write vbcrlf & "<tr class=""excel"">"
		response.write "<td class=""excelheaderleft"">&nbsp;</td>"

		For Each fldLoop In oRs.Fields 
			response.write "<td class=""excelheader"">" & fldLoop.Name & "</td>"
		Next 

		' Total Document File size
		response.write "<td class=""excelheader"">Document<br />Size</td>"

		response.write "</tr>"

		'Set Base Record Count
		iRecordNumber = (oRs.AbsolutePage * iPageSize) - iPageSize

		For irows = 1 To oRs.PageSize
			If Not oRs.eof Then 
				bgcolor       = "#eeeeee"
				iRecordNumber = iRecordNumber + 1

				response.write vbcrlf & "<tr>"
				response.write "<td class=""exceldataleft"">" & iRecordNumber & "</td>"

				For Each fldLoop In oRs.Fields
					iType = fldLoop.type
					response.write "<td class=""exceldata"">" & FormatData(iType,trim(fldLoop.Value)) & "&nbsp;</td>"
				Next

				' Document size
				response.write "<td class=""exceldata"" nowrap=""nowrap"">" & GetTotalDocumentsSize(oRs("Org ID") ) & "</td>"

				response.write "</tr>"

				oRs.MoveNext
			End If 
		Next 	
		response.write vbcrlf & "</table>"
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' string FormatData( iType, sData )
'-------------------------------------------------------------------------------------------------
Function FormatData( ByVal iType, ByVal sData )
	Dim sReturnValue

	sReturnValue = sData

	'Format according to data type
	Select Case iType
		Case 6
			'Format display as Currency
			If sData = "" Or IsNull(sData) Then 
				'If NULL or EMPTY set to zero (0)
				sData = 0
			End If 

			sReturnValue = FormatCurrency(sData,2)
	End Select 

	FormatData = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' string GetTotalDocumentsSize( sOrgId )
'-------------------------------------------------------------------------------------------------
Function GetTotalDocumentsSize( ByVal sOrgId )
	Dim sSql, oRs

	sSql = "SELECT SUM(documentsize) AS documentsize "
	sSql = sSql & "FROM documents WHERE orgid = " & sOrgId
	sSql = sSql & " GROUP BY orgid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CDbl(oRs("documentsize")) > CDbl(1024) Then 
			GetTotalDocumentsSize = FormatNumber((CDbl(oRs("documentsize")) / CDbl(1024)),0)  & " KB"
		Else
			GetTotalDocumentsSize = FormatNumber(oRs("documentsize"),0) & " Bytes"
		End If 
	Else
		GetTotalDocumentsSize = "0 Bytes"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>
