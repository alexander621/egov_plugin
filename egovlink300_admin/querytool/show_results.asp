<!-- #include file="../includes/common.asp" //-->
<%

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "data query" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 
%>

<html>
<head>
	<title>E-Gov Application</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="querytool.css" />

	<script language="Javascript" src="dates.js"></script>

</head>

<body>

	<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<form name="frmcustomquery" action="show_results.asp?custom=true" method="post">
<div class=title>QUERYTOOL:  Select Output</div>


<!--BEGIN: SELECTION LIST-->
<font class=subtitle>Select Output (<%=session("view")%>):</font>
<br>
<img src=../images/spacer.gif height=16 width=1 border=0>
<br>


<!--BEGIN:  TAB SELECTIONS-->
<table cellpadding=0 cellspacing=0 border=0 class=navigation>
	<tr>
		<td valign=top align=right style="padding: 2px;">
			<!-- List of links at the top of the page -->
			<a class=nav href="default.asp">Select Input</a>
			<a class=nav href="selection_criteria.asp">Selection Criteria</a>
			<a class=nav href="select_output.asp">Select Output</a>
			<a class=nav href="select_seq.asp">Select Sequencing</a>
			<a class=navh href="show_results.asp">Show Results</a>
		</td>
	</tr>
</table>
<!--END:  TAB SELECTIONS-->

<div class=group>

<table cellspacing=0 cellpadding=10>
<tr>
  <td valign=top>
		<% Call DisplayInfo(session("view"))%>
  </td>
</tr>
</table>

<%
If trim(session("FULLQUERY"))  = "" Then
	sDISABLED = "DISABLED"
Else
	sDISABLED = ""
End If
%>

	<table>
	 <tr><td><input <%=sDISABLED%> class="button" type=button onClick="location.href='csv_export.asp'" value="Export to CSV File"></td></tr>
	</table>



</form>
</div>

<!--END: SELECTION LIST-->

</div>
</div>

<!--#Include file="../admin_footer.asp"-->

</body>

</html>


<%
' -------------------------------------------------------------------------------------------------
' BEGIN USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------
Function DisplayInfo(sValue)

	' ERROR CHECK FOR REQUIRED VALUES
	If trim(session("SELECTCLAUSE")) = "" and request("custom_query") = "" Then
		response.write "<b>Note:</b><i> No fields have been added to query. Click 'Select Output' and select at least one field to include in the query OR enter a 'saved query' in the 'custom query' box below.</i><p>"

		Response.write "Custom Query:<br><textarea name=custom_query class=query> " & session("FULLQUERY") & "</textArea><br><input class=""button"" onClick=""frmcustomquery.submit();"" type=button value=""Show Customized Query""><br><br></small>"
		Exit Function
	End If 

	' DECLARE OBJECTS	
	Set oSchema = Server.CreateObject("ADODB.Recordset")
	If request("custom") <> "" Then
		sSQL = TRIM(request("custom_query"))
		sQuery = "&custom=true"
		
		' IF NOT POST MUST BE NAVIGATION CLICK
		If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
			sSQL = session("FULLQUERY") 
		End If

		' CHECK STRING FORM WRONG DATA SOURCE
		sPosBegin = instr(sSQL,"FROM") + 5
		sPosEnd = instr(sSQL,"WHERE")
		
		If sPosEnd = o Then
			sPosEnd = instr(sSQL,"ORDER BY")
		End If

		If sPosEnd = 0 Then
			sPosEnd = Len(sSQL) + 1
		End If

		sReplaceString = Mid(sSQL,sPosBegin,(sPosEnd-sPosBegin))
		sValidView = "dbo.qry_action_line_" & session("orgid") & " "
		sSQL = replace(sSQL,sReplaceString,sValidView)

	Else

		If session("WHERECLAUSE") <> "" Then
			sAnd = " AND "
		Else
			sAnd = " WHERE "
		End If
		sSQL = "SELECT " & session("SELECTCLAUSE") & " FROM dbo." & sValue & " " & session("WHERECLAUSE") & sAnd & " [Organization Code]='" & session("orgid") & "' " &  session("ORDERCLAUSE")
		sQuery = ""
	End If


	' STORE QUERY IN SESSION FOR USE BY EXPORT SCRIPTS
	session("FULLQUERY") = sSQL


	' SET PAGE SIZE AND RECORDSET PARAMETERS
	oSchema.PageSize = 500
	oSchema.CacheSize = 500
	oSchema.CursorLocation = 3


	' OPEN RECORDSET
	'If UCASE(Left(sValue,4)) = "QDF2" Then
		'oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	'Else
		'oSchema.Open sSQL, DSN, 3, 1
	'End If
	on error resume next
	oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	
	querytoolerror = false
	if err.number <> 0 then
		querytoolerror = true
		querytoolerrornumber = err.number
		querytooldescription = err.description
		'response.write err.description
	end if

	If NOT oSchema.EOF THEN
	
	 ' SET PAGE TO VIEW
	 If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
		oSchema.AbsolutePage = 1
	 Else
		If clng(Request("pagenum")) <= oSchema.PageCount Then
			oSchema.AbsolutePage = Request("pagenum")
		Else
			oSchema.AbsolutePage = 1
		End If
	 End If

	
	' DISPLAY RECORDSET STATISTICS
	Dim abspage, pagecnt
	abspage = oSchema.AbsolutePage
	pagecnt = oSchema.PageCount

	'if querytoolerror then
		'response.write querytooldescription
	'end if

	' DISPLAY FORWARD AND BACKWARD NAVIGATION 
	Response.write "<small><b>Customize Query: </b><br>To add SQL commands/syntax not supported by the Querytool Builder's visual interface such a parenthesis modify text and press 'Show Customized Query'.<br>"
	if querytoolerror then
		response.write "<font color=RED>There was an error with the query you built.  Please correct: " & querytooldescription & "</font><br>"
	end if
	Response.write "<textarea name=custom_query class=query> " & session("FULLQUERY") & "</textArea><br><input class=""button"" onClick=""frmcustomquery.submit();"" type=button value=""Show Customized Query""><br><br></small>"

	if querytoolerror then response.end
	Response.write "<table><tr><td valign=top><a href=""show_results.asp?pagenum=1""><img border=0 src=""images/nav_first.gif""></a><a href=""show_results.asp?pagenum="&abspage - 1&sQuery&"""><img border=0 src=""images/nav_back.gif""></a></td>"
	Response.Write "<TD><B>Number of pages: <font color=blue> " & oSchema.PageCount & "</font> | " & vbcrlf
	Response.Write "Current page: <font color=blue>" & oSchema.AbsolutePage & "</font> | " & vbcrlf
	Response.Write "Record Count : <font color=blue>" & oSchema.RecordCount
	Response.write "</B></TD><td><a href=""show_results.asp?pagenum="&abspage + 1&sQuery&"""><img border=0 src=""images/nav_forward.gif"" valign=bottom></a><a href=""show_results.asp?pagenum="&oSchema.PageCount&"""><img border=0 src=""images/nav_last.gif"" valign=bottom></a></td></tr></table>"

	' BEGIN DISPLAYING REPORT RESULTS
	response.write "<div style=""width:668;height:225px; overflow:auto; border: solid #000000 1px;background-color:#eeeeee;"">" 
	response.write "<div style=""margin-right: 15px; overflow: none"">" 
	
	response.write "<table class=excel cellspacing=0 cellpadding=2>"

	If NOT oSchema.EOF Then
	
		' WRITE COLUMN HEADINGS
		response.write "<tr class=excel><td class=excelheader>&nbsp;</td>"
		For Each fldLoop in oSchema.Fields
			response.write "<td class=excelheader>" & fldLoop.Name & "</td>"
		Next
		response.write "</tr >"

		' WRITE DATA ROWS
		For intRec=1 To oSchema.PageSize
			If NOT oSchema.EOF Then
				response.write "<tr><td class=excelheader>"  & clng(oSchema.AbsolutePage - 1) * (oSchema.PageSize) + intRec & "</td>"
				For Each fldLoop in oSchema.Fields
					response.write "<td class=exceldata>" & trim(fldLoop.Value) & "&nbsp;</td>"
				Next
				response.write "</tr>"
				oSchema.MoveNext
			End If
		Next 

	End If

	response.write "</table></div></div>"

	Else

	response.write "<b><font color=red>No Matching Records Found</font></b>"

	End If

	Set oSchema = Nothing

End Function
%>


