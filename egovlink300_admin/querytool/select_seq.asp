<!-- #include file="../includes/common.asp" //-->

<%

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "data query" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' IF FORM SUBMITTED DISPLAYED SELECTED CRITERIA
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	' BUILD ORDER CLAUSE 
	If trim(session("ORDERCLAUSE")) = "" Then
		session("ORDERCLAUSE") = " ORDER BY [" & request("field1") & "] " & request("orderbyclause") 
	Else
		session("ORDERCLAUSE") = session("ORDERCLAUSE") & ",[" & request("field1") & "] " & request("orderbyclause") 
	End If

	' CLEAR ORDER INFORMATION
	If request("clearform") <> "" Then
		session("ORDERCLAUSE") = ""
	End If
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


<form name=frmdataselection action="select_seq.asp" method="post">
<div class=title>QUERYTOOL:  Select Output</div>


<!--BEGIN: SELECTION LIST-->
<font class=subtitle>Select Output (<%=session("view")%>):</font>
<br>
<img src=../images/spacer.gif height=16 width=1 border=0>
<br>


<!--BEGIN:  TAB SELECTIONS-->
<table cellpadding=0 cellspacing=0 border=0 width=100% class=navigation>
	<tr>
		<td valign=top align=right style="padding: 2px;">
			<!-- List of links at the top of the page -->
			<a class=nav href="default.asp">Select Input</a>
			<a class=nav href="selection_criteria.asp">Selection Criteria</a>
			<a class=nav href="select_output.asp">Select Output</a>
			<a class=navh href="select_seq.asp">Select Sequencing</a>
			<a class=nav href="show_results.asp">Show Results</a>
		</td>
	</tr>
</table>
<!--END:  TAB SELECTIONS-->


<div class="group">
<table cellspacing="0" cellpadding="10" border="0">
	<tr>
		<td valign=top>
		 <b>Available Fields for Sorting By:</b>
		<% Call DrawFieldSelection(session("view"))%>
		</td>
		<td valign=top>
			<table width=100%>
				<tr><td>
					<p>
					<b>Sequence Output Options:</b><br /><br />
					<div style="border: solid #000000 1px;">
					<table>
						<tr><td>Field Name:</td><td><input name=field1 value="<%=request("name")%>"></td></tr>
						<tr>
							<td>Option:</td>
							<td>
								<select name=orderbyclause>
								<option value="ASC">Ascending
								<option value="DESC">Descending
							</td>
						</tr>
						<tr>
						<td colspan=2><input class=smallbtn name=addordervalue type=button value="Add Sort Field" onClick="document.frmdataselection.submit();" > <input class=smallbtn name=reset type=button value="Clear Sort Fields" onClick="document.frmdataselection.clearform.value='true';document.frmdataselection.submit();" >
						</td></tr>
					</table>
					</div>
					<input type=hidden name="clearform" value="">
					<input type=hidden name="ORDERCLAUSE" value="<%=session("ORDERCLAUSE")%>">
					</p>
					<b>Sort Fields:</b><br>
					<textarea name=ORDERCLAUSE rows=12 cols=30><%=session("ORDERCLAUSE")%></textarea>
				</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</div>
</form>
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
Function DrawFieldSelection(sValue)
	
	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = "select column_name,ordinal_position,column_default,data_type,CHARACTER_MAXIMUM_LENGTH from information_schema.columns where table_name='" & sValue & "' ORDER BY ordinal_position"
	
	If UCASE(Left(sValue,4)) = "QDF2" Then
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	Else
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	End If

	If NOT oSchema.EOF Then
		'response.write "<b>Number of Fields: " & oSchema.Recordcount & "</b>"
	End If
	
	response.write "<div style=""width:375px;height:325px; overflow:auto; border: solid #000000 1px;background-color:#eeeeee;"">" 
	response.write "<div style=""margin-right: 15px; overflow: none"">" 
	response.write "<table cellspacing=0 cellpadding=2>"
	response.write "<tr><td class=excelheader>&nbsp;</td><td align=left nowrap class=excelheader >Field Name</td><!--<td class=excelheader align=right nowrap>Field Size</td>--><td class=excelheader align=center nowrap>Field Type</td><!--<td class=excelheader align=right nowrap>Ordinal Pos</td>--></tr>"


	If NOT oSchema.EOF Then
		selecteditem = false
		Do while not selecteditem and not oSchema.EOF
			If request("name") = oSchema("column_name") Then
				selecteditem = true
			else
				oSchema.MoveNext
			end if
		loop

		Do While NOT oSchema.EOF 
			sSize = oSchema("CHARACTER_MAXIMUM_LENGTH")
			Select Case oSchema("data_type")
			Case "int"
				sSize = "4"
			Case "bit"
				sSize = "1"
			Case "datetime"
				sSize = "8"
			Case "currency","money"
				sSize = "8"
			End Select
			
			' DETERMINE IF ROW IS SELECTED
			If request("name") = oSchema("column_name") Then
				sSelected = "_selected"
			Else	
				sSelected = ""
			End If
			
			response.write "<tr onClick=""location.href='select_seq.asp?name=" & oSchema("column_name") & "&type=" & oSchema("data_type") & "'"" onMouseOver=""this.style.backgroundColor='#FAF8CC';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td class=excelheader" & sSelected & ">&nbsp;</td><td class=exceldata" & sSelected & " align=left>" & oSchema("column_name") & "</td><!--<td class=exceldata" & sSelected & " align=right>" & sSize & "</td>--><td class=exceldata" & sSelected & " align=center>" & oSchema("data_type") & "</td><!--<td class=exceldata" & sSelected & "  align=right>" & oSchema("ordinal_position") & "</td>--></tr>"
			oSchema.MoveNext
		Loop
		
		oSchema.MoveFirst

		selecteditem = false
		Do while not selecteditem and not oSchema.EOF
			If request("name") = oSchema("column_name") Then
				selecteditem = true
			else
				sSize = oSchema("CHARACTER_MAXIMUM_LENGTH")
				Select Case oSchema("data_type")
				Case "int"
					sSize = "4"
				Case "bit"
					sSize = "1"
				Case "datetime"
					sSize = "8"
				Case "currency","money"
					sSize = "8"
				End Select
	
				' DETERMINE IF ROW IS SELECTED
				If request("name") = oSchema("column_name") Then
					sSelected = "_selected"
				Else	
					sSelected = ""
				End If
	
				response.write "<tr onClick=""location.href='select_seq.asp?name=" & oSchema("column_name") & "&type=" & oSchema("data_type") & "'"" onMouseOver=""this.style.backgroundColor='#FAF8CC';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td class=excelheader" & sSelected & ">&nbsp;</td><td class=exceldata" & sSelected & " align=left>" & oSchema("column_name") & "</td><!--<td class=exceldata" & sSelected & " align=right>" & sSize & "</td>--><td class=exceldata" & sSelected & " align=center>" & oSchema("data_type") & "</td><!--<td class=exceldata" & sSelected & "  align=right>" & oSchema("ordinal_position") & "</td>--></tr>"
				oSchema.MoveNext
			end if
		loop

	End If

	response.write "</table></div></div>"

	Set oSchema = Nothing

End Function

Function DrawDateChoices(sName)

	response.write "<select onChange=""getDates(document.frmdataselection." & sName & ".value);"" class=calendarinput Name=" & sName & ">"
	response.write "<option value=0>&nbsp;</option>"
	response.write "<option value=1>This Month</option>"
	response.write "<option value=2>Last Month</option>"
	response.write "<option value=3>This Quarter</option>"
	response.write "<option value=4>Last Quarter</option>"
	response.write "<option value=5>Last Year</option>"
	response.write "<option value=6>Year to Date</option>"
	response.write "<option value=7>All Dates to Date</option>"
	response.write "<option value=8>User specified</option>"
	response.write "</select>"

End Function
%>

