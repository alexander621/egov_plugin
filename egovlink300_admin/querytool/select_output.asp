<!-- #include file="../includes/common.asp" //-->
<%

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "data query" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' IF FORM SUBMITTED DISPLAYED SELECTED CRITERIA
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	' ORDER CHANGE
	If request("SELECTLISTSTRING") <> "" THEN
			
			arrSelectList = split(request("SELECTLISTSTRING"),",")
			session("SELECTCLAUSE") = ""

			If NOT ISNULL(arrSelectList) Then
				For i=0 To UBOUND(arrSelectList)

					If (UBOUND(arrSelectList) > 1) AND i <> UBOUND(arrSelectList)-1 Then
						sComma = ","
					Else
						sComma = ""
					End If
					
					If trim(arrSelectList(i)) <> "" Then
						session("SELECTCLAUSE") = session("SELECTCLAUSE") & arrSelectList(i) & sComma
					End If
						
				Next

			End If

	ELSE
	
			' BUILD ORDER CLAUSE 
			If trim(session("SELECTCLAUSE")) = "" Then
				session("SELECTCLAUSE") = "[" & request("columnname") & "]"
			Else
				session("SELECTCLAUSE") = session("SELECTCLAUSE") & ",[" & request("columnname") & "]"
			End If

			' CLEAR ORDER INFORMATION
			If request("clearform") <> "" Then
				session("SELECTCLAUSE") = ""
			End If

	END IF 
End If
%>

<html>
<head>
	<title>E-Gov Application</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="querytool.css" />

	<script language="Javascript" src="dates.js"></script>
	<script language="Javascript" src="selectbox_script.js"></script>
</head>

<body>

<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">


<!--BEGIN: HEADER-->
<form name=frmdataselection action="select_output.asp" method="post">
<div class=title>QUERYTOOL:  Select Output</div>
<!--END: HEADER-->


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
			<a class=navh href="select_output.asp">Select Output</a>
			<a class=nav href="select_seq.asp">Select Sequencing</a>
			<a class=nav href="show_results.asp">Show Results</a>
		</td>
	</tr>
</table>
<!--END:  TAB SELECTIONS-->


<div class=group>

<table cellspacing=0 cellpadding=10>
<tr>
  <td valign=top>
	<b>Available Fields for Output:</b>
	<% Call DrawFieldSelection(session("view"))%>
  </td>
  <td valign=top>
    <table width=100%>
		<tr><td>
		<p>
			<div style="border: solid #000000 1px;">
			<table>
			<tr><td colspan=2><input class=smallbtn name=reset type=button value="Clear Selected Fields" onClick="document.frmdataselection.clearform.value='true';document.frmdataselection.submit();" ></td></tr>
			</table>
			</div>
			<input type=hidden name="clearform" value="">
			<input type=hidden name="columnname" value="">
			<input type=hidden name="SELECTCLAUSE" value="<%=session("SELECTCLAUSE")%>">
			<input type=hidden name="SELECTLISTSTRING" value="">
		</p>
		<b>Selected Fields:</b><br>
		<!--BEGIN: OUTPUT SELECT BOX-->
		<p><input class=smallbtn name=reset type=button value="MOVE UP" onClick="moveUpList(frmdataselection.selectlist) " >
		<input class=smallbtn name=reset type=button value="MOVE DOWN" onClick="moveDownList(frmdataselection.selectlist) " >
		<input class=smallbtn name=reset type=button value="REMOVE" onClick="removeFromList(frmdataselection.selectlist) " ></P>

		<select name=selectlist size=15 style="width:300px;">
			<%DrawSelectList(session("SELECTCLAUSE"))%>
		</option>
		<!--END:  OUTPUT SELECT BOX-->
       </td>
      </tr>
   </table>
   </td>
  </tr>
</table>

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
	
	response.write "<div style=""width:330px;height:305px; overflow:auto; border: solid #000000 1px;background-color:#eeeeee;"">" 
	response.write "<div style=""margin-right: 15px; overflow: none"">" 
	response.write "<table cellspacing=0 cellpadding=2>"
	response.write "<tr><td class=excelheader>&nbsp;</td><td align=left nowrap class=excelheader >Field Name</td><!--<td class=excelheader align=right nowrap>Field Size</td>--><td class=excelheader align=center nowrap>Field Type</td><!--<td class=excelheader align=right nowrap>Ordinal Pos</td>--></tr>"


	If NOT oSchema.EOF Then
		selecteditem = false
		Do while not selecteditem and not oSchema.EOF
			If request("columnname") = oSchema("column_name") Then
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
			If request("columnname") = oSchema("column_name") Then
				sSelected = "_selected"
			Else	
				sSelected = ""
			End If

			response.write "<tr onClick=""document.frmdataselection.columnname.value='" & oSchema("column_name") & "';document.frmdataselection.submit();"" onMouseOver=""this.style.backgroundColor='#FAF8CC';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td class=excelheader" & sSelected & ">&nbsp;</td><td align=left class=exceldata" & sSelected & ">" & oSchema("column_name") & "</td><!--<td align=right class=exceldata" & sSelected & ">" & sSize & "</td>--><td align=center class=exceldata" & sSelected & ">" & oSchema("data_type") & "</td><!--<td align=right class=exceldata" & sSelected & ">" & oSchema("ordinal_position") & "</td>--></tr>"
			oSchema.MoveNext
		Loop
		
		oSchema.MoveFirst

		selecteditem = false
		Do while not selecteditem and not oSchema.EOF
			If request("columnname") = oSchema("column_name") Then
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
				If request("columnname") = oSchema("column_name") Then
					sSelected = "_selected"
				Else	
					sSelected = ""
				End If
	
				response.write "<tr onClick=""document.frmdataselection.columnname.value='" & oSchema("column_name") & "';document.frmdataselection.submit();"" onMouseOver=""this.style.backgroundColor='#FAF8CC';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td class=excelheader" & sSelected & ">&nbsp;</td><td align=left class=exceldata" & sSelected & ">" & oSchema("column_name") & "</td><!--<td align=right class=exceldata" & sSelected & ">" & sSize & "</td>--><td align=center class=exceldata" & sSelected & ">" & oSchema("data_type") & "</td><!--<td align=right class=exceldata" & sSelected & ">" & oSchema("ordinal_position") & "</td>--></tr>"
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


Function DrawSelectList(sValue)

	arrList = split(sValue,",")

	If NOT ISNULL(arrList) Then
		For i=0 To UBOUND(arrList)
			response.write "<option>" & arrList(i)
		Next
	End If

End Function
%>

