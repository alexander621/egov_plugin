<!-- #include file="../includes/common.asp" //-->

<%

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "data query" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' PROCESS COMBO SELECTION
If request("combo") <> "" Then
	If session("showcombo") <> "FALSE" Then
		session("showcombo") = "FALSE"
	Else
		session("showcombo") = "TRUE"
	End If
End If 


' IF FORM SUBMITTED DISPLAYED SELECTED CRITERIA
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	
		' APPEND TO EXISTING CLAUSE
		sFieldValue = request("fieldvalue1")
		sValue = UCASE(request("FIELDTYPE"))

		' DETERMINE FIELD DATA TYPE
		Select Case sValue

			Case "MONEY","INT"
			' NUMBERS

			Case "NVARCHAR","VARCHAR","TEXT"
			' TEXT - ADD WRAPPING '
			sFieldValue = "'" & request("fieldvalue1") & "'"

			Case "DATETIME","SMALLDATETIME"
			' DATES

		End Select

		
		' HANDLE SPECIAL SQL OPERATORS
		Select Case request("operation")
		
		Case "LIKE" 
			sFieldValue = "'%" & request("fieldvalue1") & "%'"
		Case "IS NULL"
			sFieldValue = ""
		Case "IS NOT NULL"
			sFieldValue = ""
		End Select

		
		' HANDLE DATE RANGES
		If sValue = "DATETIME" OR sValue = "SMALLDATETIME" Then

			' CHECK TO SEE IF WE NEED TO DISPLAY NULL DATES AS WELL AS DATE RANGE
			If request("nodate") <> "" Then
				sNullCheck = " OR [" & request("field1") & "] IS NULL" 
			Else
				sNullCheck = ""
			End If	
			
			' IS NULL/NOT NULL CHECKS
			If LEFT(request("dates"),2) = "IS" Then
				
				Select Case request("dates")
					Case "IS NULL"
						sFieldValue = ""
					Case "IS NOT NULL"
						sFieldValue = ""
				End Select
				
				sParms =  " ([" & request("field1") & "] " & request("dates") & " " & sFieldValue & ")"
			Else
				' DATE RANGE CHECKS
				sParms =  " ([" & request("field1") & "] >= '" & request("fromdate") & "' AND [" & request("field1") & "] <= '" & request("thrudate") & " 23:59:59.99'" & sNullCheck & ")"
			End If

		Else
			' NOT A DATETIME VALUE
			sParms =  " ([" & request("field1") & "] " & request("operation") & " " & sFieldValue & ")"
		End If

		' NEW OR EXISTING CLAUSE
		If session("WHERECLAUSE") <> "" Then
			' EXISTING
			session("WHERECLAUSE") = session("WHERECLAUSE") & " " & request("ANDORVALUE") & sParms
		Else
			' NEW
			session("WHERECLAUSE") = " WHERE " & sParms
		End If

	' CLEAR ORDER INFORMATION
	If request("clearform") <> "" Then
		session("WHERECLAUSE") = ""
	End If

End If

'DETERMINE WHICH DISPLAY OPTIONS TO TURN ON
If UCASE(request("type")) = "DATETIME" OR UCASE(request("type")) = "SMALLDATETIME" Then
	normalstyle = "display:none;"
	datestyle = ""
Else
	normalstyle = ""
	datestyle = "display:none;"
End If

%>

<html>
<head>
	<title>E-Gov Application</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="querytool.css" />

<script language="Javascript" src="dates.js"></script>

<script>
<!--
	function CheckForValue() 
	{
		if(document.frmdataselection.fieldvalue1.value != "" || document.frmdataselection.operation.value == "IS NOT NULL" || document.frmdataselection.operation.value == "IS NULL" || document.frmdataselection.Dates.value == "IS NOT NULL" || document.frmdataselection.Dates.value == "IS NULL" || (document.frmdataselection.fromdate.value != "" && document.frmdataselection.thrudate.value != "" )) 
		{
			document.frmdataselection.submit();
		}
	}
//-->
</script>

</head>

<body>


<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<form name=frmdataselection action="selection_criteria.asp" method="post">
<!--BEGIN: HEADER-->
<div class=title>QUERYTOOL:  Selection Criteria</div>
<!--END: HEADER-->


<!--BEGIN: SELECTION LIST-->
<font class=subtitle>Selection Criteria (<%=session("view")%>):</font>
<br>
<img src=../images/spacer.gif height=16 width=1 border=0>
<br>

<!--BEGIN:  TAB SELECTIONS-->
<table cellpadding=0 cellspacing=0 border=0 class=navigation>
	<tr>
		<td valign=top align=right style="padding: 2px;">
			<!-- List of links at the top of the page -->
			<a class=nav href="default.asp">Select Input</a>
			<a class=navh href="selection_criteria.asp">Selection Criteria</a>
			<a class=nav href="select_output.asp">Select Output</a>
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
		<b>Fields Available to Select By:</b><br><small><input type=checkbox name=showcombos <%If session("showcombo") <> "FALSE" Then response.write "CHECKED" %> onClick="location.href='selection_criteria.asp?combo=change';" > Show combos for code table fields.</small>
		<% Call DrawFieldSelection(session("view"))%>
		
		<br>

		<table width=100%>
		<tr><td>
		
		<div name=normalselections style="<%=normalstyle%>">
		<p>
		<b>Selection Options:</b><br /><br />
		<div style="border: solid #000000 1px;">
		<table>
		<tr><td>Field Name:</td><td><input name=field1 value="<%=request("name")%>"></td></tr>
		<tr><td>Operation:</td>
		<td>
			<select name=operation>
			<%DrawOperators(request("type"))%>
			</select>
		</td></tr>
		
		<%
		' DETERMINE TO WHETHER TO DISPLAY SELECT OR TEXT BOX
		If UCASE(request("type")) <> "BIT" THEN
					
			If request("name") <> "" Then
				
				' CHECK FOR COMBO FIELD IF COMBOS ON CHECKED 
				If (UCASE(request("name")) = "ACTION FORM NAME") OR (UCASE(request("name")) = "EMPLOYEE ASSIGNED NAME") OR (UCASE(request("name")) = "EMPLOYEE ASSIGNED NAME 2") OR (UCASE(request("name")) = "TRACKING NUMBER")  OR (UCASE(request("name")) = "EMPLOYEE ASSIGNED EMAIL") OR (UCASE(request("name")) = "STATUS") AND (session("showcombo") <> "FALSE") Then 
					
					sSelectList = LoadCombo(request("name"))
					

					If TRIM(sSelectList) <> "" Then
						' DISPLAY SELECT WITH SPECIFIED VALUES IN VIEW
						response.write "<tr><td>Value:</td><td><select name=fieldvalue1>"
						response.write sSelectList
						response.write "</select></td></tr>"
					Else
						' DISPLAY TEXT BOX
						response.write "<tr><td>Value:</td><td><input name=fieldvalue1 value=""""></td></tr>"
					End If
				Else
						' DISPLAY TEXT BOX
						response.write "<tr><td>Value:</td><td><input name=fieldvalue1 value=""""></td></tr>"
				End If

			End If
		Else
			' DISPLAY TRUE/FALSE SELECT
			response.write "<tr><td>Value:</td><td><select name=fieldvalue1><option SELECTED value="" ""> <option value=""1"">TRUE<option value=""0"">FALSE</select></td></tr>"
		End If
		%>


		
		</table>
		</div>
		</p>
		</div>



		<div name=dateselections style="<%=datestyle%>" >
		<p>
		<b>Date Selection Options:</b><br /><br />
		<div style="border: solid #000000 1px;">
		<table>
			<tr><td>Field Name:</td><td><input name=field2 value="<%=request("name")%>"></td></tr>
			<tr><td>Date Value Options:</td><td><%DrawDateChoices("Dates")%></td>
			<tr>
				<td><b>From Date</b></td>
				<td><input id=date1 name="fromdate" class=calendarinput size=10 >
				<input class=calendarbtn onclick="doDate('fromdate',1);" type=button value="..."></td>
			</tr>
			<tr>
				<td><b>Thru Date</b></td>
				<td><input id=date2 name="thrudate" class=calendarinput size=10 > <input    class=calendarbtn onclick="doDate('thrudate',1);" type=button value="..."></td>
			</tr>
			<tr>
				<td colspan=2><small>Include when NO date is present? <input type=checkbox name=nodate value=yes></small></td>
			</tr>
		</table>
		</div>
		</p>
		</div>
		
		</td>
		<td align=right valign=top><br>
		<p><input name="ADDAND" class=smallbtn type=button style="width:100px;" value="Add as AND >>" onClick="document.getElementById('date1').disabled=false;document.getElementById('date2').disabled=false;document.frmdataselection.ANDORVALUE.value='AND';CheckForValue();" ></p>
		
		<p><input class=smallbtn onClick="document.getElementById('date1').disabled=false;document.getElementById('date2').disabled=false;document.frmdataselection.ANDORVALUE.value='OR';CheckForValue();" name="ADDOR" type=button style="width:100px;" value="Add as OR >>"></p>
		</td>
		</tr>
		</table>
  </td>
  <td valign=top>
	<textarea name=DISPLAYWHERECLAUSE rows=18 cols=30><%=session("WHERECLAUSE")%></textarea>
	<br><input class=smallbtn name=reset type=button value="Clear Selection Criteria" onClick="document.frmdataselection.clearform.value='true';document.frmdataselection.submit();" >
  </td>
  
</tr>
</table>

<input type=hidden name="WHERECLAUSE" value="<%=session("WHERECLAUSE")%>">
<input type=hidden name="ANDORVALUE" value="">
<input type=hidden name="FIELDTYPE" value="<%=request("type")%>">
<input type=hidden name="clearform" value="">


</form>
</div>
</p>
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
	sSQL = "select column_name,ordinal_position,column_default,data_type,CHARACTER_MAXIMUM_LENGTH from information_schema.columns where table_name='" & sValue & "' order by ordinal_position"
	
	If UCASE(Left(sValue,4)) = "QDF2" Then
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	Else
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	End If

	If NOT oSchema.EOF Then
		'response.write "<b>Number of Fields: " & oSchema.Recordcount & "</b>"
	End If
	
	response.write "<div style=""width:375px;height:200px; overflow:auto; border: solid #000000 1px;background-color:#eeeeee;"">" 
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
			Case "int","real"
				sSize = "4"
			Case "bit"
				sSize = "1"
			Case "datetime","smalldatetime"
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

			response.write "<tr onClick=""location.href='selection_criteria.asp?name=" & oSchema("column_name") & "&type=" & oSchema("data_type") & "'"" onMouseOver=""this.style.backgroundColor='#FAF8CC';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td class=excelheader" & sSelected & ">&nbsp;</td><td class=exceldata" & sSelected & " align=left>" & oSchema("column_name") & "</td><!--<td class=exceldata" & sSelected & " align=right>" & sSize & "</td>--><td class=exceldata" & sSelected & " align=center>" & oSchema("data_type") & "</td><!--<td class=exceldata" & sSelected & "  align=right>" & oSchema("ordinal_position") & "</td>--></tr>"
			oSchema.MoveNext
		Loop

		oSchema.MoveFirst

		selecteditem = false
		Do while not selecteditem and not oSchema.EOF
			If request("name") = oSchema("column_name") Then
				selecteditem = true
				'response.write "HERE"
				'response.end
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
	
				response.write "<tr onClick=""location.href='selection_criteria.asp?name=" & oSchema("column_name") & "&type=" & oSchema("data_type") & "'"" onMouseOver=""this.style.backgroundColor='#FAF8CC';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td class=excelheader" & sSelected & ">&nbsp;</td><td class=exceldata" & sSelected & " align=left>" & oSchema("column_name") & "</td><!--<td class=exceldata" & sSelected & " align=right>" & sSize & "</td>--><td class=exceldata" & sSelected & " align=center>" & oSchema("data_type") & "</td><!--<td class=exceldata" & sSelected & "  align=right>" & oSchema("ordinal_position") & "</td>--></tr>"
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
	response.write "<option value=11>This Week</option>"
	response.write "<option value=12>Last Week</option>"
	response.write "<option value=2>Last Month</option>"
	response.write "<option value=1>This Month</option>"
	response.write "<option value=13>Next Month</option>"
	response.write "<option value=3>This Quarter</option>"
	response.write "<option value=4>Last Quarter</option>"
	response.write "<option value=5>Last Year</option>"
	response.write "<option value=6>Year to Date</option>"
	response.write "<option value=7>All Dates to Date</option>"
	response.write "<option value=8>User specified</option>"
	response.write "<option value=""IS NULL"">Date IS NULL</option>"
	response.write "<option value=""IS NOT NULL"">Date IS NOT NULL</option>"
	response.write "</select>"

End Function


Function DrawOperators(sValue)

sValue = UCASE(sValue)

Select Case sValue

	Case "BIT"
	' BOOLEAN
	response.write "<option value=""="">="
	response.write "<option value=""<>""><>"

	Case "MONEY","INT","TINYINT"
	' NUMBERS
	response.write "<option value=""="">="
	response.write "<option value=""<>""><>"
	response.write "<option value=""<""><"
	response.write "<option value="">"">>"
	response.write "<option value=""In"">In"

	Case "NVARCHAR","NTEXT","VARCHAR","TEXT"
	' TEXT
	response.write "<option value=""="">="
	response.write "<option value=""<>""><>"
	response.write "<option value=""<""><"
	response.write "<option value="">"">>"
	response.write "<option value=""LIKE"">LIKE"
	response.write "<option value=""In"">In"
	response.write "<option value=""IS NULL"">IS NULL"
	response.write "<option value=""IS NOT NULL"">IS NOT NULL"

	Case "DATETIME","SMALLDATETIME"
	' DATES

	Case Else
	' NONE
	response.write "<option value="" "">&nbsp;&nbsp;"

End Select

End Function


Function LoadCombo(sName)
	
	' GET LIST OF TABLE NAMES FIND MATCHING SET WITH VALUES
	blnFound = False
	sMatchName = ""
	sNameRoot = LEFT(sName,Len(sName) - 4) ' CODE
	sNameRootID = LEFT(sName,Len(sName) - 2) ' ID
	sList = ""

	' LOAD VIEW SCHEMA TO FIND MATCHING COLUMN
	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT TOP 1 * FROM " & session("view") 
	oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	

	' LOOP THRU COLUMNS AND FIND MATCHING COLUMN
	For i=0 to oSchema.fields.count-1
		' CODE EXTENSION FIELDS - (CODE MATCH DESC - OR - CODE MATCH DESCRIPTION)
		If len(sName) > 4 Then
			If (UCASE(LEFT(oSchema.fields(i).name,LEN(sNameRoot))) = UCASE(sNameRoot)) AND (ucase(oSchema.fields(i).name) <> ucase(sName)) AND ((UCASE(RIGHT(oSchema.fields(i).name,4)) = "DESC") OR UCASE(sNameRoot & "DESCRIPTION") = ucase(oSchema.fields(i).name)) Then
				blnFound = True	
				sMatchName = oSchema.fields(i).name
				
			End If
		End If
		' ID EXTENSION FIELDS - (ID MATCH FULLNAME)
		If len(sName) > 8 Then
			If (UCASE(LEFT(oSchema.fields(i).name,LEN(sNameRootID))) = UCASE(sNameRootID)) AND (ucase(oSchema.fields(i).name) <> ucase(sName)) AND (UCASE(RIGHT(oSchema.fields(i).name,8)) = "FULLNAME") Then
				blnFound = True	
				sMatchName = oSchema.fields(i).name
			End If
		End If
		' ID EXTENSION FIELDS - (ID MATCH NAME)
		If len(sName) > 4 Then
			If (UCASE(LEFT(oSchema.fields(i).name,LEN(sNameRootID))) = UCASE(sNameRootID)) AND (ucase(oSchema.fields(i).name) <> ucase(sName)) AND (UCASE(RIGHT(oSchema.fields(i).name,4)) = "NAME") Then
				blnFound = True	
				sMatchName = oSchema.fields(i).name
		    End If
		End If
	Next 

	Set oSchema = Nothing


	' SPECIAL HANDING
	
	' HANDLE ACTIONFORM NAMES
	If UCASE(sName) = "ACTION FORM NAME" Then
		blnFound = True
		sMatchName = "ACTION FORM NAME"
	End If

	If UCASE(sName) = "EMPLOYEE ASSIGNED NAME" Then
		blnFound = True
		sMatchName = "EMPLOYEE ASSIGNED NAME"
	End If

	If UCASE(sName) = "EMPLOYEE ASSIGNED NAME 2" Then
		blnFound = True
		sMatchName = "EMPLOYEE ASSIGNED NAME 2"
	End If

	If UCASE(sName) = "EMPLOYEE ASSIGNED EMAIL" Then
		blnFound = True
		sMatchName = "EMPLOYEE ASSIGNED EMAIL"
	End If


	If UCASE(sName) = "TRACKING NUMBER" Then
		blnFound = True
		sMatchName = "TRACKING NUMBER"
	End If

	If UCASE(sName) = "STATUS" Then
		blnFound = True
		sMatchName = "STATUS"
	End If


	' IF FOUND DISPLAY MATCH THEN GENERATE SELECT LIST 
	If blnFound = True Then
		
		Set oSchema = Server.CreateObject("ADODB.Recordset")
		sSQL = "SELECT DISTINCT [" & sMatchName  & "] , [" & sName & "] FROM " & session("view") & " WHERE [Organization Code] = '" & session("orgid") & "' ORDER BY [" & sMatchName  & "]"
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
		
		If NOT oSchema.EOF Then
			Do while NOT oSchema.EOF
				optiondisplay = oSchema.fields(0).value
				optionvalue = oSchema.fields(1).value
				sList = sList &  "<option value=""" & optionvalue & """>" & optiondisplay
			oSchema.MoveNext
			Loop
		End If

		Set oSchema = Nothing

	End If

	LoadCombo = sList

End Function
%>

