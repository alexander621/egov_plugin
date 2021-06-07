<!-- #include file="../includes/common.asp" //-->


<%

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "data query" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

' SET DEFAULT VIEW
session("view") = "qry_Action_Line_" & session("orgid")
%>


<html>
<head>
	<title>E-Gov Querytool</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="querytool.css" />
	<link rel="stylesheet" type="text/css" href="../global.css">

	<script src="layers.js"></script>

	<script language="JavaScript" type="text/javascript"> 
	<!--  
		function HandleForm()
		{ 
			var sForm = document.frmdataselection.view.options[document.frmdataselection.view.selectedIndex].text;
			
			if(sForm.indexOf('[PWD]') > 1)
			{
				// PASSWORD CHECK
				toggleDisplay('passbox');
				document.frmdataselection.passcheck.value = 'true';
			}
			else
			{
				// NO PASSWORD CHECK
				document.frmdataselection.submit();}
			} 
	// --> 
	</script> 
</head>

<body>

	<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">


<!--BEGIN: HEADER-->
<!--<div style="padding: 20px;">-->

<form name="frmdataselection" action="default.asp" method="post">
<div class=title>QUERYTOOL:  Selection Criteria</div>
<!--END: HEADER-->


<!--BEGIN: DATA SELECTION LIST-->
<input type=hidden name=passcheck value="false">



<!--BEGIN:  TAB SELECTIONS-->
<table cellpadding=0 cellspacing=0 border=0 class=navigation>
	<tr>
		<td valign=top align=right style="padding: 2px;">
			<!-- List of links at the top of the page -->
			<a class=navh href="default.asp">Select Input</a>
			<a class=nav href="selection_criteria.asp">Selection Criteria</a>
			<a class=nav href="select_output.asp">Select Output</a>
			<a class=nav href="select_seq.asp">Select Sequencing</a>
			<a class=nav href="show_results.asp">Show Results</a>
		</td>
	</tr>
</table>
<!--END:  TAB SELECTIONS-->


<!--BEGIN:  DATA TABLE/VIEW SOURCE SELECTION-->
<div class=group>
<!--<b>Select data module to work with: </b><br>
<select name="view" onChange="HandleForm();" >
<option value="Action Line Query">Action Line Query</option>-->
<!--<option value="00000">Please select data source...</option>
<optgroup ID="ACTIONLINE" LABEL="ACTION LINE...">
<% 'Call ListTargetInput()%>
</optgroup>-->
<!--</select>-->
<input type=hidden name=qdfpass value="">


<!--BEGIN: PASSWORD CHECK-->
<p>
<div id=passbox name=passbox style="display:none;" >
<b>This information is password protected. Please enter the required password to continue.</b></font><br><input type=password value="" name=sPassBox><input class=smallbtn type=button onClick="document.frmdataselection.submit()" value="Continue">
</div></p>
<!--END: PASSWORD CHECK-->


</form>
<!--END:  DATA TABLE/VIEW SOURCE SELECTION-->


<%
' IF FORM SUBMITTED DISPLAY COLUMN INFORMATION FOR SELECTED VIEW/TABLE
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

	session("view") = ""

	If CheckPass(request("sPassBox")) OR request("passcheck") = "false" Then

		session("view") = request("view")
		session("WHERECLAUSE") = ""
		session("ORDERCLAUSE") = ""
		session("SELECTCLAUSE") = ""
	
		If request("view") <> "00000" Then%>
			<font class=subtitle><%=UCASE(request("view"))%></font>
			<p><% GetFieldInformation(request("view"))%></p>
		<%End If
	
	End If
Else

		Response.write "<form name=""frmcustomquery"" action=show_results.asp?custom=true method=""post"">"
		Response.write "Custom Query:<br><textarea name=custom_query class=query> " & session("FULLQUERY") & "</textArea><br><input onClick=""frmcustomquery.submit();"" type=button class=""button"" value=""Show Customized Query""><br><br></small>"
		Response.write "</form>"
End If

' SECURITY PASSWORD CHECK
If (request("passcheck") = "true" AND NOT CheckPass(request("sPassBox"))) THen
	response.write "<b><font color=red> << INVALID PASSWORD >> </font></b>"
End IF
%>
</div>
</p>
<!--END: DATA SELECTION LIST-->


</div>
</div>

<!--#Include file="../admin_footer.asp"-->

</body>

</html>


<%
' -------------------------------------------------------------------------------------------------
' BEGIN USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------
Function ListTargetInput()
	
	Set oReportList = Server.CreateObject("ADODB.Recordset")
	sSQL = "sp_tables @table_type = ""'VIEW'"""
	oReportList.Open sSQL, Application("QUERYTOOLDSN"), 3, 1

	If NOT oReportList.EOF Then

		Do While NOT oReportList.EOF 
			If Left(oReportList("TABLE_NAME"),3) = "qry" Then
				sTableName = oReportList("TABLE_NAME")
				response.write "<option value=""" & oReportList("TABLE_NAME") & """>" & sTableName & "</option>"
			End If
			oReportList.MoveNext
		Loop

	End If

	Set oReportList = Nothing

End Function



Function GetFieldInformation(sValue)
	
	Set oSchema = Server.CreateObject("ADODB.Recordset")
	sSQL = "select column_name,ordinal_position,column_default,data_type,CHARACTER_MAXIMUM_LENGTH from information_schema.columns where table_name='" & sValue & "' ORDER BY ordinal_position"
	
	If UCASE(Left(sValue,4)) = "QDF2" Then
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	Else
		oSchema.Open sSQL, Application("QUERYTOOLDSN"), 3, 1
	End If

	If NOT oSchema.EOF Then
		response.write "<b>Number of Fields: " & oSchema.Recordcount & "</b>"
	End If
	
	response.write "<div style=""width:600px;height:225px; overflow:auto; border: solid #000000 1px;background-color:#eeeeee;"">" 
	response.write "<div style=""margin-right: 15px; overflow: none"">" 
	response.write "<table cellspacing=0 cellpadding=2>"
	response.write "<tr><td class=excelheader>&nbsp;</td><td align=left nowrap class=excelheader >Field Name</td><td class=excelheader align=right nowrap>Field Size</td><td class=excelheader align=center nowrap>Field Type</td><td class=excelheader align=right nowrap>Ordinal Pos</td></tr>"


	If NOT oSchema.EOF Then

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
			
			response.write "<tr><td class=excelheader" & sSelected & ">&nbsp;</td><td class=exceldata" & sSelected & " align=left>" & oSchema("column_name") & "</td><td class=exceldata" & sSelected & " align=right>" & sSize & "</td><td class=exceldata" & sSelected & " align=center>" & oSchema("data_type") & "</td><td class=exceldata" & sSelected & "  align=right>" & oSchema("ordinal_position") & "</td></tr>"
			oSchema.MoveNext
		Loop

	End If

	response.write "</table></div></div>"

	Set oSchema = Nothing

End Function



Function CheckPass(sValue)
	blnReturnValue = False

	' NEED TO REPLACE WITH DATABASE CALL
	If sValue = "Smile" Then
		blnReturnValue = True
	End If

	CheckPass = blnReturnValue
End Function


%>

