<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_statisticsreport.asp
' AUTHOR: Steve Loar
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the statistics report for classes and events
'
' MODIFICATION HISTORY
' 1.0   05/09/06   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iYear, iAdminCnt, iOnlineCnt

If request("iyear") <> "" Then 
	iYear = request("iyear")
Else
	iYear = Year(Now())
End If 

iAdminCnt = 0
iOnlineCnt = 0

%>


<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link href="../global.css" rel="stylesheet" type="text/css" />
	<link href="classes.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />
</head>


<body>

 
<%DrawTabs tabRecreation,1%>


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<h3><%=GetOrgName( Session("orgid") )%> Class/Events Statistics</h3>

		<div id="receiptlinks">
			<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
		</div>
		<div id="topbuttons">
			<input type="button" onclick="javascript:window.print();" value="Print" />
		</div>

		<p>
			<form name="YearForm" method="post" action="class_statisticsreport.asp">
			Year: <% ShowYearChoices( iYear ) %>
			</form>
		</p>
		<p>
			Percentage of programs meeting minimum registrations: <% =GetMinimumRegistrationPercents( iYear ) %>
		</p>
		<p>
			Percentage of programs meeting maximum registrations: <% =GetMaximumRegistrationPercents( iYear ) %>
		</p>
		<p>
			Class count by category: <br /><br />
			<% ShowClassByCategory iYear %>
		</p>
		<p>
			Total revenue for the year: <% =FormatCurrency(GetClassRevenue( iYear ),2) %>
		</p>
		<p>
			Revenue per program: <% =FormatCurrency(GetRevenuePerClass( iYear ),2) %>
		</p>
		<p>
			Number of participants in all programs: <% =GetClassParticipantsCount( iYear )%>
		</p>
		
		<% GetRegistrationByType iYear, iAdminCnt, iOnlineCnt %>
		<p>
			Number of participants registering online: <%=iOnlineCnt %>
		</p>
		<p>
			Number of participants signed up by administration personnel: <%=iAdminCnt %>
		</p>
		<p>
			Number of evaluations sent: 0
		</p>
		<p>
			Number of evaluations returned: 0
		</p>

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>


</html>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' Sub ShowYearChoices( iDefaultYear )
'--------------------------------------------------------------------------------------------------
Sub ShowYearChoices( iDefaultYear )
	Dim sSql, oYears

	sSql = "Select distinct year(startdate) as year from egov_class where startdate is not null and orgid = " & session("orgid")

	Set oYears = Server.CreateObject("ADODB.Recordset")
	oYears.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""iyear"">"

	Do While NOT oYears.EOF 
		response.write vbcrlf & "<option value=""" & oYears("year") & """ "  
		If clng(iDefaultYear) = clng(oYears("year")) Then
			response.write " selected=""selected"" "
		End If 
		response.write " >" & oYears("year") & "</option>"
		oYears.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oYears.close
	Set oYears = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetMinimumRegistrationPercents( iYear )
'--------------------------------------------------------------------------------------------------
Function GetMinimumRegistrationPercents( iYear )
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetMinimumRegistrationPercent"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		.Parameters.Append oCmd.CreateParameter("@iMinPercent", 5, 2)
		.Execute
	End With
	
	GetMinimumRegistrationPercents = oCmd.Parameters("@iMinPercent").Value & "%"

	Set oCmd = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetMaximumRegistrationPercents( iYear )
'--------------------------------------------------------------------------------------------------
Function GetMaximumRegistrationPercents( iYear )
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetMaximumRegistrationPercent"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		.Parameters.Append oCmd.CreateParameter("@iMaxPercent", 5, 2)
		.Execute
	End With
	
	GetMaximumRegistrationPercents = oCmd.Parameters("@iMaxPercent").Value & "%"

	Set oCmd = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassRevenue( iYear )
'--------------------------------------------------------------------------------------------------
Function GetClassRevenue( iYear )
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetClassRevenueByYear"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		.Parameters.Append oCmd.CreateParameter("@iTotalRevenue", 5, 2)
		.Execute
	End With
	
	GetClassRevenue = oCmd.Parameters("@iTotalRevenue").Value 

	Set oCmd = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetRevenuePerClass( iYear )
'--------------------------------------------------------------------------------------------------
Function GetRevenuePerClass( iYear )
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetClassRevenuePerClass"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		.Parameters.Append oCmd.CreateParameter("@iRevenuePerClass", 5, 2)
		.Execute
	End With
	
	GetRevenuePerClass = oCmd.Parameters("@iRevenuePerClass").Value 

	Set oCmd = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassParticipantsCount( iYear )
'--------------------------------------------------------------------------------------------------
Function GetClassParticipantsCount( iYear )
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetClassParticipantsByYear"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		.Parameters.Append oCmd.CreateParameter("@iTotalParticipants", 3, 2)
		.Execute
	End With
	
	GetClassParticipantsCount = oCmd.Parameters("@iTotalParticipants").Value

	Set oCmd = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassByCategory( iYear )
'--------------------------------------------------------------------------------------------------
Sub ShowClassByCategory( iYear )
	Dim oCmd, oClass, iRow

	iRow = 0

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetClassCountByCategory"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		Set oClass = .Execute
	End With
	
	response.write vbcrlf & "<div id=""classcount"">"
	response.write vbcrlf & "<table id=""classcount"" cellpadding=""5"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr><th>Category</th><th>Classes</th></tr>"
	Do While Not oClass.EOF
		iRow = iRow + 1
		If iRow Mod 2 = 1 Then 
			response.write vbcrlf & "<tr>"
		Else
			response.write vbcrlg & "<tr class=""alt_row"">"
		End If 
		response.write "<td>" & oClass("categorytitle") & "</td><td>" & oClass("classesofferedcount") & "</td></tr>"
		oClass.MoveNext
	Loop 
	response.write vbcrlf & "</table>"
	response.write vbcrlf & "</div>"

	oClass.close 
	Set oClass = Nothing 
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassByCategory( iYear )
'--------------------------------------------------------------------------------------------------
Sub GetRegistrationByType( iYear, ByRef iAdminCnt, ByRef iOnlineCnt )
	Dim  oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetRegistrationsByType"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgId", 3, 1, 4, session("orgid"))
		.Parameters.Append oCmd.CreateParameter("@iYear", 3, 1, 4, iYear)
		.Parameters.Append oCmd.CreateParameter("@iOnlineTotal", 3, 2)
		.Parameters.Append oCmd.CreateParameter("@iAdminTotal", 3, 2)
		.Execute
	End With
	
	iAdminCnt = oCmd.Parameters("@iAdminTotal").Value
	iOnlineCnt = oCmd.Parameters("@iOnlineTotal").Value

	Set oCmd = Nothing

End Sub 
%>


