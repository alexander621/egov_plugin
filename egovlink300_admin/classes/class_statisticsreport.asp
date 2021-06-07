<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
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
' 1.0   05/09/06	Steve Loar - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "class statistics" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

Dim iYear, iAdminCnt, iOnlineCnt, dAnnualRevenue

If request("iyear") <> "" Then 
	iYear = request("iyear")
Else
	iYear = Year(Now())
End If 

iAdminCnt = 0
iOnlineCnt = 0
dAnnualRevenue = 0

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="Javascript">
	<!--
	function reloadpage()
	{
		//var iYear = document.frmcal.selyear.options[document.frmcal.selyear.selectedIndex].value;
		//location.href='facility_calendar.asp?iYear=' + iYear ;
		document.YearForm.submit();
	}
	//-->
	</script>

	<script defer>
	function window.onload() 
	{
	  //factory.printing.header = ''
	  //factory.printing.footer = '&b<%=Session("sOrgName")%> Class/Events Statistics - Printed on &d - Page:&p/&P'
	  //factory.printing.portrait = true
	  //factory.printing.leftMargin = 0.5
	  //factory.printing.topMargin = 0.5
	  //factory.printing.rightMargin = 0.5
	  //factory.printing.bottomMargin = 0.5
	 
	  // enable control buttons
	  //var templateSupported = factory.printing.IsTemplateSupported();
	  //var controls = idControls.all.tags("input");
	  //for ( i = 0; i < controls.length; i++ ) {
		//controls[i].disabled = false;
		//if ( templateSupported && controls[i].className == "ie55" )
		//  controls[i].style.display = "inline";
	  //}
	}
	</script>

</head>

<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
<%
'	<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
'	<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
%>
</div>

<%
'<object id="factory" viewastext  style="display:none"
'  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
'   codebase="../includes/smsx.cab#Version=6,3,434,12">
'</object>
%>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<h3><%=GetOrgName( Session("orgid") )%> Class/Events Statistics</h3>

		<!--<div id="receiptlinks">
			<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
		</div>-->

		<!--<div id="topbuttons">
			<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
		</div>-->

		<p>
			<form name="YearForm" method="post" action="class_statisticsreport.asp">
			Year: <% ShowYearChoices iYear %>
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
		
		<% If dAnnualRevenue > 0 Then %>
			<% dAnnualRevenue = GetRevenuePerClass( iYear )%>
			<p>
				Revenue per program: <% =FormatCurrency(dAnnualRevenue,2) %>
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
			<!--
			<p>
				Number of evaluations returned: 0
			</p>
			-->
		<% End If %>

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

	response.write vbcrlf & "<select name=""iyear"" onChange='reloadpage();'>"
	
	If Not oYears.EOF Then 
		Do While NOT oYears.EOF 
			response.write vbcrlf & "<option value=""" & oYears("year") & """ "  
			If clng(iDefaultYear) = clng(oYears("year")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oYears("year") & "</option>"
			oYears.MoveNext
		Loop
	Else
		response.write vbcrlf & "<option value=""" & Year(Date()) & """  selected=""selected"" >" & Year(Date()) & "</option>"
	End If 

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
	Dim cPurchaseAmount, cRefundAmount

	cPurchaseAmount = GetYearAmount( "credit", iYear, 1 )
	'response.write cPurchaseAmount & "<br />"
	cRefundAmount = GetYearAmount( "debit", iYear, 2 )
	'response.write cRefundAmount & "<br />"

	GetClassRevenue = CDbl(cPurchaseAmount) - CDbl(cRefundAmount)
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetYearAmount( sType, iYear )
'--------------------------------------------------------------------------------------------------
Function GetYearAmount( sType, iYear, iJournalEntryTypeID )
	Dim sSql, oAmount
	' and journalentrytypeid = 2

	sSql = "select sum(L.amount) as amount from egov_accounts_ledger L, egov_class_payment J "
	sSql = sSql & " where L.orgid = " & session("orgid") & " and L.itemtypeid = 1 and L.ispaymentaccount = 0 and L.entrytype = '" & sType & "' "
	sSql = sSql & " and L.paymentid = J.paymentid and Year(J.paymentdate) = " & iYear & " and journalentrytypeid = " & iJournalEntryTypeID
	sSql = sSql & " group by L.orgid"

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 0, 1

	If Not oAmount.EOF Then 
		GetYearAmount = CDbl(oAmount("amount"))
	Else
		GetYearAmount = CDbl(0.00)
	End If 

	oAmount.Close
	Set oAmount = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassRevenue_old( iYear )
'--------------------------------------------------------------------------------------------------
Function GetClassRevenue_old( iYear )
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
	If IsNull(GetClassRevenue) Then 
		GetClassRevenue = 0
	End If 

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
' Sub GetRegistrationByType( iYear, ByRef iAdminCnt, ByRef iOnlineCnt )
'--------------------------------------------------------------------------------------------------
Sub GetRegistrationByType( ByVal iYear, ByRef iAdminCnt, ByRef iOnlineCnt )
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


