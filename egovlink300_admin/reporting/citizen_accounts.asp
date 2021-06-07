<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: citizen_accounts.asp
' AUTHOR: SteveLoar
' CREATED: 11/12/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This report displays the accounts receivables and accounts payable for citizens. 
'				Part of the Menlo Park Project.
'
' MODIFICATION HISTORY
' 1.0   11/12/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "citizen accounts rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>
<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--

		window.onload = function()
		{
		  //factory.printing.header = "Printed on &d"
		  //factory.printing.footer = "&bPrinted on &d - Page:&p/&P";
		  //factory.printing.portrait = true;
		  //factory.printing.leftMargin = 0.5;
		  //factory.printing.topMargin = 0.5;
		  //factory.printing.rightMargin = 0.5;
		  //factory.printing.bottomMargin = 0.5;
		 
		  // enable control buttons
		  //var templateSupported = factory.printing.IsTemplateSupported();
		  //var controls = idControls.all.tags("input");
		  //for ( i = 0; i < controls.length; i++ ) 
		  //{
			//controls[i].disabled = false;
			//if ( templateSupported && controls[i].className == "ie55" )
			//  controls[i].style.display = "inline";
		  //}
		}

	  //-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

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

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Citizen Accounts Report</strong></font></td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
				
				<p class="executivesummary">
					Citizen Accounts Receivable..........................<%=GetAccountsReceivable( ) %><br />
					Citizen Accounts Payable..............................<%=GetAccountsPayable( ) %><br />
				</p>
				<!-- END: DISPLAY RESULTS -->
      
			</td>
		 </tr>
	</table>
  </form>
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
' Function GetAccountsReceivable( )
'--------------------------------------------------------------------------------------------------
Function GetAccountsReceivable( )
	Dim sSql, oRs, sAmount

	sSql = "SELECT SUM(accountbalance) AS accountbalance FROM egov_users WHERE orgid = " & session("orgid") & " AND accountbalance < 0"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("accountbalance")) Then
			sAmount = 0.00
		Else 
			sAmount = Abs(oRs("accountbalance"))
		End If 
	Else
		sAmount = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetAccountsReceivable = FormatNumber(sAmount,2)
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAccountsPayable( )
'--------------------------------------------------------------------------------------------------
Function GetAccountsPayable( )
	Dim sSql, oRs, sAmount

	sSql = "SELECT SUM(accountbalance) AS accountbalance FROM egov_users WHERE orgid = " & session("orgid") & " AND accountbalance > 0"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("accountbalance")) Then
			sAmount = 0.00
		Else 
			sAmount = oRs("accountbalance")
		End If 
	Else
		sAmount = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 
	
	GetAccountsPayable = FormatNumber(sAmount,2)
End Function 


%>