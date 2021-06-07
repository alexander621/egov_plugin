<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: gl_account_mgmt.asp
' AUTHOR: Steve Loar
' CREATED: 02/14/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the management of gl accounts
'
' MODIFICATION HISTORY
' 1.0   02/14/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMaxAccountId, iAccountCount, sDisplay, sDisplayNew, sShow

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "gl accounts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iMaxAccountId = GetMaxAccountId()
iAccountCount = 0

If Request("display") = "" Or request("display") = "A" Then
	sDisplay = "A"
	sDisplayNew = "D"
	sShow = "Show All"
Else
	sDisplay = "D"
	sDisplayNew = "A"
	sShow = "Show Active Only"
End If 

%>


<html>
<head>
	<title>E-Gov GL Account Management</title>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<style>
		#content * { font-family:'Open Sans', sans-serif; font-size:14px; }
		.fa {font: normal normal normal 14px/1 FontAwesome !important;}
		.ui-button, .ui-button:hover, .ui-button:active, .ui-button:focus { font-size:14px;background-color: #2C5F93 !important; color:white !important; border:inherit !important; margin-bottom:5px; padding: .4em 1em !important; }
		.ui-button:disabled {
			background-color: #ccc;
			color:#999;
			margin-bottom:5px;
			cursor: not-allowed;
		}
		div.glaccounts_shadow {background:none; width:100%;}
		div.glaccounts_shadow table {width:100% !important;}

.dropdown {
    position: relative;
    display: inline-block;
    margin-left:40px;

}
.dropdown.right {float:right;}
#bottom .dropdown
{
	display:none;
}

.dropdown-content {
    display: none;
    position: absolute;
    background-color: #f1f1f1;
    min-width: 280px;
    box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
    z-index: 1;
}
.dropdown.right .dropdown-content {
    margin-left:-198px;
	margin-top:-5px;
}

.dropdown-content a {
    color: black;
    padding: 12px 16px;
    text-decoration: none;
    display: block;
}

.dropdown-content a:hover {background-color: #ddd}

.dropdown:hover .dropdown-content {
    display: block;
}

.dropdown:hover .dropbtn, .ui-state-active {
    background-color: #2C5F93 !important;
    
}
.dd-green, .dd-green:hover
{
	background-color: green !important;
}
.showiniframe {display:none;}
	</style>

	<script language="Javascript">
	<!--

		function DeactivateCheck( sDisplay )
		{
			var sMsg = 'Are you certain you want to deactivate the marked accounts?';
			if (sDisplay == 'D')
			{
				sMsg = 'Are you certain you want to change the status of the marked accounts?';
			}
			if(confirm(sMsg)) 
			{
				// tell the save routine to work in delete mode
				document.AccountForm.action.value = 'deactivate';
				//alert(document.AccountForm.action.value);
				document.AccountForm.submit();
			}
		}

		function SaveChanges()
		{
			// check all the accounts so they get processed in the save routine
			for (var j = 0; j <= <%=iMaxAccountId - 1%>; j++) 
			{
				var exists = eval("document.AccountForm.accountid[" + j + "]");
				if (exists)
				{
					box = eval("document.AccountForm.accountid[" + j + "]"); 
					if (box.checked == false) box.checked = true;
				}
			}
			document.AccountForm.submit();
		}

		function ChangeShow( sDisplayNew )
		{
			//alert(sDisplayNew);
			location.href='gl_account_mgmt.asp?display=' + sDisplayNew;
		}
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}
		function commonIFrameUpdateFunction()
		{
			UpdateParentGLAccounts('glaccounts','glaccountDD')
		}
		function UpdateParentGLAccounts(poptype, classname)
		{

			//Get New Values
			var request = new XMLHttpRequest();
			request.open('GET', 'popselectbox.asp?type='+poptype, false);  // `false` makes the request synchronous
			request.send();

			if (request.status === 200) {
  				newDDVals = request.responseText;

				//Get elements from parent
				var pfDD = parent.document.getElementsByClassName(classname);
				for (var i = 0; i < pfDD.length; i++) {
					//Get Selected Value
  					//pfDD[i].style.display = 'inline-block';
					var selVal = pfDD[i].options[pfDD[i].selectedIndex].value;
					
					//Update The Values
					pfDD[i].innerHTML = newDDVals;
	
					//Select Previous Option
					pfDD[i].value = selVal;
				}
			}

		}


	//-->
	</script>

</head>


<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<font size="+1"><strong>GL Account Management</strong></font><br />
		
		<form name="AccountForm" method="post" action="save_accounts.asp">
		<br />
			<input type="hidden" name="action" value="save" />
				<input type="button" class="button ui-button ui-widget ui-corner-all" name="create" value="New Account" onClick="javascript:window.location='new_account.asp';" /> &nbsp;
				<% If iMaxAccountId > 0 Then %>
					<input type="button" class="button ui-button ui-widget ui-corner-all" name="save1" value="Save Changes" onClick="SaveChanges();" /> &nbsp; 
					<input type="button" class="showiniframe button ui-button ui-widget ui-corner-all" value="Close" onClick="doClose();" />
     					<div class="dropdown right">
  						<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-bars" aria-hidden="true"></i> Tools</button>
  						<div class="dropdown-content">
							<a href="../export/excel_export.asp">Export to EXCEL</a>
							<a href="javascript:ChangeShow('<%=sDisplayNew%>');"><%=sShow%></a>
							<a href="javascript:DeactivateCheck('<%=sDisplay%>');">Deactivate/Activate Marked</a>
						</div>
					</div>
				<% End If %>
				
			<p>
				<div class="glaccounts_shadow">
					<table id="glaccounts" border="0" cellpadding="3" cellspacing="0">
						<tr><th>Deactivate/Activate</th><th>Account #</th><th>Account Name</th><th>Status</th></tr>
						<%	If iMaxAccountId > 0 Then
								iAccountCount = ShowAccounts( sDisplay )
							Else %>
								<tr><td colspan="3"> &nbsp; No Accounts Exist.</td></tr>
						<%	End If %>
					</table>
				</div>
			</p>
			<p>
				<% If iAccountCount > 29 Then %>
					<input type="button" class="button" name="save2" value="Save Changes" onClick="SaveChanges();" />
					<input type="button" class="button" name="deactivate2" value="Deactivate/Activate Marked" onClick="DeactivateCheck('<%=sDisplay%>');" /> &nbsp;
				<% End If %>
			</p>
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
' Function ShowAccounts( sDisplay )
'--------------------------------------------------------------------------------------------------
Function ShowAccounts( sDisplay )
	Dim sSql, oAccounts, iRows, sStatusFilter

	iRows = 0
	If sDisplay = "A" Then
		sStatusFilter = " and accountstatus <> 'D' "
	Else
		sStatusFilter = ""
	End If 

	sSql = "Select accountid, accountname, accountnumber, isnull(accountstatus,'') as accountstatus from egov_accounts where orgid = " & session("orgid") & sStatusFilter & " Order By accountname, accountnumber"
	' BEGIN: STORE QUERY FOR EXPORT TO CSV
	session("DISPLAYQUERY") = "SELECT '=""' + accountnumber + '""' as [account number], accountname as [account name], isnull(accountstatus,'') as [account status] from egov_accounts where orgid = " & session("orgid") & sStatusFilter & " Order By accountname, accountnumber"
	' END: STORE QUERY FOR EXPORT TO CSV


	Set oAccounts = Server.CreateObject("ADODB.Recordset")
	oAccounts.Open sSQL, Application("DSN"), 0, 1

	Do While Not oAccounts.EOF
		iRows = iRows + 1
		response.write vbcrlf & "<tr"
		If iRows Mod 2 = 0 Then response.write " class=""altrow"" "
		response.write "><td align=""center""><input type=""checkbox"" name=""accountid"" value=""" & oAccounts("accountid") & """></td>"
		response.write "<td align=""center""><input type=""text"" name=""accountnumber" & oAccounts("accountid") & """ value=""" & Replace(oAccounts("accountnumber"),"""","&quot;") & """ size=""20"" maxlength=""20"" /></td>"
		response.write "<td align=""center""><input type=""text"" name=""accountname" & oAccounts("accountid") & """ value=""" &  Replace(oAccounts("accountname"),"""","&quot;") & """ size=""50"" maxlength=""50"" /></td>"
		response.write "<td align=""center"">" 
		If oAccounts("accountstatus") = "A" Then
			response.write "Active<input type=""hidden"" name=""accountstatus" & oAccounts("accountid") & """ value=""A"" />"
		Else
			response.write "Deactivated<input type=""hidden"" name=""accountstatus" & oAccounts("accountid") & """ value=""D"" />"
		End if
		response.write "</td>"
		response.write "</tr>"
		oAccounts.MoveNext
	Loop

	oAccounts.close
	Set oAccounts = Nothing 

	ShowAccounts = iRows 

End Function   


'--------------------------------------------------------------------------------------------------
' Function GetMaxAccountId()
'--------------------------------------------------------------------------------------------------
Function GetMaxAccountId()
	Dim sSql, oAccounts

	sSql = "Select MAX(accountid) as MaxId from egov_accounts where orgid = " & session("orgid") 

	Set oAccounts = Server.CreateObject("ADODB.Recordset")
	oAccounts.Open sSQL, Application("DSN"), 0, 1
	
	If Not oAccounts.EOF And Not IsNull(oAccounts("MaxId")) Then 
		GetMaxAccountId = clng(oAccounts("MaxId"))
	Else
		GetMaxAccountId = 0
	End If 

	oAccounts.close
	Set oAccounts = Nothing 
End Function 


%>
