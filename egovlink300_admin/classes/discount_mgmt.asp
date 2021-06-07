<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DISCOUNT_MGMT.ASP
' AUTHOR: TERRY FOSTER
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/26/06   TERRY FOSTER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "discounts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	
	<script language="Javascript" src="tablesort.js"></script>

	<script language="Javascript">
	<!--

		function mouseOverRow( oRow )
		{
			oRow.style.backgroundColor='#93bee1';
			oRow.style.cursor='pointer';
			oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
			if (oNextRow)
			{
				oNextRow.style.backgroundRepeat="repeat-x";
				oNextRow.style.backgroundImage="url(../images/shadow.png)";
			}
		}

		function mouseOutRow( oRow )
		{	
			oRow.style.backgroundColor='';
			oRow.style.cursor='';
			oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
			if (oNextRow)
			{
				oNextRow.style.backgroundImage="";
			}
		}

		function deleteconfirm(ID, sName) 
		{
			if(confirm('Do you wish to delete \'' + sName + '\'?')) {
				window.location="discount_delete.asp?iPriceDiscountid=" + ID;
			}
		}

	//-->
	</script>

</head>

<body>

 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: Discount Management</strong></font><br />
	<!--<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS LIST-->
	<% ListDiscounts %> 
<!--END: CLASS LIST-->

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
' SUB LISTDISCOUNTS()
'--------------------------------------------------------------------------------------------------
Sub ListDiscounts()
	Dim sSql, iRowCount, sClass

	iRowCount = 0
	' GET ALL DISCOUNTS FOR ORG
	sSQL = "SELECT D.*, T.discounttype FROM egov_price_discount D, egov_price_discount_types T "
	sSql = sSql & " WHERE D.discounttypeid = T.discounttypeid and D.orgid = " & SESSION("orgid") & ""
	'response.write sSQL

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	' DRAW LINK TO NEW DISCOUNT
	response.write "<div id=""functionlinks""><a href=""discount_edit.asp?discountid=0""><img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;New Discount</a></div>"

	If NOT oList.EOF Then

		' DRAW TABLE 
		response.write "<div class=""shadow""><table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin style-alternate sortable-onload-2"" width=""100%"">"
		
		' HEADER ROW
		response.write "<tr>"
		response.write "<th>Name</th><th>Type</th><th>Qty. Required</th><th>Amount</th><th>Description</th><th>Delete</th>"
		response.write "</tr>"
		
		' LOOP THRU AND DISPLAY ROWS
		Do While Not oList.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sClass = " class=""altrow"" "
			Else
				sClass = ""
			End If 

			response.write "<tr " & sClass & " id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			'response.write "<td nowrap><a href=discount_edit.asp?PriceDiscountid=" & oList("PriceDiscountid") & ">Edit</a> | "
			response.write "<td title=""click to edit"" onClick=""location.href='discount_edit.asp?PriceDiscountid=" & oList("PriceDiscountid") & "';"">&nbsp;" & oList("DiscountName") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='discount_edit.asp?PriceDiscountid=" & oList("PriceDiscountid") & "';"">" & oList("discounttype") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='discount_edit.asp?PriceDiscountid=" & oList("PriceDiscountid") & "';"">" & oList("qtyrequired") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='discount_edit.asp?PriceDiscountid=" & oList("PriceDiscountid") & "';"">" & FormatCurrency(oList("discountamount"),2) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='discount_edit.asp?PriceDiscountid=" & oList("PriceDiscountid") & "';"">" & oList("discountdescription") & "</td>"
			response.write "<td><a  title=""click to delete"" href=""javascript:deleteconfirm(" & oList("PriceDiscountid") & ", '" & FormatForJavaScript(oList("DiscountName")) & "')"">Delete</a></td>"
			response.write "</tr>"
			oList.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write "</table></div>"
		oList.close
		Set oList = Nothing 
	
	Else
		' NO categoryS WERE FOUND
		response.write "<font color=""red""><b>There are no discounts to display.</b></font>"
	
	End If

End Sub
%>


