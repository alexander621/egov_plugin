<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: merchandiseedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the creating and editing of merchandise
'
' MODIFICATION HISTORY
' 1.0  04/24/2009 Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sTitle, iMerchandiseId, sMerchandise, sPrice, bShowPublic, sShowPublic

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "merchandise setup", sLevel	' In common.asp

iMerchandiseId = CLng(request("merchandiseid") )

If CLng(iMerchandiseId) > CLng(0) Then
	sTitle = "Edit"
	GetMerchandiseInfo iMerchandiseId, sMerchandise, sPrice, bShowPublic
Else
	sTitle = "New"
	sMerchandise = ""
	sPrice = "0.00"
	bShowPublic = False 
End If 

If bShowPublic Then
	sShowPublic = " checked=""checked"" "
Else
	sShowPublic = ""
End If 

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="merchandise.css" />
	
	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function Delete() 
		{
			if (confirm("Do you wish to delete this merchandise item?"))
			{
				location.href='merchandisedelete.aspx?merchandiseid=<%=iMerchandiseId%>&orgid=<%=Session("orgid")%>';
			}
		}

		function NewOfferingRow()
		{
			document.frmMerchandise.maxofferings.value = parseInt(document.frmMerchandise.maxofferings.value) + 1;
			var tbl = document.getElementById("offeringsedit");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmMerchandise.maxofferings.value);
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.align = 'center';
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'remove' + newRow;
			e.id = 'remove' + newRow;
			cellZero.appendChild(e);
			//Add the hidden field
			e = document.createElement('input');
			e.type = 'hidden';
			e.name = 'merchandisecatalogid' + newRow;
			e.id = 'merchandisecatalogid' + newRow;
			e.value = '0';
			cellZero.appendChild(e);

			//color pick here
			cellZero = row.insertCell(1);
			cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'merchandisecolorid' + newRow;
			e0.id = 'merchandisecolorid' + newRow;
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmMerchandise.maxofferings.value); t++ )
			{
				if (document.getElementById("merchandisecolorid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("merchandisecolorid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("merchandisecolorid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("merchandisecolorid" + t).options[s].value );
				e0.appendChild(op);
			}

			//size pick here
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'merchandisesizeid' + newRow;
			e1.id = 'merchandisesizeid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("merchandisesizeid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("merchandisesizeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("merchandisesizeid" + t).options[s].value );
				e1.appendChild(op);
			}

			// In Stock checkbox
			var cellTwo = row.insertCell(3);
			cellTwo.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'checkbox';
			e1.name = 'instock' + newRow;
			e1.id = 'instock' + newRow;
			cellTwo.appendChild(e1);

			// Show Public checkbox
			var cellTwo = row.insertCell(4);
			cellTwo.align = 'center';
			e1 = document.createElement('input');
			e1.type = 'checkbox';
			e1.name = 'showpublic' + newRow;
			e1.id = 'showpublic' + newRow;
			cellTwo.appendChild(e1);
		}

		function RemoveOfferingRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("offeringsedit");
			// Check the Inspection rows for any selected for removal
			for (var t = 0; t <= parseInt(document.frmMerchandise.maxofferings.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("remove" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("remove" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("remove" + t).checked = false;
							document.getElementById("merchandisecolorid" + t).options[0].selected = true;
							document.getElementById("merchandisesizeid" + t).options[0].selected = true;
							document.getElementById("instock" + t).checked = false;
							document.getElementById("showpublic" + t).checked = false;
						}
					}
				}
			}
		}

		function Validate()
		{
			// Check for a merchandise
			if (document.frmMerchandise.merchandise.value == '')
			{
				//alert("Please provide a merchandise name, then try saving again.");
				inlineMsg($("merchandise").id,'<strong>Invalid Value: </strong>Please provide a merchandise name, then try saving again.',10,$("merchandise").id);
				document.frmMerchandise.merchandise.focus();
				return;
			}


			// Check for a price
			if ($("price").value != '')
			{
				// Remove any extra spaces
				document.frmMerchandise.price.value = removeSpaces(document.frmMerchandise.price.value);
				//Remove commas that would cause problems in validation
				document.frmMerchandise.price.value = removeCommas(document.frmMerchandise.price.value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.getElementById("price").value);
				if ( ! Ok )
				{
					inlineMsg($("price").id,'<strong>Invalid Value: </strong>The price must be a number in currency format. Please correct this and try saving again.',10,$("price").id);
					document.getElementById("price").focus();
					return;
				}
				else
				{
					document.getElementById("price").value = format_number(Number(document.getElementById("price").value),2);
					if (Number(document.getElementById("price").value) > Number(999.99))
					{
						document.getElementById("price").value = format_number(0,2);
						inlineMsg($("price").id,'<strong>Invalid Value: </strong>The price must be less than $1000. Please correct this and try saving again.',10,$("price").id);
						document.getElementById("price").focus();
						return;
					}

					if (Number(document.getElementById("price").value) == Number(0))
					{
						document.getElementById("price").value = format_number(0,2);
						inlineMsg($("price").id,'<strong>Invalid Value: </strong>The price must be more than $0. Please correct this and try saving again.',10,$("price").id);
						document.getElementById("price").focus();
						return;
					}
				}
			}
			else 
			{
				inlineMsg($("price").id,'<strong>Invalid Value: </strong>Please provide a price, then try saving again.',10,$("price").id);
				document.frmMerchandise.price.focus();
				return;
			}

			//alert("OK to submit");
			document.frmMerchandise.submit();
		}

<%		If request("success") <> "" Then 
			DisplayMessagePopUp 
		End If 
%>

	//-->
	</script>

</head>

<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%> Merchandise</strong></font><br /><br />
				<a href="merchandiselist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
			</p>
			<!--END: PAGE TITLE-->

<%		If CLng(iMerchandiseId) = CLng(0) Then %>
			<input type="button" id="savebutton" class="button" onclick="javascript:Validate();" value="Create" /><br />
<%		Else %>
			<input type="button" id="savebutton" class="button" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
			<input type="button" class="button" onclick="javascript:Delete();" value="Delete" /> &nbsp; &nbsp;
			<br />
<%		End If %>

		<form name="frmMerchandise" action="merchandiseupdate.asp" method="post">
		<input type="hidden" name="merchandiseid" value="<%=iMerchandiseId%>" />
		
		<p>
			Merchandise: &nbsp;&nbsp; <input type="text" id="merchandise" name="merchandise" value="<%=sMerchandise%>" size="50" maxlength="50" />
		</p>
		<p>
			Price: &nbsp;&nbsp; <input type="text" id="price" name="price" value="<%=sPrice%>" size="6" maxlength="6" />
		</p>
		<p>
			<input type="checkbox" name="showpublic" <%=sShowPublic%> /> &nbsp; Display on purchase pages (Admin and Public)
		</p>

		<p>
			Use this section below to set up the merchandise offerings. Add rows for each combination of color 
			and size that will be offered. To remove an offering check the &quot;Select For Removal&quot; box 
			at the start of the row, click the &quot;Remove Selected Rows&quot; button then save your changes.
		</p>
		
		<p>
			<input type="button" class="button" value="Add A Row" onclick="NewOfferingRow()" /> &nbsp; &nbsp;
			<input type="button" class="button" value="Remove Selected Rows" onclick="RemoveOfferingRows()" />
		</p>

		<div class="shadow" id="offeringseditshadow">
		<table id="offeringsedit" cellpadding="0" cellspacing="0" border="0">
			<tr><th>Select For<br />Removal</th><th>Color</th><th>Size</th><th>In Stock</th><th>Show</th></tr>
<%			iMaxOfferings = ShowOfferings( iMerchandiseId )		%>
		</table>
		</div>
		<input type="hidden" name="maxofferings" id="maxofferings" value="<%=iMaxOfferings%>" />

		</form>
		<!--END: EDIT FORM-->

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

<%	If request("success") <> "" Then 
		SetupMessagePopUp request("success")
	End If	
%>

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetMerchandiseInfo( iMerchandiseId, sMerchandise, sPrice, bShowPublic )
'--------------------------------------------------------------------------------------------------
Sub GetMerchandiseInfo( ByVal iMerchandiseId, ByRef sMerchandise, ByRef sPrice, ByRef bShowPublic )
	Dim sSql, oRs

	sSql = "SELECT merchandise, ISNULL(description,'') AS description, price, showpublic "
	sSql = sSql & " FROM egov_merchandise WHERE orgid = " & session("orgid") & " AND merchandiseid = " & iMerchandiseId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sMerchandise = oRs("merchandise")
		sPrice = FormatNumber(oRs("price"),2,,,0)
		If oRs("showpublic") Then 
			bShowPublic = True 
		Else
			bShowPublic = False 
		End If 
	Else
		sMerchandise = ""
		sPrice = "0.00"
		bShowPublic = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Sub


'--------------------------------------------------------------------------------------------------
' Function GetMerchandiseInfo( iMerchandiseId, sMerchandise, sPrice, bShowPublic )
'--------------------------------------------------------------------------------------------------
Function ShowOfferings( iMerchandiseId )
	Dim sSql, oRs, sInStock, sShowPublic, iRowCount

	iRowCount = CLng(0) 

	sSql = "SELECT merchandisecatalogid, ISNULL(merchandisecolorid,0) AS merchandisecolorid, ISNULL(merchandisesizeid,0) AS merchandisesizeid, instock, showpublic "
	sSql = sSql & " FROM egov_merchandisecatalog WHERE orgid = " & session("orgid") & " AND merchandiseid = " & iMerchandiseId
	sSql = sSql & " ORDER BY merchandisecolor, displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then 
				sRowClass = " class=""altrow"" "
			Else
				sRowClass = ""
			End If 
			If oRs("instock") Then
				sInStock = " checked=""checked"" "
			Else
				sInStock = ""
			End If 
			If oRs("showpublic") Then
				sShowPublic = " checked=""checked"" "
			Else
				sShowPublic = ""
			End If 
			response.write vbcrlf & "<tr" & sRowClass & "><td align=""center"">"
			response.write "<input type=""checkbox"" name=""remove" & iRowCount & """ id=""remove" & iRowCount & """ />"
			response.write "<input type=""hidden"" id=""merchandisecatalogid" & iRowCount & """ name=""merchandisecatalogid" & iRowCount & """ value=""" & oRs("merchandisecatalogid") & """ />"
			response.write "</td><td align=""center"">"
			ShowMerchandiseColors oRs("merchandisecolorid"), iRowCount
			response.write "</td><td align=""center"">"
			ShowMerchandiseSizes oRs("merchandisesizeid"), iRowCount
			response.write "</td><td align=""center"">"
			response.write "<input type=""checkbox"" name=""instock" & iRowCount & """ " & sInStock & " />"
			response.write "</td><td align=""center"">"
			response.write "<input type=""checkbox"" name=""showpublic" & iRowCount & """ " & sShowPublic & " />"
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
	Else
		' output a starter row
		response.write vbcrlf & "<tr><td align=""center"">"
		response.write "<input type=""checkbox"" name=""remove1"" id=""remove1"" value=""0"" />"
		response.write "<input type=""hidden"" id=""merchandisecatalogid1"" name=""merchandisecatalogid1"" value=""0"" />"
		response.write "</td><td align=""center"">"
		ShowMerchandiseColors 0, 1
		response.write "</td><td align=""center"">"
		ShowMerchandiseSizes 0, 1
		response.write "</td><td align=""center"">"
		response.write "<input type=""checkbox"" id=""instock1"" name=""instock1"" " & sInStock & " />"
		response.write "</td><td align=""center"">"
		response.write "<input type=""checkbox"" id=""showpublic1"" name=""showpublic1"" " & sShowPublic & " />"
		response.write "</td></tr>"
		iRowCount = CLng(1) 
	End If
	
	oRs.Close
	Set oRs = Nothing 

	ShowOfferings = iRowCount
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowMerchandiseColors( iMerchandiseColorId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowMerchandiseColors( iMerchandiseColorId, iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT merchandisecolorid, merchandisecolor FROM egov_merchandisecolors "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY isnocolor DESC, merchandisecolor"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""merchandisecolorid" & iRowCount & """ name=""merchandisecolorid" & iRowCount & """>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("merchandisecolorid") & """ "  
			If CLng(iMerchandiseColorId) = CLng(oRs("merchandisecolorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("merchandisecolor")
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' Sub ShowMerchandiseSizes( iMerchandiseSizeId, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowMerchandiseSizes( iMerchandiseSizeId, iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT merchandisesizeid, merchandisesize FROM egov_merchandisesizes "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""merchandisesizeid" & iRowCount & """ name=""merchandisesizeid" & iRowCount & """>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("merchandisesizeid") & """ "  
			If CLng(iMerchandiseSizeId) = CLng(oRs("merchandisesizeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("merchandisesize")
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing

End Sub  



%>