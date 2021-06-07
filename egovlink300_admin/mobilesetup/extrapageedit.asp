<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: extrapageedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/15/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the creation and editing of ad hoc mobile pages
'
' MODIFICATION HISTORY
' 1.0   04/15/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iPageId, sPageBody, sContainsHTML, sLoadMsg, sDisplayPage, sPageTitle, sBtnText
Dim sTitle

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "mobileextrapages", sLevel	' In common.asp

iPageId = CLng(request("pageid"))

GetPageValues iPageId, sPageTitle, sPageBody, sContainsHTML, sDisplayPage

If iPageId = CLng(0) Then
	' the default should be that the page is displayed
	sDisplayPage = " checked=""checked"" "
	sBtnText = "Create Page"
	sTitle = "New"
Else
	sBtnText = "Save Changes"
	sTitle = "Edit an"
End If 

If request("s") = "u" Then
	sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
End If 
If request("s") = "n" Then
	sLoadMsg = "displayScreenMsg('This Page Has Been Created');"
End If

' <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />'
%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	
	<title>E-GovLink Administration Console</title>

	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../classes/classes.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="mobilesetupstyles.css" />

	<script src="../scripts/jquery-1.9.1.min.js"></script>

	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	<!--
		
		function validate() 
		{
			var bOk = true;
			// validate stuff and then fire off the post
			if ($("#pagebody").val() == "")
			{
				$("#pagebody").focus();
				inlineMsg("pagebody",'<strong>Missing Value: </strong>Please enter some text into the page body.',8,"pagebody");
				bOk = false;
			}

			if ($("#pagetitle").val() == "")
			{
				$("#pagetitle").focus();
				inlineMsg("pagetitle",'<strong>Missing Value: </strong>Please enter some text into the page title.',8,"pagetitle");
				bOk = false;
			}

			if (bOk == false)
			{
				return false;
			}
			else
			{
				document.frmExtraPage.submit();
			}
		}

		function checkMsgLength() 
		{

			var displayLength = $("#pagebody").val().length;

			$("#charactercount").html( displayLength );
		}

		function addHTMLTag( sTag ) 
		{
			var pagebody = $("#pagebody").val();

			switch ( sTag )
			{
			case "BOLD":
				pagebody = pagebody + " <strong></strong>";
				break;
			case "ITALICS":
				pagebody = pagebody + " <em></em>";
				break;
			case "H1": 
				pagebody = pagebody + " <h1></h1>";
				break;
			
			case "H2": 
				pagebody = pagebody + " <h2></h2>";
				break;

			case "H3":
				pagebody = pagebody + " <h3></h3>";
				break;

			case "LINK":
				pagebody = pagebody + " <a href=\"url goes here\"></a>";
				break;

			case "IMG":
				pagebody = pagebody + " <img src=\"image filename goes here\" width=\"0\" height=\"0\" />";
				break;

			case "FONT":
				pagebody = pagebody + " <font style=\"font-size: 10pt;\"></font>";
				break;

			case  "BR":
				pagebody = pagebody + "<br />";
				break;

			case "P":
				pagebody = pagebody + "<p>text goes here</p>";
				break;

			case "P_LEFT":
				pagebody = pagebody + " <p align=\"left\"></p>";
				break;

			case "P_CENTER":
				pagebody = pagebody + " <p align=\"center\"></p>";
				break;

			case "P_RIGHT":
				pagebody = pagebody + " <p align=\"right\"></p>";
				break;

			}
	
			$("#pagebody").val( pagebody );
			checkMsgLength();
		}

		function doPicker( sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL ) 
		{
			w = 600;
			h = 400;
			l = (screen.AvailWidth/2)-(w/2);
			t = (screen.AvailHeight/2)-(h/2);
			lcl_showFolderStart = "";
			lcl_folderStart     = 0;

			//Determine which options will be displayed
			if((p_displayDocuments=="")||(p_displayDocuments==undefined)) 
			{
				lcl_displayDocuments = "";
			}
			else
			{
				lcl_displayDocuments = "&displayDocuments=Y";
				lcl_folderStart = lcl_folderStart + 1;
			}

			if((p_displayActionLine=="")||(p_displayActionLine==undefined)) 
			{
				lcl_displayActionLine = "";
			}
			else
			{
				lcl_displayActionLine = "&displayActionLine=Y";
				lcl_folderStart = lcl_folderStart + 1;
			}

			if((p_displayPayments=="")||(p_displayPayments==undefined)) 
			{
				lcl_displayPayments = "";
			}
			else
			{
				lcl_displayPayments = "&displayPayments=Y";
				lcl_folderStart = lcl_folderStart + 1;
			}

			if((p_displayURL=="")||(p_displayURL==undefined)) 
			{
				lcl_displayURL = "";
			}
			else
			{
				lcl_displayURL = "&displayURL=Y";
			}

			if(lcl_folderStart > 0) 
			{
				//lcl_showFolderStart = "&folderStart=published_documents";
				lcl_showFolderStart = "&folderStart=CITY_ROOT";
			}

			pickerURL  = "../picker_new/default.asp";
			pickerURL += "?name=" + sFormField;
			pickerURL += "&returnAsHTMLLink=<%=lcl_returnAsHTMLLink%>";
			pickerURL += lcl_showFolderStart;
			pickerURL += lcl_displayDocuments;
			pickerURL += lcl_displayActionLine;
			pickerURL += lcl_displayPayments;
			pickerURL += lcl_displayURL;

			eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
		}

		function insertAtCaret( textEl, text ) 
		{
			if (textEl.createTextRange && textEl.caretPos) {
				var caretPos = textEl.caretPos;
				caretPos.text =
				caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				text + ' ' : text;
			}
			else
			// Append the link to the textarea
			textEl.value = textEl.value + text;
			checkMsgLength();
		}

		function displayScreenMsg( iMsg ) 
		{
			if( iMsg != "" ) 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
		}

		function deletePage()
		{
			if (confirm("Delete this page?"))
			{
				// Off to the delete script
				//alert("Deleting");
				location.href = "extrapagedelete.asp?pageid=<%=iPageId%>";
			}
		}

		function goBack()
		{
			location.href = "extrapagelist.asp";
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
			checkMsgLength();
		}

		$(document).ready(function() {
			$("#save1").click(function() { validate(); });
			$("#save2").click(function() { validate(); });
			$("#back").click(function() { goBack(); });
		});

	//-->
	</script>

</head>
<body onload="SetUpPage();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<h3><%=sTitle%> Extra Mobile Page</h3>

		<form name="frmExtraPage" method="post" action="extrapageupdate.asp">
			<input type="hidden" name="pageid" value="<% =iPageId %>" />
			<div id="topbtnsholder">
				<span id="screenMsg"></span>
				<input type="button" class="button" id="back" name="back" value="<< Back" /> &nbsp; 
				<input type="button" class="button" id="save1" name="save1" value="<%=sBtnText%>" /> &nbsp; 
<%				If iPageId > CLng(0) Then			%>
					<input type="button" class="button" name="deletepage" value="Delete This Page" onclick="deletePage()" /> &nbsp; 
<%				End If								%>
				<input type="button" value="Add a Link" class="button" onClick="doPicker('frmExtraPage.pagebody','Y','Y','Y','Y');" />
			</div>

			<div id="htmlonlycheck">
				<strong>Page Title:</strong> <input type="text" name="pagetitle" id="pagetitle" value="<% =sPageTitle %>" size="25" maxlength="25" /><br /><br />
				<input type="checkbox" name="displaypage" id="displaypage" <%=sDisplayPage%> />&nbsp;Display this page<br />
				<!-- <input type="hidden" name="containsHTML" id="containsHTML" value="on" /> -->
				<input type="checkbox" name="containsHTML" id="containsHTML" <%=sContainsHTML%> />&nbsp;Contains HTML in your page text
			</div>

			<table id="tagtable" border="0" cellspacing="0" cellpadding="2">
				<tr>
					<td align="right" nowrap="nowrap" valign="bottom"><strong>Common HTML formatting tags:</strong></td>
					<td valign="bottom"><% displayButton "Bold" %></td>
					<td valign="bottom"><% displayButton "Italics" %></td>
					<td valign="bottom"><% displayButton "FONT" %></td>
					<td valign="bottom"><% displayButton "H1" %></td>
					<td valign="bottom"><% displayButton "H2" %></td>
					<td valign="bottom"><% displayButton "H3" %></td>
					<td valign="bottom"><% displayButton "LINK" %></td>
					<td valign="bottom"><% displayButton "IMG" %></td>
					<td valign="bottom"><% displayButton "BR" %></td>
					<td valign="bottom"><% displayButton "P" %></td>
					<td align="center" valign="bottom">Alignment:<br />
						<select name="p_format_alignment" onchange="addHTMLTag( this.value );">
							<option value=""></option>
							<option value="P_LEFT">LEFT</option>
							<option value="P_CENTER">CENTER</option>
							<option value="P_RIGHT">RIGHT</option>
						</select>
					</td>
					<td nowrap="nowrap" valign="bottom">[<a href="http://www.w3schools.com/tags/default.asp" target="_blank">Additional Tags</a>]</td>
				</tr>
			</table>

			<p>Note: Your page text will always be contained within other HTML tags. All tags MUST be closed or the page will not function properly.</p>

			<textarea name="pagebody" id="pagebody" onkeyup="checkMsgLength();"><%=sPageBody%></textarea>
            <div id="pagebodycount">Total Character Count: [<span id="charactercount"></span>]</div>

			<div id="bottombtnsholder">
				<input type="button" class="button" id="save2" name="save2" value="<%=sBtnText%>" />
			</div>
		</form>

		
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'------------------------------------------------------------------------------
' GetPageValues iPageId, sPageTitle, sPageBody, sContainsHTML, sDisplayPage
'------------------------------------------------------------------------------
Sub GetPageValues( ByVal iPageId, ByRef sPageTitle, ByRef sPageBody, ByRef sContainsHTML, ByRef sDisplayPage )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(pagetitle,'') AS pagetitle, containshtml, displaypage, ISNULL(pagebody,'') AS pagebody "
	sSql = sSql & "FROM egov_extramobilepages WHERE pageid = " & iPageId
	sSql = sSql & " AND orgid = " & SESSION("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sPageTitle = oRs("pagetitle")

		'response.write oRs("pagebody") & "<br /><br />"
		sPageBody = Replace(oRs("pagebody"),"&","&amp;")

		If oRs("containshtml") Then 
			sContainsHTML = " checked=""checked"" "
		Else
			sContainsHTML = ""
		End If 

		If oRs("displaypage") Then 
			sDisplayPage = " checked=""checked"" "
		Else
			sDisplayPage = ""
		End If 
	Else
		sPageTitle = ""
		sPageBody = ""
		sContainsHTML = ""
		sDisplayPage = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub  


'------------------------------------------------------------------------------
' displayButton sBtnName
'------------------------------------------------------------------------------
Sub displayButton( ByVal sBtnName )

	If sBtnName <> "" Then 
		response.write "<input type=""button"" name=""" & sBtnName & "Button"" id=""" & sBtnName & "Button"" value=""" & sBtnName & """ class=""button"" onclick=""addHTMLTag('" & UCase(Trim(sBtnName)) & "');"" />"
	End If 

End Sub 



%>