<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: contactusedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the changing of the contact us mobile page
'
' MODIFICATION HISTORY
' 1.0   04/14/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sContactUsText, sContainsHTML, sLoadMsg

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "contact us edit", sLevel	' In common.asp

GetContactUsValue sContactUsText, sContainsHTML 

If request("s") = "u" Then
	sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
End If 

%>

<html lang="en">
<head runat="server">
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title>E-GovLink Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="contactusstyles.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.5.min.js"></script>

	<script language="Javascript">
	<!--
		
		function validate() 
		{
			// validate stuff and then fire off the post
			document.contactForm.submit();
		}

		function checkMsgLength() 
		{

			var displayLength = $("#contactusdisplay").val().length;

			$("#charactercount").html( displayLength );
		}

		function addHTMLTag( sTag ) 
		{
			var contactusdisplay = $("#contactusdisplay").val();

			switch ( sTag )
			{
			case "BOLD":
				contactusdisplay = contactusdisplay + " <strong></strong>";
				break;
			case "ITALICS":
				contactusdisplay = contactusdisplay + " <em></em>";
				break;
			case "H1": 
				contactusdisplay = contactusdisplay + " <h1></h1>";
				break;
			
			case "H2": 
				contactusdisplay = contactusdisplay + " <h2></h2>";
				break;

			case "H3":
				contactusdisplay = contactusdisplay + " <h3></h3>";
				break;

			case "LINK":
				contactusdisplay = contactusdisplay + " <a href=\"url goes here\"></a>";
				break;

			case "IMG":
				contactusdisplay = contactusdisplay + " <img src=\"image filename goes here\" width=\"0\" height=\"0\" />";
				break;

			case "FONT":
				contactusdisplay = contactusdisplay + " <font style=\"font-size: 10pt;\"></font>";
				break;

			case  "BR":
				contactusdisplay = contactusdisplay + "<br />";
				break;

			case "P":
				contactusdisplay = contactusdisplay + "<p>text goes here</p>";
				break;

			case "P_LEFT":
				contactusdisplay = contactusdisplay + " <p align=\"left\"></p>";
				break;

			case "P_CENTER":
				contactusdisplay = contactusdisplay + " <p align=\"center\"></p>";
				break;

			case "P_RIGHT":
				contactusdisplay = contactusdisplay + " <p align=\"right\"></p>";
				break;

			}
	
			$("#contactusdisplay").val( contactusdisplay );
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

		function SetUpPage()
		{
			<%=sLoadMsg%>
			checkMsgLength();
		}

		$(document).ready(function() {
			$("#save1").click(function() { validate(); });
			$("#save2").click(function() { validate(); });
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

		<h3>Mobile Contact Us Display</h3>

		<form name="contactForm" method="post" action="contactuseditsave.asp">
			<div id="topbtnsholder">
				<span id="screenMsg"></span>
				<input type="button" class="button" id="save1" name="save1" value="Save Changes" />
				<input type="button" value="Add a Link" class="button" onClick="doPicker('contactForm.contactusdisplay','Y','Y','Y','Y');" />
			</div>

			<div id="htmlonlycheck">
				<input type="checkbox" name="containsHTML" id="containsHTML" <%=sContainsHTML%> />&nbsp;The display text contains HTML
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

			<textarea name="contactusdisplay" id="contactusdisplay" onkeyup="checkMsgLength();"><%=sContactUsText%></textarea>
            <div id="contactusdisplaycount">Total Character Count: [<span id="charactercount"></span>]</div>

			<div id="bottombtnsholder">
				<input type="button" class="button" id="save2" name="save2" value="Save Changes" />
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
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void GetContactUsValue( sContactUsText, sContainsHTML )
'------------------------------------------------------------------------------
Sub GetContactUsValue( ByRef sContactUsText, ByRef sContainsHTML )
	Dim sSql, oRs

	sSql = "SELECT containshtml, ISNULL(contactusdisplay,'') AS contactusdisplay "
	sSql = sSql & "FROM egov_mobilecontactus WHERE orgid = " & SESSION("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		'response.write oRs("contactusdisplay") & "<br /><br />"
		sContactUsText = oRs("contactusdisplay")
		If oRs("containshtml") Then 
			sContainsHTML = " checked=""checked"" "
		Else
			sContainsHTML = ""
		End If 
	Else
		sContactUsText = ""
		sContainsHTML = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub  


'------------------------------------------------------------------------------
' void displayButton sBtnName
'------------------------------------------------------------------------------
Sub displayButton( ByVal sBtnName )

	If sBtnName <> "" Then 
		response.write "<input type=""button"" name=""" & sBtnName & "Button"" id=""" & sBtnName & "Button"" value=""" & sBtnName & """ class=""button"" onclick=""addHTMLTag('" & UCase(Trim(sBtnName)) & "');"" />"
	End If 

End Sub 



%>
