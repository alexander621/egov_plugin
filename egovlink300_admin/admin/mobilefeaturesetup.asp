<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mobilefeaturesetup.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the setup of mobile features for clients
'
' MODIFICATION HISTORY
' 1.0   04/14/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iFeatureCount

sLevel = "../"  'Override of value from common.asp
iFeatureCount = CLng(0)

If Not UserIsRootAdmin( session("UserID") ) Then 
  	response.redirect "../default.asp"
End If 

If request("orgid") <> "" Then 
  	iOrgId = request("orgid")
Else 
   If session("orgid") <> "" Then 
      iOrgID = session("orgid")
   Else 
     	iOrgId = GetMaxOrgId
   End If 
End If 

If request("s") <> "" Then
	If request("s") = "upd" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If
End If 

%>

<html lang="en">
<head runat="server">
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title>E-GovLink Administration Console {Mobile Feature Setup}</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.5.min.js"></script>

	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>


	<script language="Javascript">
	<!--

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
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
		}

		function goBack() 
		{
			window.location='featureselection.asp?orgid=<%=iOrgId%>';
		}

		function validate()
		{
			var maxFeatureCount = parseInt($("#maxfeaturecount").val())
			var rege = /^\d*$/;
			var showid;

			if ( maxFeatureCount > parseInt('0') )
			{
				for( x = 1; x <= maxFeatureCount; x++)
				{
					if ( $("#mobiledisplayorder" + x).val() != '' )
					{
					
						// Remove any extra spaces
						$("#mobiledisplayorder" + x).val( removeSpaces($("#mobiledisplayorder" + x).val()) );
						//Remove commas that would cause problems in validation
						$("#mobiledisplayorder" + x).val( removeCommas($("#mobiledisplayorder" + x).val()) );

						Ok = rege.test( $("#mobiledisplayorder" + x).val() );
						if ( ! Ok )
						{
							showid = "mobiledisplayorder" + x;
							inlineMsg( document.getElementById(showid).id,'<strong>Invalid: </strong>The display order must be a positive integer.',10,showid );
							return false;
						}
					}

					if ( $("#mobileitemcount" + x).val() != '' )
					{
					
						// Remove any extra spaces
						$("#mobileitemcount" + x).val( removeSpaces($("#mobileitemcount" + x).val()) );
						//Remove commas that would cause problems in validation
						$("#mobileitemcount" + x).val( removeCommas($("#mobileitemcount" + x).val()) );

						Ok = rege.test( $("#mobileitemcount" + x).val() );
						if ( ! Ok )
						{
							showid = "mobileitemcount" + x;
							inlineMsg( document.getElementById(showid).id,'<strong>Invalid: </strong>The items shown must be a positive integer.',10,showid );
							return false;
						}
					}

					if ( $("#mobilelistcount" + x).val() != '' )
					{
					
						// Remove any extra spaces
						$("#mobilelistcount" + x).val( removeSpaces($("#mobilelistcount" + x).val()) );
						//Remove commas that would cause problems in validation
						$("#mobilelistcount" + x).val( removeCommas($("#mobilelistcount" + x).val()) );

						Ok = rege.test( $("#mobilelistcount" + x).val() );
						if ( ! Ok )
						{
							showid = "mobilelistcount" + x;
							inlineMsg( document.getElementById(showid).id,'<strong>Invalid: </strong>The items shown must be a positive integer.',10,showid );
							return false;
						}
					}
				}
			}
			//alert( 'Ok');
			document.FeatureForm.submit();
		}

		$(document).ready(function() {
			$('#orgid').change(function() { document.pickForm.submit(); });
			$("#backbtn").click(function() { goBack(); });
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

		<h3>Mobile Feature Setup</h3>

		
		
		<form name="pickForm" method="post" action="mobilefeaturesetup.asp">
			<p>
				<span id="screenMsg"></span>
				Organization: <% ShowOrgDropDown iOrgId %>  
			</p>
		</form>

		<form name="FeatureForm" method="post" action="mobilefeaturesetupsave.asp">
			<input type="hidden" id="orgid" name="orgid" value="<%=iOrgId%>" />
			<div id="topbtnsholder">
				<input type="button" class="button" id="backbtn" value="<< Back" /> &nbsp;
				<input type="button" class="button" id="save1" name="save1" value="Save Changes" />
			</div>

			<div id="mobilefeatureholder">
				<table id="mobilefeaturesetup" border="1" cellpadding="3" cellspacing="0">
					<tr>
						<th>Activated</th><th>Mobile Name</th><th>Display Order</th><th>Has Items On Main Page</th><th>Main Page<br />Items shown</th><th>List Page<br />Items Shown</th><th>Id</th><th>Feature</th><th width="250">Feature Notes</th>
					</tr>
					<% iFeatureCount = ShowMobileFeatures( iOrgId ) %>
				</table>
				<input type="hidden" id="maxfeaturecount" name="maxfeaturecount" value="<%=iFeatureCount%>" />
			</div>

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
' void ShowOrgDropDown iOrgId 
'------------------------------------------------------------------------------
Sub  ShowOrgDropDown( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT orgname, orgcity, orgid, defaultstate FROM organizations "
	sSql = sSql & "WHERE isdeactivated = 0 ORDER BY orgcity, defaultstate"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""orgid"" id=""orgid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("orgid") & """ "
			If CLng(iOrgId) = CLng(oRs("orgid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("orgcity") & ", " & oRs("defaultstate") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' integer ShowMobileFeatures iOrgId
'------------------------------------------------------------------------------
Function ShowMobileFeatures( ByVal iOrgId )
	Dim sSql, oRs, iRows, sMobiledisplayorder, sMobileitemcount, sMobilelistcount

	iRows = CLng(0)

	sSql = "SELECT O.featureid, ISNULL(O.mobilename, F.featurename) AS mobilename, F.feature, mobileurl, "
	sSql = sSql & "ISNULL(O.mobileitemcount, 9999) AS mobileitemcount, ISNULL(O.mobiledisplayorder, F.mobiledefaultdisplayorder) AS displayorder, "
	sSql = sSql & "ISNULL(O.mobiledisplayorder, 9999) AS mobiledisplayorder, ISNULL(O.mobilelistcount, 9999) AS mobilelistcount, "
	sSql = sSql & "ismobilenavonly, mobileisactivated, F.mobiledefaultitemcount, F.mobiledefaultlistcount, "
	sSql = sSql & "F.mobiledefaultdisplayorder, F.featurename, featurenotes "
	sSql = sSql & "FROM egov_organizations_to_features O, egov_organization_features F "
	sSql = sSql & "WHERE O.featureid = F.featureid AND O.orgid = " & iOrgId
	sSql = sSql & " AND F.hasmobileview = 1 AND (F.feature_offline = 'N' OR F.feature_offline IS NULL) "
	sSql = sSql & "ORDER BY 6 ASC, 2"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iRows = iRows + CLng(1)
		response.write vbcrlf & "<tr"
		If iRows Mod 2 = 0 Then 
			response.write " class=""altrow"" "
		End If 
		response.write ">"

		' activated
		response.write "<td align=""center"" valign=""top"">"
		If oRs("mobileisactivated") Then
			sChecked = " checked=""checked"" "
		Else
			sChecked = ""
		End If 
		response.write "<input type=""checkbox"" name=""mobileisactivated" & iRows & """" & sChecked & " />" 
		response.write "<input type=""hidden"" name=""featureid" & iRows & """ value=""" & oRs("featureid") & """ />"
		response.write "</td>"

		' mobile name
		response.write "<td align=""center"" valign=""top"" nowrap=""nowrap"">"
		response.write "<input type=""input"" name=""mobilename" & iRows & """ value=""" & oRs("mobilename") & """ size=""15"" maxlength=""50"" />"
		response.write "<br />[" & oRs("featurename") & "]"
		response.write "</td>"

		' display order
		response.write "<td align=""center"" valign=""top"">"
		If CLng(oRs("mobiledisplayorder")) < CLng(9999) Then
			sMobiledisplayorder = oRs("mobiledisplayorder")
		Else
			sMobiledisplayorder = ""
		End If 
		response.write "<input type=""input"" id=""mobiledisplayorder" & iRows & """ name=""mobiledisplayorder" & iRows & """ value=""" & sMobiledisplayorder & """ size=""3"" maxlength=""3"" />"
		response.write "<br />[" & oRs("mobiledefaultdisplayorder") & "]"
		response.write "</td>"

		' is mobile nav only
		response.write "<td align=""center"" valign=""top"">"
		If oRs("ismobilenavonly") Then 
			response.write "no"
		Else
			response.write "yes"
		End If 
		response.write "</td>"

		' main page items shown
		response.write "<td align=""center"" valign=""top"">"
		If CLng(oRs("mobileitemcount")) < CLng(9999) Then
			sMobileitemcount = oRs("mobileitemcount")
		Else
			sMobileitemcount = ""
		End If 
		response.write "<input type=""input"" id=""mobileitemcount" & iRows & """ name=""mobileitemcount" & iRows & """ value=""" & sMobileitemcount & """ size=""3"" maxlength=""3"" />"
		response.write "<br />[" & oRs("mobiledefaultitemcount") & "]"
		response.write "</td>"

		' list page items shown
		response.write "<td align=""center"" valign=""top"">"
		If CLng(oRs("mobilelistcount")) < CLng(9999) Then
			sMobilelistcount = oRs("mobilelistcount")
		Else
			sMobilelistcount = ""
		End If 
		response.write "<input type=""input"" id=""mobilelistcount" & iRows & """ name=""mobilelistcount" & iRows & """ value=""" & sMobilelistcount & """ size=""3"" maxlength=""3"" />"
		response.write "<br />[" & oRs("mobiledefaultlistcount") & "]"
		response.write "</td>"

		' feature id
		response.write "<td align=""center"" valign=""top"">"
		response.write oRs("featureid")
		response.write "</td>"

		' feature
		response.write "<td align=""center"" valign=""top"">"
		response.write oRs("feature")
		response.write "</td>"

		' feature notes
		response.write "<td valign=""top"">"
		response.write oRs("featurenotes")
		response.write "</td>"

		response.write "</tr>"

		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	ShowMobileFeatures = iRows

End Function 



%>
