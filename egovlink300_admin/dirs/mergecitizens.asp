<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: mergecitizens.asp
' AUTHOR: Steve Loar
' CREATED: 12/23/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Merge Citizen records together. Merges the entire family.
'
' MODIFICATION HISTORY
' 1.0   12/23/2008	Steve Loar - INITIAL VERSION Started
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iKeepFamilyId, iMergeFamilyId, bHasClasses

sLevel = "../" ' Override of value from common.asp

' Check the page availability and user access rights in one call
PageDisplayCheck "merge registered users", sLevel	' In common.asp

If request("mergefamilyid") <> "" Then 
	iMergeFamilyId = CLng(request("mergefamilyid")) ' Familyid
Else
	iMergeFamilyId = 0
End If 
If request("keepfamilyid") <> "" Then 
	iKeepFamilyId = CLng(request("keepfamilyid")) ' Familyid
Else
	iKeepFamilyId = 0
End If 

bHasClasses = OrgHasFeature( "activities" )

%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="mergecitizens.css" />

	<script src="../scripts/ajaxLib.js"></script>
	<script src="../prototype/prototype-1.6.0.2.js"></script>
	<script src="../scriptaculous/src/scriptaculous.js"></script>

	<script>
	<!--

		function UserPick( sType, iKeepInfo )
		{
			if ($(sType + "familyid").value != 0)
			{
				//alert($(sType + "userid").value);
				// fire off ajax job to get details
				doAjax('getcitizendetails.asp', 'familyid=' + $(sType + "familyid").value + '&sType=' + sType, 'ShowAddress', 'get', '0');
			}
			else
			{
				$(sType + "info").innerHTML = '';
				$(sType + "info").style.visibility = 'hidden';
			}
			if (iKeepInfo == 0)
			{
				$(sType + "results").value = '';
				$(sType + "searchresults").innerHTML = '';
				$(sType + "searchstart").value = -1;
				$(sType + "searchname").value = '';
			}
		}

		function ShowAddress( sReturnJSON )
		{
			var json = sReturnJSON.evalJSON(true); 
			//alert( json.flag );
			//alert( json.type );
			if (json.flag == 'success')
			{
				$(json.type + "info").innerHTML = json.familyid + json.userid + json.useremail + json.useraddress + json.userhomephone + json.family + json.mergecount;
				$(json.type + "info").style.visibility = 'visible';
			}
			else
			{
				$(json.type + "info").innerHTML = 'No Information Found.';
				$(json.type + "info").style.visibility = 'visible';
			}
		}

		function SearchCitizens( sType, iSearchStart )
		{
			var optiontext;
			var optionchanged;

			var searchtext = $(sType + "searchname").value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < $(sType + "familyid").length ; x++)
			{
				optiontext = $(sType + "familyid").options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					$(sType + "familyid").selectedIndex = x;
					$(sType + "results").value = 'Possible Match Found.';
					$(sType + "searchresults").innerHTML = 'Possible Match Found.';
					$(sType + "searchstart").value = x;
					UserPick( sType, 1 );
					return;
				}
			}
			$(sType + "familyid").selectedIndex = 0;
			$(sType + "results").value = 'No Match Found.';
			$(sType + "searchresults").innerHTML = 'No Match Found.';
			$(sType + "searchstart").value = -1;
			$(sType + "info").innerHTML = '';
			$(sType + "info").style.visibility = 'hidden';
		}

		function ClearSearch( sType )
		{
			$(sType + "searchstart").value = -1;
		}

		function UserPicker( sType )
		{
			$(sType + "searchname").value = '';
			$(sType + "results").value = '';
			$(sType + "searchresults").innerHTML = '';
			$(sType + "searchstart").value = -1;
		}

		function continueMerge( )
		{
			if ($("keepfamilyid").selectedIndex == 0 || $("mergefamilyid").selectedIndex == 0)
			{
				alert("Please select both a citizen user to keep and a citizen user to merge.");
				return;
			}
			else
			{
<%				if not bHasClasses then		%>
					// Only Heads of households so goto merge 
					document.frmMerge.action = "mergecitizensmerge.asp";
					if (! confirm("Merging cannot be undone.\nAre you certain you want to merge these Citizens?"))
					{
						return;
					}
<%				end if		%>
				document.frmMerge.submit();
			}
		}

		function setPicks()
		{
			// Show the info for the keep household
			if ($("mergefamilyid").value != 0)
			{
				UserPick( 'merge', 0 );
			}
			// Show the info for the merge household
			if ($("keepfamilyid").value != 0)
			{
				window.setTimeout('UserPick( "keep", 0 )',400);
			}
		}

	//-->
	</script>

</head>
<body onload="setPicks();">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Merge Citizen Users</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmMerge" action="mergecitizensmatchup.asp" method="post">
				<p>
				<span class="mergepicktitle">Select the Household that will be Merged</span><br />
				<% ShowRegisteredUserPicks "merge", iMergeFamilyId %> &nbsp; &nbsp;
				<!-- <input type="button" class="button" value="Details" /> &nbsp; &nbsp; 
				<input type="button" class="button" value="Family" />--> <br />
				<span class="searcharea">
					<input type="text" id="mergesearchname" name="mergesearchname" value="" size="50" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchCitizens('merge', document.frmMerge.mergesearchstart.value);return false;}" /> &nbsp;&nbsp; 
					<input type="button" class="button" value="Search" onclick="SearchCitizens('merge', document.frmMerge.mergesearchstart.value);" /> &nbsp; &nbsp; 

					<input type="hidden" id="mergesearchstart" name="mergesearchstart" value="-1" />
					<input type="hidden" id="mergeresults" name="mergeresults" value="" />
					<span id="mergesearchresults" class="searchresults"></span>
				</span>
				<div id="mergeinfo">
				</div>
			</p>

			<p>
				<span class="mergepicktitle">Select the Household to Merge Into</span><br />
				<% ShowRegisteredUserPicks "keep", iKeepFamilyId %> &nbsp; &nbsp;
				<!-- <input type="button" class="button" value="Details" /> &nbsp; &nbsp; 
				<input type="button" class="button" value="Family" /> --> <br />
				<span class="searcharea">
					<input type="text" id="keepsearchname" name="keepsearchname" value="" size="50" maxlength="50" onkeypress="if(event.keyCode=='13'){SearchCitizens('keep', document.frmMerge.keepsearchstart.value);return false;}" /> &nbsp;&nbsp; 
					<input type="button" class="button" value="Search" onclick="SearchCitizens('keep', document.frmMerge.keepsearchstart.value);" /> &nbsp; &nbsp; 

					<input type="hidden" id="keepsearchstart" name="keepsearchstart" value="-1" />
					<input type="hidden" id="keepresults" name="keepresults" value="" />
					<span id="keepsearchresults" class="searchresults"></span>
				</span>
				<div id="keepinfo">
				</div>
			</p>
			<p>
				<input type="button" class="button" id="savebutton" value="Continue Merging" onclick="continueMerge()" />
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
'  ShowRegisteredUserPicks sType, iFamilyid 
'--------------------------------------------------------------------------------------------------
Sub ShowRegisteredUserPicks( ByVal sType, ByVal iFamilyid )
	Dim sSql, oRs

	sSql = "SELECT userid, userfname, userlname, familyid FROM egov_users "
	sSql = sSql & " WHERE headofhousehold = 1 AND orgid = " & session("orgid")
	sSql = sSql & " AND userlname IS NOT NULL AND familyid IS NOT NULL AND userlname <> '' AND isdeleted = 0"
	sSql = sSql & " ORDER BY userlname, userfname" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""" & sType & "familyid"" name=""" & sType & "familyid"" onchange=""javascript:UserPick('" & sType & "', 0);"">"
		If sType = "keep" Then
			response.write vbcrlf & "<option value=""0"">Select the registered user to keep</option>"
		Else 
			response.write vbcrlf & "<option value=""0"">Select the registered user to merge</option>"
		End If 
		
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("familyid") & """"
			If CLng(iFamilyid) = CLng(oRs("familyid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("userlname") & ", " & oRs("userfname") & " (" & oRs("familyid") & ")</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>