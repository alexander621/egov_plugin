<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: mergecitizensmatchup.asp
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
Dim sSearch, iMergeFamilyId, iKeepFamilyId, bSameFamilyId, iMaxMerge, iMaxKeep, bOnePick

sLevel = "../" ' Override of value from common.asp

' Check the page availability and user access rights in one call
PageDisplayCheck "merge registered users", sLevel	' In common.asp

iMergeFamilyId = CLng(request("mergefamilyid")) ' FamilyId
iKeepFamilyId = CLng(request("keepfamilyid")) ' Familyid
iMaxMerge = CLng(request("maxmerge")) 
iMaxKeep = CLng(request("maxkeep")) 

If iMergeFamilyId = iKeepFamilyId Then
	bSameFamilyId = True 
Else
	bSameFamilyId = False 
End If 

If iMaxKeep = iMaxMerge Then
	' If there are only head of households
	bOnePick = True 
Else 
	bOnePick = False 
End If 
%>

<html>
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
		
		function completeMerge()
		{
			if (! confirm("Merging cannot be undone.\nAre you certain you want to merge these households?"))
			{
				return;
			}
			//Check the picks and if OK, then submit the form
			// See if the Head of Household is not set to "No Merge"
			if (parseInt($("keepuserid1").value) != -2)
			{
				// If there is more that just the Head of Household to merge
				if (parseInt($("maxmerge").value) > 1)
				{
					// Loop through the family, skipping the head of household
					for (var t = 2; t <= parseInt($("maxmerge").value); t++)
					{
						// If any are set to "No Merge" then we have a problem
						if (parseInt($("keepuserid" + t).value) == -2)
						{
							alert("You cannot have any family member selected for 'No Merge' when the Head of Household is not also selected for 'No Merge'. Please correct this and try the merge again.");
							$("keepuserid" + t).focus();
							return;
						}
					}
				}
			}
			// Submit the form here
			//alert("OK to submit");
			document.frmMerge.submit();
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

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Map the Merging Household Members</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p>
				<input type="button" class="button" value="<< Back" onclick="javascript:window.location='mergecitizens.asp?keepfamilyid=<%=iKeepFamilyId%>&mergefamilyid=<%=iMergeFamilyId%>';" />
			</p>

			<form name="frmMerge" action="mergecitizensmerge.asp" method="post">
				
					<table cellpadding="0" cellspacing="0" border="0" id="mergemaptable">
						<tr><th colspan="3">Merging Household <%'=iMergeFamilyId%></th><th>Merge Into <%'=iKeepFamilyId%></th></tr>
<%							iMaxMerge = ShowMergeHousehold( iMergeFamilyId, iKeepfamilyId, bSameFamilyId, bOnePick )		%>
					</table>
				

				<p>
					<input type="button" class="button" id="savebutton" value="Complete The Merge" onclick="completeMerge()" />
					<input type="hidden" id="maxmerge" name="maxmerge" value="<%=iMaxMerge%>" />
					<input type="hidden" name="keepfamilyid" value="<%=iKeepFamilyId%>" />
					<input type="hidden" name="mergefamilyid" value="<%=iMergeFamilyId%>" />
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
' Function ShowMergeHousehold( iMergeFamilyId, iKeepFamilyId, bSameFamilyId, bOnePick )
'--------------------------------------------------------------------------------------------------
Function ShowMergeHousehold( ByVal iMergeFamilyId, ByVal iKeepFamilyId, ByVal bSameFamilyId, ByVal bOnePick )
	Dim sSql, oRs, iRecCount, bFirstPick

	iRecCount = 0
	bFirstPick = True 

	sSql = "SELECT U.userid, ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname, U.headofhousehold, ISNULL(F.relationship,'') AS relationship, U.birthdate "
	sSql = sSql & " FROM egov_users U, egov_familymembers F WHERE U.familyid = " & iMergeFamilyId
 sSql = sSql & " AND U.isdeleted = 0 AND U.userid = F.userid AND F.isdeleted = 0 "
'	sSql = sSql & " AND U.isdeleted = 0 AND U.userid = F.userid "
	sSql = sSql & " ORDER BY U.headofhousehold DESC, U.userfname, U.userlname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		iRecCount = iRecCount + 1
		response.write vbcrlf & "<tr"
		If iRecCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write "><td align=""center"">" & oRs("userfname") & " " & oRs("userlname")
		response.write "<input type=""hidden"" id=""mergeuserid" & iRecCount & """ name=""mergeuserid" & iRecCount & """ value=""" & oRs("userid") & """ />"
		
		response.write "</td><td align=""center"">"
		If oRs("headofhousehold") Then
			response.write "Head of Household"
			response.write "<input type=""hidden"" id=""headofhouseholdflag" & iRecCount & """ name=""headofhouseholdflag" & iRecCount & """ value=""1"" />"
		Else
			response.write oRs("relationship")
			response.write "<input type=""hidden"" id=""headofhouseholdflag" & iRecCount & """ name=""headofhouseholdflag" & iRecCount & """ value=""0"" />"
		End If 
		response.write "</td><td align=""center"">"
		If LCase(oRs("relationship")) = "child" And Not IsNull(oRs("birthdate")) Then
			response.write FormatNumber(GetAgeOnDate( oRs("birthdate"), Now ),1) & " yrs"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td><td align=""center"">"
		ShowKeepHouseholdPicks iKeepFamilyId, iRecCount, bSameFamilyId, oRs("headofhousehold"), bFirstPick, oRs("userid")
		response.write "</td></tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	ShowMergeHousehold = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' ShowKeepHouseholdPicks( iKeepFamilyId )
'--------------------------------------------------------------------------------------------------
Sub ShowKeepHouseholdPicks( ByVal iKeepFamilyId, ByVal iRow, ByVal bSameFamilyId, ByVal bHeadOfHousehold, ByRef bFirstPick, ByVal iMergeUserId )
	Dim sSql, oRs

	sSql = "SELECT U.userid, ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname, U.headofhousehold, ISNULL(F.relationship,'') AS relationship, U.birthdate "
	sSql = sSql & " FROM egov_users U, egov_familymembers F WHERE U.familyid = " & iKeepFamilyId
	sSql = sSql & " AND U.isdeleted = 0 AND U.userid = F.userid AND F.isdeleted = 0 "
	sSql = sSql & " ORDER BY U.headofhousehold DESC, U.userfname, U.userlname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""keepuserid" & iRow & """ name=""keepuserid" & iRow & """>"
		If Not bSameFamilyId Then 
			'response.write vbcrlf & "<option value=""-2"">No Merge &mdash; Keep In Seperate Household*</option>"
			response.write vbcrlf & "<option value=""-1"">Add As A New Family Member</option>"
		'Else
			'response.write vbcrlf & "<option value=""-2"">No Merge &mdash; Keep As Is</option>"
		End If 
		'If Not bSameFamilyId Or (bSameFamilyId And Not bHeadOfHousehold) Then
			Do While Not oRs.EOF
			 
				response.write vbcrlf & "<option value="""& oRs("userid") & """"
				If bFirstPick Then 
					response.write " selected=""selected"" "
					bFirstPick = False 
				Else
					If CLng(iMergeUserId) = CLng(oRs("userid")) Then 
						response.write " selected=""selected"" "
					End If 
				End If 
				response.write ">" & oRs("userfname") & " " & oRs("userlname") & " ("
				If oRs("headofhousehold") Then
					response.write "Head of Household"
				Else
					response.write oRs("relationship")
				End If 
				If LCase(oRs("relationship")) = "child" And Not IsNull(oRs("birthdate")) Then
					response.write ", " & FormatNumber(GetAgeOnDate( oRs("birthdate"), Now ),1) & " yrs"
				End If 
				response.write ")</option>"
				oRs.MoveNext 
			Loop 
		'End If 
		
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetAgeOnDate( dBirthDate, dCompareDate )
'--------------------------------------------------------------------------------------------------
Function GetAgeOnDate( ByVal dBirthDate, ByVal dCompareDate )
	Dim iMonths, iAge

	iAge = (DateValue(dCompareDate) - DateValue(dBirthDate)) / 365.25
	GetAgeOnDate = iAge

End Function 



%>
