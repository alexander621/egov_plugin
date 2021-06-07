<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">


<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->


<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_ISSUE_LOCATION.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/5/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0	02/05/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' INITIALIZE AND DECLARE VARIABLES
Dim sError
sLevel = "../../" ' OVERRIDE OF VALUE FROM COMMON.ASP

' SET TIMEZONE INFORMATION INTO SESSION
Session("iUserOffset") = request.cookies("tz")
%>



<html>

<head>

  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="Javascript" src="../../scripts/modules.js"></script>

  <script language="Javascript" > 
  <!--

	//Set timezone in cookie to retrieve later
	var d=new Date();
	if (d.getTimezoneOffset)
	{
		var iMinutes = d.getTimezoneOffset();
		document.cookie = "tz=" + iMinutes;
	}

  //-->


  function save_address(){

	// CHECK TO SEE IF WE HAVE ADDRESS ON FILE OR IF IT IS A CUSTOM ADDRESS
	if (document.frmlocation.select_address.options[document.frmlocation.select_address.selectedIndex].value!=0) 
		{
			// IF SELECTED ADDRESS CLEAR OUT CUSTOM ADDRESS FIELD
			document.frmlocation.ques_issue2.value = '';

		}
	else
		{
			// SUBMIT FORM AS IS
		}

  }
  </script>

  <STYLE>
		div.correctionsbox {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
		div.correctionsboxnotfound  {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
		td.correctionslabel {font-weight:bold;}
		th.corrections {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
		input.correctionstextbox {border: solid 1px #336699;width:400px;}
		textarea.correctionstextarea {border: solid 1px #336699;width:600px;height:150px;}
		p.instructions {padding: 10px;}
		span.maptrue {color:green;font-weight:bold;}
		span.mapfalse {color:red;font-weight:bold;}
		.savemsg {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
  </STYLE>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >


<% ShowHeader sLevel %>


<!--#Include file="../../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		

		<h3>Edit Request Issue Location</h3>
		<a href="../action_respond.asp?control=<%=request("requestid")%>">Return to Request Details</a> | <a href="../action_line_list.asp">View All Requests List</a> 


		<%
		' DISPLAY TO USER THAT VALUES WERE SAVED
		If request("r") = "save" Then 
			response.write "<P><span class=""savemsg"">Saved " & Now() & ".</span></P>"
		End If

		' GET CONTACT INFORMATION
		SubDrawEditIssueLocationInformation(request("requestid"))
		%>


	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../../admin_footer.asp"-->  

</body>
</html>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' FUNCTION FNDISPLAYUSERINFO(IID)
'--------------------------------------------------------------------------------------------------
Function SubDrawEditIssueLocationInformation(irequestid)

	' CHECK FOR EMPTY OR MISSING USERID
	If IsNull(irequestid) or irequestid="" then
		response.write "<div class=correctionsboxnotfound>No information available for the issue location for this request.</div>"
	Else
			' GET INFORMATION FOR SPECIFIED USER
			sSQL = "SELECT * FROM egov_action_response_issue_location WHERE actionrequestresponseid='" & iRequestID & "'"
			
			' OPEN RECORDSET
			Set oIssueLocation = Server.CreateObject("ADODB.Recordset")
			oIssueLocation.Open sSQL, Application("DSN"), 3, 1
			
			If NOT oIssueLocation.EOF Then

				response.write "<form name=""frmlocation"" action=""correction_issue_location_cgi.asp"" method=""post"">"
				response.write "<div class=shadow>"
				response.write "<table class=tablelist cellspacing=0 cellpadding=2 >"
				response.write "<tr><th class=corrections colspan=2 align=left>&nbsp;Request Issue Location</th></tr>"
				
				' DISPLAY INSTRUCTIONS
				response.write "<tr><td colspan=2><P class=instructions>Please select street number/streetname of problem location from list or select ""*not on list"". Provide any additional information on problem location in the box below.</p></td></tr>"

				' DISPLAY STREET
				sAddress = oIssueLocation("streetnumber") & " " & oIssueLocation("streetaddress")
				response.write "<tr><td class=correctionslabel align=""right"" valign=top>Street:</td><td>"
				Call DisplayAddress( session("orgid"), "R", sAddress )
				response.write "</td></tr>"

				' DISPLAY CITY
				response.write "<tr><td class=correctionslabel align=""right"">City:</td><td><input name=city type=text class=correctionstextbox value=""" & oIssueLocation("city") & """></td></tr>"
				
				' DISPLAY STATE
				response.write "<tr><td class=correctionslabel align=""right"">State:</td><td><input name=state type=text class=correctionstextbox value=""" & oIssueLocation("state") & """></td></tr>"
				
				' DISPLAY ZIP
				response.write "<tr><td class=correctionslabel align=""right"">Zip:</td><td><input name=""zip"" type=text class=correctionstextbox value=""" & oIssueLocation("zip") & """></td></tr>"
				
				' DISPLAY ADDITIONAL INFORMATION
				response.write "<tr><td valign=top class=correctionslabel align=""right"">Additional Information:</td><td><textarea class=""correctionstextarea"" name=""comments"" >" & oIssueLocation("comments") & "</textarea></td></tr>"


				' DISPLAY MAPPING OPTIONS
				' LATITUDE
				'response.write "<tr><td class=correctionslabel align=""right"">Latitude:</td><td><input type=text class=correctionstextbox value=""" & oIssueLocation("latitude") & """></td></tr>"
				
				' LONGITUDE
				'response.write "<tr><td class=correctionslabel align=""right"">Longitude:</td><td><input type=text class=correctionstextbox value=""" & oIssueLocation("longitude") & """></td></tr>"

				' MAPPABLE
				'response.write "<tr><td class=correctionslabel align=""right"">Mapping:</td><td>"
				If TRIM(oIssueLocation("latitude")) <> "" AND TRIM(oIssueLocation("longitude")) <> "" Then
					' IS MAPPABLE
					'response.write "<span class=maptrue><img src=""images\mappable_true.gif""> This location can be shown using ""Map It"" feature.</span>"
				Else
					' IS NOT MAPPABLE
					'response.write "<span class=mapfalse><img src=""images\mappable_false.gif""> Latitude and\or longitude value(s) missing. Location cannot be shown using ""Map It"" feature.</span>"
				End If
				response.write "</td></tr>"

				' DISPLAY SAVE AND CANCEL BUTTONS
				response.write "<tr><td class=correctionslabel align=""left"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='../action_respond.asp?control=" & request("requestid") & "';""></td><td align=""right""><input type=submit value=""Save"" >&nbsp;&nbsp;</td></tr>"


				response.write "</table>"
				response.write "</div>"
				response.write "<input type=hidden value=""" & request("status") & """ name=""status"">"
				response.write "<input type=hidden value=""" & iRequestID & """ name=""irequestid"">"
				response.write "</form>"

			Else
				' NO MATCHING USER FOUND
				response.write "<div class=correctionsboxnotfound>No information available for the issue location for this request.</div>"
			End If
		End If

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION DISPLAYADDRESS( IORGID, SRESIDENTTYPE, SADDRESS )
'--------------------------------------------------------------------------------------------------
Function DisplayAddress( iorgid, sResidenttype, sAddress )
	
	Dim sNumber, sSQL, oAddressList, blnFound

	' GET LIST OF ADDRESSES FOR ORGANIZATION
	sSQL = "SELECT residentaddressid, isnull(residentstreetnumber,'') as residentstreetnumber, residentstreetname FROM egov_residentaddresses where orgid = " & iorgid & " and residenttype='" & sResidenttype & "' and residentstreetname is not null order by residentstreetname, residentstreetnumber"
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

	' DISPLAY ADDRESS SELECT BOX
	response.write "<select name=""select_address"" onChange=""save_address();"">"
	response.write  vbcrlf & vbtab & "<option value=""0000"">*not on list</option>"
	
	' LOOP THRU RESIDENT ADDRESS FOR CITY
	Do While NOT oAddressList.EOF 

		' CHECK TO SEE IF WE HAVE MATCHING ADDRESS
		If UCASE(sAddress) = UCASE(oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")) Then
			sSelected  = " SELECTED "
			blnFound = True
		Else
			sSelected = " "
		End If

		response.write vbcrlf & vbtab & "<option " & sSelected & " value=""" & oAddressList("residentaddressid") & """>"
		
		If oAddressList("residentstreetnumber") <> "" Then 
			response.write oAddressList("residentstreetnumber") & " " 
		End If 
		
		response.write oAddressList("residentstreetname") & "</option>"
		
		oAddressList.MoveNext
	Loop

	response.write vbcrlf & "</select>"


	' DISPLAY OTHER TEXTBOX
	response.write " <br> - Or Other Not Listed - <br> "
	If NOT blnFound Then
		' ADDRESS NOT FOUND IN ADDRESS LIST SHOW IN OTHER BOX
		response.write "<input class=correctionstextbox align=""right"" name=""ques_issue2"" type=""text"" value=""" & trim(sAddress) & """ />"
	Else
		' ADDRESS FOUND IN ADDRESS LIST
		response.write "<input class=correctionstextbox align=""right"" name=""ques_issue2"" type=""text""   />"
	End If

	
	' CLEAN UP OBJECTS
	oAddressList.close
	Set oAddressList = Nothing 


End Function
%>


