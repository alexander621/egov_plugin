<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: newpermit.asp
' AUTHOR: Steve Loar
' CREATED: 01/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates permits
'
' MODIFICATION HISTORY
' 1.0   02/14/2008	Steve Loar - INITIAL VERSION
' 1.1	04/01/2008	Steve Loar - Table structure changed. No landvalue, totalvalue, tax district; added streetdirection
' 2.0	09/23/2008	Steve Loar - Changed to allow contractors to be applicants
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sTitle, iPermitInspectionTypeid, sPermitInspectionType, sInspectionDescription, sIsBuildingPermitType
Dim bFound, sIsFinal, sAddress, sStreetNumber, sStreetName, iResidentAddressId, sUnit

sLevel = "../" ' Override of value from common.asp
bFound = False 

PageDisplayCheck "create building permits", sLevel	' In common.asp

If request("permitaddresstypeid") <> "" Then
	iResidentAddressId = CLng(request("permitaddresstypeid"))
Else
	iResidentAddressId = CLng(0)
End If 

'Check for OrgFeatures
 lcl_orghasfeature_hidenewaddress = orghasfeature("hidenewaddress")
%>
<html>
<head>
	<title>E-Gov Administration Console {New Permit}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

<%	If CLng(iResidentAddressId) = CLng(0) Then %>
<style type="text/css">
  input#editaddress,
  input#copyaddress,
		input#newaddress {
		   visibility: hidden;
		}

		span#applicantaddresssearch {
			visibility: hidden;
			}

		div#addressdisplay {
			display: none;
			}

		div#locationdisplay {
			display: none;
			}

		div#neither {
			display: none;
			}

		textarea#location {
			width: 500px;
			height: 80px;
		}

	</style>
<%	End If %>

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>

<!--jquery-1.4.2.min.js-->
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		var w = (screen.width - 640)/2;
		var h = (screen.height - 450)/2;

		//document.getElementById('edituser').style.visibility = 'hidden';
		//document.getElementById('newuser').style.visibility = 'hidden';

 function EditAddress() {
		 myRand = parseInt(Math.random() * 99999999 );
			//eval('window.open("residentaddressedit.asp?residentaddressid=' + $("#residentaddressid").val() + '&rand=' + myRand + '", "_picker", "width=900,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('residentaddressedit.asp?residentaddressid=' + $("#residentaddressid").val() + '&rand=' + myRand, 'Edit Address', 50, 70);
 }

 function copyAddress() {
   var lcl_copyid = $('#residentaddressid').val();

   var lcl_url  = 'residentaddressedit.asp';
       lcl_url += '?residentaddressid=0';
       lcl_url += '&copyid=' + lcl_copyid;

   //eval('window.open("' + lcl_url + '", "_picker", "width=900,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal(lcl_url, 'Copy Address', 50, 70);
 }

	function NewAddress() {
  	//eval('window.open("residentaddressedit.asp?residentaddressid=0", "_picker", "width=900,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('residentaddressedit.asp?residentaddressid=0', 'New Address', 50, 70);
	}

		function searchName()
		{
			if ($("#searchname").val() != "")
			{
				// Try to get a drop down of names
				doAjax('getpermitapplicants.asp', 'searchname=' + $("#searchname").val(), 'UpdateApplicants', 'get', '0');
			}
			else
			{
				alert('Please enter a name before searching.');
				$("#searchname").focus();
			}
		}

		function UpdateApplicants( sResult )
		{
			// change the pick list
			$("#applicant").html( sResult );

			// show and hide buttons to match what was returned
			if (sResult.indexOf('No Match Found') > -1)
			{
				// this is 'not found'
				$("#edituser").css({visibility: "hidden"});
				$("#newuser").css({visibility: "visible"});
				$("#newcontractor").css({visibility: "visible"});
				$("#neworganization").css({visibility: "visible"});
				$("#applicantaddresssearch").css({visibility: "hidden"}); 
			}
			else
			{
				// this is 'some found'
				$("#edituser").css({visibility: "visible"});
				$("#newuser").css({visibility: "visible"});
				$("#newcontractor").css({visibility: "visible"});
				$("#neworganization").css({visibility: "visible"});

				var strPickedApplicant = document.frmPermit.userid.options[0].value;
				//alert(strPickedApplicant.substring(0,1));
				if (strPickedApplicant.substring(0,1) == 'U')
				{
					$("#applicantaddresssearch").css({visibility: "visible"}); 
				}
				else
				{
					$("#applicantaddresssearch").css({visibility: "hidden"}); 
				}
			}

		}

		function toggleAddressSearch()
		{
			//var strPickedApplicant = document.frmPermit.userid.options[document.frmPermit.userid.selectedIndex].value;
			var strPickedApplicant = $("#userid").val();
			if (strPickedApplicant.substring(0,1) == 'U')
			{
				//$('applicantaddresssearch').style.visibility = 'visible';
				$("#applicantaddresssearch").css({visibility: "visible"}); 
			}
			else
			{
				//$('applicantaddresssearch').style.visibility = 'hidden';
				$("#applicantaddresssearch").css({visibility: "hidden"}); 
			}

		}

		function searchStreet()
		{
			if ($("#searchstreet").val() == "" && $("#searchnumber").val() == "" && $("#searchowner").val() == "" )
			{
				alert('Please enter either a number, street name or listed owner before searching.');
				//document.frmPermit.searchstreet.focus();
				$("#searchstreet").focus();
			}
			else 
			{
				// Try to get a drop down of names
				doAjax('getaddresses.asp', 'searchstreet=' + $("#searchstreet").val() + '&searchnumber=' + $("#searchnumber").val() + '&searchowner=' + $("#searchowner").val(), 'UpdateAddresses', 'get', '0');
			}
			
		}

		function UpdateAddresses( sResult )
		{
			//document.getElementById('residentaddress').innerHTML = sResult;
			$("#residentaddress").html( sResult );
<%		if not lcl_orghasfeature_hidenewaddress then	%>
			if (sResult.indexOf('No Match Found') > -1)
			{
				//document.getElementById('editaddress').style.visibility = 'hidden';
				//document.getElementById('newaddress').style.visibility = 'visible';
				$("#editaddress").css({visibility: "hidden"});
				$("#copyaddress").css({visibility: "hidden"});
				$("#newaddress").css({visibility: "visible"});
			}
			else
			{
				//document.getElementById('editaddress').style.visibility = 'visible';
				//document.getElementById('newaddress').style.visibility = 'visible';
				$("#editaddress").css({visibility: "visible"});
				$("#copyaddress").css({visibility: "visible"});
				$("#newaddress").css({visibility: "visible"});
			}
<%		end if		%>
		}

		function EditApplicant()
		{
			//var strPickedApplicant = document.frmPermit.userid.options[document.frmPermit.userid.selectedIndex].value;
			var strPickedApplicant = $("#userid").val();
			//alert(strPickedApplicant.substring(0,1));
			//alert(strPickedApplicant.substring(1));
			myRand = parseInt(Math.random() * 99999999 );
			if (strPickedApplicant.substring(0,1) == 'U')
			{
				//eval('window.open("permitapplicantedit.asp?userid=' + strPickedApplicant.substring(1) + '&detailid=applicantdetails&rand=' + myRand + '", "_picker", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
				showModal('permitapplicantedit.asp?userid=' + strPickedApplicant.substring(1) + '&detailid=applicantdetails&rand=' + myRand , 'Edit Applicant', 20, 30);
			}
			else
			{
				//eval('window.open("permitapplicantcontacttypeedit.asp?permitcontacttypeid=' + strPickedApplicant.substring(1) +'&type=isapplicant&rand=' + myRand + '", "_picker", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
				showModal('permitapplicantcontacttypeedit.asp?permitcontacttypeid=' + strPickedApplicant.substring(1) +'&type=isapplicant&rand=' + myRand , '', 50, 70);
			}
		}

		function NewApplicant()
		{
			//eval('window.open("permitapplicantedit.asp?userid=0&detailid=applicantdetails", "_picker", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
			showModal('permitapplicantedit.asp?userid=0&detailid=applicantdetails&updatetitle=yes' , 'New Applicant', 50, 70);
		}

		function NewContractor()
		{
			//eval('window.open("permitcontactedit.asp?permitcontactid=0&isorganization=0&type=isapplicant", "_picker", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
			showModal('permitcontactedit.asp?permitcontactid=0&isorganization=0&type=isapplicant&createpermit=yes' , 'New Contractor', 20, 30);
		}

		function NewOrganization()
		{
			//eval('window.open("permitcontactedit.asp?permitcontactid=0&isorganization=1&type=isapplicant", "_picker", "width=900,height=600,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
			showModal('permitcontactedit.asp?permitcontactid=0&isorganization=1&type=isapplicant&createpermit=yes' , 'New Organization', 20, 30);
		}

		function validate()
		{
			//alert("validating");
			// Make sure an applicant was selected by seeing if there is an edit button and it is visible
			//  If so, there is a list of valid names and one is always selected
			if ( $("#userid").is(':hidden') )
			{
				//alert(document.getElementById('userid').type);
				alert("Please select an applicant by searching for a name match first.");
				$("#searchname").focus();
				return false;
			}

			// Check that some userid has been selected from the dropdown list
			if ( $("#userid").val() == "0" )
			{
				alert("Please select an applicant by searching for a name match first.");
				$("#searchname").focus();
				return false;
			}

			// Make sure a permit type was selected
			if ($("#permittypeid").val() == "0")
			{
				alert("Please select a permit type from the list.");
				$("#permittypeid").focus();
				return false;
			}

			if ($("#locationrequired").val() == 'address')
			{
				// Check the address that was input
				if ($("#residentaddressid").val() == 0)
				{
					alert("This permit type requires an address.\nPlease select an address.");
					$("#searchstreet").focus();
					return false;
				}
			}

			if ($("#locationrequired").val() == 'location')
			{
				if ($("#location").val() == '')
				{
					alert("This permit type requires a location.\nPlease enter the location information.");
					$("#location").focus();
					return false;
				}
			}

			//alert("Validated OK.");
			document.frmPermit.submit();
		}

		function GetApplicantNumber()
		{
			var strPickedApplicant = $("#userid").val();
			//alert(strPickedApplicant.substring(0,1));
			//alert(strPickedApplicant.substring(1));
			doAjax('getapplicantaddressnumber.asp', 'userid=' + strPickedApplicant.substring(1), 'GetApplicantName', 'get', '0');
		}

		function GetApplicantName( sReturn )
		{
			//alert(sReturn);
			if (sReturn != 'NONUMBER')
			{
				$("#searchnumber").val(sReturn);
			}
			else
			{
				$("#searchnumber").val('');
			}

			var strPickedApplicant = $("#userid").val();
			//alert(strPickedApplicant.substring(1));
			doAjax('getapplicantaddressname.asp', 'userid=' + strPickedApplicant.substring(1), 'doApplicantAddressSearch', 'get', '0');
		}

		function doApplicantAddressSearch( sReturn )
		{
			// change this to only happen if the location requirement is 'address'
			if (sReturn != 'NONAME')
			{
				$("#searchstreet").val(sReturn);
			}
			else
			{
				$("#searchstreet").val('');
			}

			if ( $("#searchstreet").val() != '' || $("#searchnumber").val() != '' )
			{
				searchStreet();
			}
			else
			{
				alert('There is no address for the selected applicant.\nPlease try entering the address, then doing a search.');
				$("#searchnumber").focus();
			}
		}

		function getLocationRequirement()
		{
			// Ajax Call to get the location requirement for the selected permit type
			doAjax('getlocationrequirement.asp', 'permittypeid=' + $("#permittypeid").val(), 'showRequiredLocations', 'get', '0');
		}

		function showRequiredLocations( sReturn )
		{
			$("#locationrequired").val(sReturn);

			if ( sReturn == 'address' )
			{
				$("#locationdisplay").hide("slow");
				$("#neither").hide("slow");
				$("#introtext").hide("slow", function() {
   				$("#addressdisplay").show("slow");
    });
				//$("#location").val('');
			}

			if ( sReturn == 'location' )
			{
				$("#addressdisplay").hide("slow");
				$("#neither").hide("slow");
				$("#introtext").hide("slow", function() {
   				$("#locationdisplay").show("slow");
    });
			}

			if ( sReturn == 'none' )
			{
				$("#addressdisplay").hide("slow");
				$("#locationdisplay").hide("slow");
				$("#introtext").hide("slow", function() {
   				$("#neither").show("slow");
    });
				//$("#location").val('');
			}

			if ( sReturn == 'introtext' )
			{
				$("#locationdisplay").hide("slow");
				$("#neither").hide("slow");
				$("#addressdisplay").hide("slow", function() {
   				$("#introtext").show("slow");
    });
				//$("#location").val('');
			}
		}


		$(document).ready(function(){
			setMaxLength();
		});

	//-->
	</script>

</head>
<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>New Permit Request</strong></font><br /><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: EDIT FORM-->

		<form name="frmPermit" action="newpermitcreate.asp" method="post">
			<input type="hidden" name="permitaddresstypeid" value="<%=iResidentAddressId%>" />
			<input type="hidden" id="locationrequired" name="locationrequired" value="introtext" />
		
		<p>
			<fieldset>
				<legend><font size="+1"><strong>Applicant</strong></font></legend>
				<p>
					<input type="text" id="searchname" name="searchname" value="" size="50" maxlength="50" onkeypress="if(event.keyCode=='13'){searchName();return false;}" /> &nbsp;&nbsp; 
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="searchName();" /><br /><br />
					<table id="applicantpick" cellpadding="0" cellspacing="0" border="0"><tr><td>
						<span id="applicant"><input type="hidden" value="0" name="userid" id="userid" />Search for a contractor or a registered user name</span>&nbsp;
					</td><td>
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="edituser" value="Edit Applicant" onclick="EditApplicant();" />&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="newuser" value="New Individual" onclick="NewApplicant();" />&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="newcontractor" value="New Contractor" onclick="NewContractor();" />&nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" id="neworganization" value="New Organization" onclick="NewOrganization();" />
					</td></tr></table>
				</p>
			</fieldset>
		</p>

		<p>
			<fieldset>
				<legend><font size="+1"><strong>Permit Type</strong></font></legend>
				<p>
<%					ShowPermitTypePicks 0 %>					
				</p>
			</fieldset>
		</p>

		<p>
			<fieldset>
				<legend><font size="+1"><strong>Address/Location</strong></font></legend>

				<div id="introtext">
					<p>
						<strong>The permit type selected above may require an address or a location before the permit can be created.</strong>
					</p>
				</div>

				<div id="addressdisplay">
					<table cellpadding="2" cellspacing="0" border="0" id="permitlocationsearch">
						<tr>
							<td><input type="text" id="searchnumber" name="searchnumber" value="" size="10" maxlength="10" onkeypress="if(event.keyCode=='13'){searchStreet();return false;}" /><br />number
							</td>
							<td><input type="text" id="searchstreet" name="searchstreet" value="" size="50" maxlength="50" onkeypress="if(event.keyCode=='13'){searchStreet();return false;}" /><br />street name
							</td>
							<td><input type="text" id="searchowner" name="searchowner" value="" size="25" maxlength="25" onkeypress="if(event.keyCode=='13'){searchStreet();return false;}" /><br />listed owner
							</td>
							<td valign="top" nowrap="nowrap"><input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="searchStreet();" />&nbsp;
							<span id="applicantaddresssearch">
								<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search for Applicant's Address" onclick="GetApplicantNumber();" />
							</span>
							</td>
						</tr>
					</table><br /><br />
<%
      response.write "<span id=""residentaddress"">" & vbcrlf

      if clng(iResidentAddressId) > clng(0) then
  							ShowPreSelectedAddress session("orgid"), iResidentAddressId
      else
         response.write "  <input type=""hidden"" value=""0"" name=""residentaddressid"" id=""residentaddressid"" />Search by street number, street name, or listed owner" & vbcrlf
      end if

      response.write "</span>&nbsp;&nbsp;" & vbcrlf

      if not lcl_orghasfeature_hidenewaddress then
         response.write "  <input type=""button"" class=""button ui-button ui-widget ui-corner-all"" id=""editaddress"" value=""Edit Address"" onclick=""EditAddress();"" />" & vbcrlf
         response.write "  <input type=""button"" class=""button ui-button ui-widget ui-corner-all"" id=""copyaddress"" value=""Copy Address"" onclick=""copyAddress();"" />" & vbcrlf
         response.write "  <input type=""button"" class=""button ui-button ui-widget ui-corner-all"" id=""newaddress"" value=""New Address"" onclick=""NewAddress();"" />" & vbcrlf
      end if
%>
				</div>
				
				<div id="locationdisplay">
					<table cellpadding="2" cellspacing="0" border="0" id="locationentry">
						<tr>
							<td valign="top"><strong>Location:</strong></td>
							<td>
								<textarea id="location" name="location" maxlength="1000"></textarea>
							</td>
						</tr>
					</table>
				</div>
				
				<div id="neither">
					<p>
						<strong>This permit type does not require an address or a location.</strong>
					</p>
				</div>

			</fieldset>
		</p>

		<p>
			<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="validate();" value="Create Permit Request" />
		</p>

		</form>
		<!--END: EDIT FORM-->

		</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void ShowPreSelectedAddress( iResidentAddressId )
'------------------------------------------------------------------------------
sub ShowPreSelectedAddress( ByVal iOrgID, ByVal iResidentAddressId )
	dim sSql, oRs

	sSQL = "SELECT residentaddressid, "
 sSQL = sSQL & " residentstreetnumber, "
 sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, "
 sSQL = sSQL & " ISNULL(streetdirection,'') AS streetdirection, "
 sSQL = sSQL & " ISNULL(residentunit,'') AS residentunit " 
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & iOrgID
	sSQL = sSQL & " AND residentaddressid =  " & iResidentAddressId 

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	if not oRS.eof then
  		response.write vbcrlf & "<select name=""residentaddressid"" id=""residentaddressid"">" & vbcrlf

    do while not oRs.EOF
       lcl_option_displaytext = oRs("residentstreetnumber")

    			if oRs("residentstreetprefix") <> "" then
      				lcl_option_displaytext = lcl_option_displaytext & " " & oRs("residentstreetprefix") 
    			end if

      				lcl_option_displaytext = lcl_option_displaytext & " " & oRs("residentstreetname")

    			if oRs("streetsuffix") <> "" then
      				lcl_option_displaytext = lcl_option_displaytext & " " & oRs("streetsuffix")
    			end if

    			if oRs("streetdirection") <> "" then
      				lcl_option_displaytext = lcl_option_displaytext & " " & oRs("streetdirection")
       end if

    			if oRs("residentunit") <> "" then
      				lcl_option_displaytext = lcl_option_displaytext & ", " & oRs("residentunit")
    			end if

      'Clean up before display option
       lcl_option_displaytext = trim(lcl_option_displaytext)

       if lcl_option_displaytext = "," then
          lcl_option_displaytext = ""
       end if

    			response.write "  <option value='" & oRs("residentaddressid") & "'>" & lcl_option_displaytext & "</option>" & vbcrlf

       oRs.movenext
    loop

  		response.write "</select>" & vbcrlf
 end if

	oRs.close
 set oRs = nothing

end sub

'------------------------------------------------------------------------------
' void ShowPermitTypePicks( iPermitTypeId )
'------------------------------------------------------------------------------
Sub ShowPermitTypePicks( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, ISNULL(permittype,'') AS permittype, ISNULL(permittypedesc,'') AS permittypedesc "
	sSql = sSql & " FROM egov_permittypes "
	'sSql = sSql & " WHERE isbuildingpermittype = 1 AND orgid = "& session("orgid")
	sSql = sSql & " WHERE orgid = "& session("orgid")
	sSql = sSql & " ORDER BY permittype, permittypedesc, permittypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""permittypeid"" name=""permittypeid"" onchange=""getLocationRequirement();"">"
		If CLng(iPermitTypeId) = CLng(0) Then
			response.write vbcrlf & "<option value=""0"">Please select a permit type...</option>"
		End If 
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value="""  & oRs("permittypeid") & """"
			If CLng(iPermitTypeId) = CLng(oRs("permittypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permittype") 
			If oRs("permittype") <> "" And oRs("permittypedesc") <> "" Then 
				response.write " &ndash; "
			End If 
			response.write oRs("permittypedesc")
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write vbcrlf & "There are No Permit Types to select."
		response.write vbcrlf & "<input type=""hidden"" id=""permittypeid"" name=""permittypeid"" value=""0"" />"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
