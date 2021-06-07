<!DOCTYPE html>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="rentals/rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getuseraddress.asp
' AUTHOR: Steve Loar
' CREATED: 10/02/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This collects missing address information for Montgomery. It is used in Rentals and 
'				Facilities
'
' MODIFICATION HISTORY
' 1.0   10/02/2013	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId

'Force the page to be re-loaded on back button
response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button after purchase problems

' we need to pass in the userid, orgid, URL to return to (in session var)
iUserId = clng(request( "userid"))

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If


%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title><%=sTitle%></title>

	<link rel="stylesheet" href="css/styles.css" />
	<link rel="stylesheet" href="global.css" />
	<link rel="stylesheet" href="rentals/rentalstyles.css" />
	<link rel="stylesheet" href="css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

	<script src="scripts/formvalidation_msgdisplay.js"></script>
	<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>

	<script>
		<!--

		function displayScreenMsg( iMsg ) 
		{
			if( iMsg != "" ) 
			{
				$("#screenMsg").html(iMsg);
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("&nbsp;");
		}

		function validate()
		{
			//alert("validating!");
			// no fields can be blank
			var oktosubmit = true;

			if ( $("#useraddress").val() == '')
			{
				inlineMsg("useraddress",'Please provide an address.',10,'useraddress');
				oktosubmit = false;
			}

			if ( $("#usercity").val() == '')
			{
				inlineMsg("usercity",'Please provide a city.',10,'usercity');
				oktosubmit = false;
			}

			if ( $("#userstate").val() == '')
			{
				inlineMsg("userstate",'Please provide a state.',10,'userstate');
				oktosubmit = false;
			}

			if ( $("#userzip").val() == '')
			{
				inlineMsg("userzip",'Please provide a zip code.',10,'userzip');
				oktosubmit = false;
			}

			if ( oktosubmit )
			{
				// submit the form
				document.frmAddressUpdate.submit();
			}
			else
			{
				displayScreenMsg('Some of the information we need is missing. Please provide the missing information and try again.');
			}
		}

		$( document ).ready(function() {

			$("#useraddress").autocomplete({
				source: function( request, response ) {
					//alert( request.term );
					var orgid = $("#orgid").val();
					$("#residentaddressid").val("0");
					$("#addressinsystemmsg").html("We consider this to be a non-resident address.");
					$("#addressinsystemmsg").toggleClass("addressfound", false);
					$("#addressinsystemmsg").toggleClass("addressnotfound", true);
					$.ajax({
						url: "https://api.egovlink.com/api/ActionForm/GetAddressList?callback=?",
						type: "GET",
						dataType: "jsonp",
						contentType: "application/json",
						data: {
							_OrgId : orgid,
							_MatchString: request.term,
							_MaxRows: 16
						},
						success: function( data ) {
							//alert("back");
							response( $.map( data, function( item ) {
								$('#ui-id-1').css('display', 'block');
								return {
									label: item.streetaddress,
									value: item.streetaddress, 
									residentadressid: item.residentadressid
								}
							}));
						}
					});
				},
				minLength: 1,
				select: function( event, ui ) {
					$("#residentaddressid").val( ui.item ? ui.item.residentadressid : "0" );
					$("#addressinsystemmsg").html("We consider this to be a valid resident address.");
					$("#addressinsystemmsg").toggleClass("addressfound", true);
					$("#addressinsystemmsg").toggleClass("addressnotfound", false);
				}
			});

		});

		//-->
	</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "./" ) %>

<div id="page-title">
	<h1>Address Update</h1>
</div>

<div id="screenMsg">&nbsp;</div>

<div class="pagedescription">
	Before continuing with your reservation, we need your address information. Please enter that information below and then we will take you back to where you can continue your reservation.
</div>

<form method="post" name="frmAddressUpdate" action="getuseraddressupdate.asp">
	<input type="hidden" id="userid" name="userid" value="<%=iUserId%>" />
	<input type="hidden" id="orgid" name="orgid" value="<%=iorgid%>" />

    <div class="form-element">
    	<label for="useraddress" >Address: </label>
		<input type="text" id="useraddress" name="useraddress" maxlength="100" value="" placeholder="Address" /> 
		<span id="addressinsystemmsg"></span>
		<div class="formelementdescription">
			Start typing an address in the box above and then select one from the popup list if there is a match.<br />If there is no popup or match, that's OK. Just fill in your correct address.
		</div>
		<input type="hidden" id="residentaddressid" name="residentaddressid" value="0" />

	</div>
	<table class="addressdisplay" cellpadding="2" cellspacing="0" border="0">
		<tr>
			<td class="labelcell"><label for="usercity" >City: </label></td>
			<td><input type="text" id="usercity" name="usercity" maxlength="100" value="<%= sDefaultCity %>" placeholder="City" /></td>
		</tr>
		<tr>
			<td class="labelcell"><label for="userstate" >State: </label></td>
			<td><input type="text" id="userstate" name="userstate" maxlength="2" value="<%= sDefaultState %>" placeholder="State" /></td>
		</tr>
		<tr>
			<td class="labelcell"><label for="userzip" >Zip: </label></td>
			<td><input type="text" id="userzip" name="userzip" maxlength="10" value="<%= sDefaultZip %>" placeholder="Zip" /></td>
		</tr>
	</table>	

	<input type="button" class="button" value="Save Changes" onclick="validate();" />

</form>

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="include_bottom.asp"-->  
