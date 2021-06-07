
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalslist.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals. From here you can create or edit rentals
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
' 1.1	03/24/2011	Steve Loar - deactivated functions added
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, sTitle, sRentalName, sDescription, sPublicCanView, sPublicCanReserve, sNeedsApproval
Dim iLocationid, sWidth, sLength, sCapacity, iSupervisorUserId, sReceiptNotes, sImagesPlacement
Dim sHasOffSeason, sOffSeasonStartMonth, sOffSeasonStartDay, sOffSeasonEndMonth, sOffSeasonEndDay
Dim iMaxImages, iMaxDocuments, iMaxRentals, sButtonValue, sLoadMsg, sShortDescription, iMaxItems
Dim sTerms, sResidentRentalPeriod, sNonResidentRentalPeriod, iMaxFeeRows, sOffSeasonEndYear
Dim sNoCostToRent, sIconImageUrl, sReservationsDuringSeason, sNonResidentsWait, sNonresidentWaitDays
Dim sNonresidentStartDate, bOrgHasAccounts, iMaxAlertRows, sDisabledAdd, sCheckUseHTMLonLong
Dim sDeactivatedCheck

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create edit rentals", sLevel	' In common.asp

iRentalId = CLng(request("rentalid"))

If iRentalId = CLng(0) Then
	sTitle = "Create A Rental"
	sButtonValue = "Create Rental"
Else
	sTitle = "Edit Rental"
	sButtonValue = "Save Changes"
End If 

blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")

sRentalName = ""
sDescription = ""
sPublicCanView = ""
sPublicCanReserve = ""
sNeedsApproval = ""
sNoCostToRent = ""
sIconImageUrl = ""
iLocationid = 0
sWidth = ""
sLength = ""
sCapacity = ""
iSupervisorUserId = 0
sReceiptNotes = ""
sImagesPlacement = "none"
sHasOffSeason = ""
sOffSeasonStartMonth = "1"
sOffSeasonStartDay = "1"
sOffSeasonEndMonth = "1"
sOffSeasonEndDay = "31"
sOffSeasonEndYear = "0"
iMaxImages = 1
iMaxDocuments = 1
iMaxRentals = 1
sLoadMsg = ""
sShortDescription = ""
sTerms = ""
sResidentRentalPeriod = ""
sNonResidentRentalPeriod = ""
sReservationsDuringSeason = ""
sNonResidentsWait = ""
sNonresidentWaitDays = ""
sNonresidentStartDate = ""
iMaxAlertRows = 0
sCheckUseHTMLonLong = ""
sDeactivatedCheck = ""

GetRentalValues iRentalId

If request("s") <> "" Then
	If request("s") = "n" Then
		sLoadMsg = "displayScreenMsg('This Rental Was Successfully Created');"
	End If
	If request("s") = "u" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If 
	If request("s") = "c" Then
		sLoadMsg = "displayScreenMsg('The Copy Was Successful');"
	End If 
End If 

' Not every org has general ledger accounts so we need to be able to hide/show accordingly.
bOrgHasAccounts = OrgHasFeature("gl accounts")

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--
		
		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		function getDays( sMonthPickName, sSpanId, sDayPickName )
		{
			//alert( sSpanId );
			var iMonth;
			iMonth = $("#" + sMonthPickName).val();
			//alert( iMonth );
			doAjax( 'getdaysinmonth.asp', 'imonth=' + iMonth + '&spickname=' + sDayPickName, 'Replace' + sSpanId, 'get', '0');
		}

		function Replaceoffseasonstart( sResults )
		{
			//alert( sResults );
			$("#offseasonstart").html(sResults); 
		}

		function Replaceoffseasonend( sResults )
		{
			//alert( sResults );
			$("#offseasonend").html(sResults);
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html("*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;");
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("");
		}

		function SetUpPage()
		{
			setMaxLength();
			<%=sLoadMsg%>
			$("#rentalname").focus();
		}

		function doImagePicker( sFormField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("imagepicker/default.asp?name=frmRental.' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doImageView( sImageurl )
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("imagedisplay.asp?url=' + sImageurl + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doDocumentPicker( sFormField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("documentpicker/default.asp?name=frmRental.' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function storeCaret (textEl) 
		{
		   if (textEl.createTextRange)
			 textEl.caretPos = document.selection.createRange().duplicate();
		 }

		 function insertAtURL (textEl, text) 
		 {
		   if (textEl.createTextRange && textEl.caretPos) 
		   {
			 var caretPos = textEl.caretPos;
			 caretPos.text =
			   caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				 text + ' ' : text;
		   }
		   else
			 textEl.value  = text;
			
			//document.getElementById(textEl.name + 'pic').src = text;
			$("#" + textEl.name + "pic").attr("src",text);
			if (textEl.name.indexOf("document") >= 0)
			{
				$("#" + textEl.name + 'pic').html('<a href="' + text + '" target="_newwindow">View Document</a>&nbsp;&nbsp;');
			}
		 }

		 function AddImageRow()
		{
			document.frmRental.maximages.value = parseInt(document.frmRental.maximages.value) + 1;
			var tbl = document.getElementById("imagelist");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmRental.maximages.value);
			var row = tbl.insertRow(lastRow);

			// Remove Image checkbox
			var newCell = row.insertCell(0);
			newCell.align = 'center';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removeimage' + newRow;
			e.id = 'removeimage' + newRow;
			newCell.appendChild(e);

			// Image URL
			newCell = row.insertCell(1);
			newCell.align = 'left';
			e = document.createElement('input');
			<%if blnHasWP then%>
			e.type = 'hidden';
			<%else%>
			e.type = 'text';
			<% end if %>
			e.name = 'imageurl' + newRow;
			e.id = 'imageurl' + newRow;
			e.size = '60';
			e.maxlength = '250';
			e.classList.add("imageurl");
			newCell.appendChild(e);

			// Image

			var img = document.createElement('img');
    			img.src = "../images/placeholder.png";
			img.width = "240";
			img.height = "180";
			img.align = "middle";
			img.id = "imageurl" + newRow + "pic";
			img.onerror = "this.src = '../images/placeholder.png';";
			newCell.appendChild(img);

			// Pick button
			e = document.createElement('input');
			e.type = 'button';
			e.value = 'Pick';
			e.className = 'button';
			<%if blnHasWP then%>
			e.onclick = function() { showModal('Pick Image', 65, 80, 'imageurl' + newRow ); };
			<% else %>
			e.onclick = function() { doImagePicker('imageurl' + newRow ); };
			<% end if %>
			newCell.appendChild(e);

			// Alt Tag
			newCell = row.insertCell(2);
			newCell.align = 'center';
			e = document.createElement('input');
			e.type = 'text';
			e.name = 'alttag' + newRow;
			e.id = 'alttag' + newRow;
			e.size = '30';
			e.maxlength = '100';
			newCell.appendChild(e);
			
			// Find the first row that exists
			for (var t = 1; t <= parseInt(document.frmRental.maximages.value); t++ )
			{
				if (document.getElementById("displayorder" + t))
				{
					break;
				}
			}
			var slength = document.getElementById("displayorder" + t).length + 1;
			var op;
			var newText;
			// Add the new display order to the existing display order picks
			var newDisplayOrder = parseInt(document.frmRental.maximages.value);
			for ( var o=1; o < newRow; o++)
			{
				if (document.getElementById("displayorder" + o))
				{
					op = document.createElement('OPTION');
					newText = document.createTextNode( newDisplayOrder );
					op.appendChild( newText );
					op.setAttribute( 'value', newDisplayOrder );
					document.getElementById('displayorder' + o).appendChild(op);
				}
			}

			// The display order pick
			newCell = row.insertCell(3);
			newCell.align = 'center';
			e = document.createElement('select');
			e.name = 'displayorder' + newRow;
			e.id = 'displayorder' + newRow;
			newCell.appendChild(e);
			//slength = document.getElementById("displayorder" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("displayorder" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("displayorder" + t).options[s].value );
				op.selected = true;
				e.appendChild(op);
			}

			// if we have six rows then disable the add button
			if (parseInt(document.frmRental.maximages.value) > parseInt(5))
			{
				$("#addanimage").prop( "disabled", true );
			}

			// Put them on the new row
			document.getElementById('imageurl' + newRow).focus();
			refreshImageURLListener();
		}

		function AddDocument()
		{
			document.frmRental.maxdocuments.value = parseInt(document.frmRental.maxdocuments.value) + 1;
			var tbl = document.getElementById("documentlist");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmRental.maxdocuments.value);
			var row = tbl.insertRow(lastRow);

			// Remove Document checkbox
			var newCell = row.insertCell(0);
			newCell.align = 'center';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removedocument' + newRow;
			e.id = 'removedocument' + newRow;
			newCell.appendChild(e);

			// Image URL
			newCell = row.insertCell(1);
			newCell.align = 'center';
			e = document.createElement('input');
			<% if blnHasWP then %>
			e.type = 'hidden';
			<% else %>
			e.type = 'text';
			<% end if %>
			e.name = 'documenturl' + newRow;
			e.id = 'documenturl' + newRow;
			e.size = '60';
			e.maxlength = '250';
			newCell.appendChild(e);

			//span target
			e = document.createElement('span');
			e.id = "documenturl" + newRow + "pic";
			newCell.appendChild(e);

			// Pick button
			e = document.createElement('input');
			e.type = 'button';
			e.value = 'Pick';
			e.className = 'button';
			<% if blnHasWP then %>
			e.onclick = function() { showModal('Pick File', 65, 80, 'documenturl' + newRow ); };
			<% else %>
			e.onclick = function() { doDocumentPicker('documenturl' + newRow ); };
			<% end if %>
			newCell.appendChild(e);

			// Alt Tag
			newCell = row.insertCell(2);
			newCell.align = 'center';
			e = document.createElement('input');
			e.type = 'text';
			e.name = 'documenttitle' + newRow;
			e.id = 'documenttitle' + newRow;
			e.size = '40';
			e.maxlength = '250';
			newCell.appendChild(e);

			// Put them on the new row
			document.getElementById('documenturl' + newRow).focus();
		}

		function AddRentalRow()
		{
			document.frmRental.maxrentals.value = parseInt(document.frmRental.maxrentals.value) + 1;
			var tbl = document.getElementById("rentallist");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmRental.maxrentals.value);
			var row = tbl.insertRow(lastRow);

			// Remove Rental checkbox
			var newCell = row.insertCell(0);
			newCell.align = 'center';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removerental' + newRow;
			e.id = 'removerental' + newRow;
			newCell.appendChild(e);

			//inspection type pick here
			newCell = row.insertCell(1);
			newCell.align = 'center';
			e = document.createElement('select');
			e.name = 'associatedrentalid' + newRow;
			e.id = 'associatedrentalid' + newRow;
			newCell.appendChild(e);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmRental.maxrentals.value); t++ )
			{
				if (document.getElementById("associatedrentalid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("associatedrentalid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("associatedrentalid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("associatedrentalid" + t).options[s].value );
				e.appendChild(op);
			}

		}

		function AddItemRow()
		{
			document.frmRental.maxitems.value = parseInt(document.frmRental.maxitems.value) + 1;
			var tbl = document.getElementById("rentalitemlist");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.frmRental.maxitems.value);
			var row = tbl.insertRow(lastRow);

			// Remove Item checkbox
			var newCell = row.insertCell(0);
			newCell.align = 'center';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removeitem' + newRow;
			e.id = 'removeitem' + newRow;
			newCell.appendChild(e);

			// Item
			newCell = row.insertCell(1);
			newCell.align = 'center';
			e = document.createElement('input');
			e.type = 'text';
			e.name = 'rentalitem' + newRow;
			e.id = 'rentalitem' + newRow;
			e.size = '50';
			e.maxlength = '50';
			newCell.appendChild(e);

			//account pick here
			newCell = row.insertCell(2);
			newCell.align = 'center';
			e = document.createElement('select');
			e.name = 'itemaccountid' + newRow;
			e.id = 'itemaccountid' + newRow;
			newCell.appendChild(e);

			// Find the first row that exists
			for (var t = 0; t <= parseInt(document.frmRental.maxitems.value); t++ )
			{
				if (document.getElementById("itemaccountid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("itemaccountid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("itemaccountid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("itemaccountid" + t).options[s].value );
				e.appendChild(op);
			}

			// Max Available
			newCell = row.insertCell(3);
			newCell.align = 'center';
			e = document.createElement('input');
			e.type = 'text';
			e.name = 'maxavailable' + newRow;
			e.id = 'maxavailable' + newRow;
			e.size = '5';
			e.maxlength = '5';
			e.onchange = function() { ValidateMaxAvailable(this); };
			newCell.appendChild(e);

			// Rate per day
			newCell = row.insertCell(4);
			newCell.align = 'center';
			e = document.createElement('input');
			e.type = 'text';
			e.name = 'amount' + newRow;
			e.id = 'amount' + newRow;
			e.size = '7';
			e.maxlength = '7';
			e.onchange = function() { ValidatePrice(this); };
			newCell.appendChild(e);

			// Put them on the new row
			document.getElementById('rentalitem' + newRow).focus();
		}

		function RemoveImages()
		{
			var iRow = 0;
			var tbl = document.getElementById("imagelist");
			// Check the Images rows for any selected for removal
			for (var t = 1; t <= parseInt(document.frmRental.maximages.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removeimage" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removeimage" + t).checked == true)
					{
						if (tbl.rows.length > 2)
						{
							// Remove the unwanted row
							tbl.deleteRow(iRow);
							// Decrement the row counter as we have one less row now
							iRow--;
							if ($("#addanimage").length)
							{
								$("#addanimage").prop( "disabled", false );
							}
						}
						else
						{
							// Down to one row, so just reset it to it's initial defaults
							document.getElementById("removeimage" + t).checked = false;
							document.getElementById("imageurl" + t).value = '';
							document.getElementById("alttag" + t).value = '';
							document.getElementById("displayorder" + t).options[0].selected = true;
						}
					}
				}
			}
		}

		function RemoveDocuments()
		{
			var iRow = 0;
			var tbl = document.getElementById("documentlist");
			// Check the Document rows for any selected for removal
			for (var t = 1; t <= parseInt(document.frmRental.maxdocuments.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removedocument" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removedocument" + t).checked == true)
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
							document.getElementById("removedocument" + t).checked = false;
							document.getElementById("documenturl" + t).value = '';
							document.getElementById("documenttitle" + t).value = '';
						}
					}
				}
			}
		}

		function RemoveRentals()
		{
			var iRow = 0;
			var tbl = document.getElementById("rentallist");
			// Check the Rentals rows for any selected for removal
			for (var t = 1; t <= parseInt(document.frmRental.maxrentals.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removerental" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removerental" + t).checked == true)
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
							document.getElementById("removerental" + t).checked = false;
							document.getElementById("associatedrentalid" + t).options[0].selected = true;
						}
					}
				}
			}
		}

		function RemoveItems()
		{
			var iRow = 0;
			var tbl = document.getElementById("rentalitemlist");
			// Check the Items rows for any selected for removal
			for (var t = 1; t <= parseInt(document.frmRental.maxitems.value); t++)
			{
				// See if a row exists for this one
				if (document.getElementById("removeitem" + t))
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if (document.getElementById("removeitem" + t).checked == true)
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
							document.getElementById("removeitem" + t).checked = false;
							document.getElementById("rentalitem" + t).value = '';
							document.getElementById("itemaccountid" + t).options[0].selected = true;
							document.getElementById("maxavailable" + t).value = '';
							document.getElementById("amount" + t).value = '';
						}
					}
				}
			}
		}

		function validate()
		{
			if ($("#rentalname").val() == '')
			{
				alert("Please provide a name, then try saving again.");
				$("#rentalname").focus();
				return;
			}
			if (document.getElementById('hasoffseason'))
			{
				if (document.getElementById('hasoffseason').checked == true)
				{
					var startmonth = document.getElementById('offseasonstartmonth').value;
					var startday = document.getElementById('offseasonstartday').value;
					var endmonth = document.getElementById('offseasonendmonth').value;
					var endday = document.getElementById('offseasonendday').value;
					var startyear = "<%=year(now())%>";
					var endyear = parseInt(startyear) + parseInt(document.getElementById('offseasonendyear').value);
	
					var startdate = new Date(startmonth + "/" + startday + "/" + startyear);
					var enddate = new Date(endmonth + "/" + endday + "/" + endyear);
	
					if (startdate > enddate)
					{
						alert("Your off season is set to end before it begins.  You may want to change it to resume \"The Next Year\" instead of \"The Same Year\" or correct your dates.");
						return;
					}
				}
			}
			
			document.frmRental.submit();
		}

		function DeleteRental()
		{
			if (confirm('Delete this rental?'))
			{
				location.href='rentaldelete.asp?rentalid=<%=iRentalId%>';
			}
		}

		function ValidatePrice( oPrice )
		{
			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			// Validate the format of the price
			if (oPrice.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPrice.value);
				if ( Ok )
				{
					oPrice.value = format_number(Number(oPrice.value),2);
				}
				else 
				{
					oPrice.value = "";
					alert("Rates must be numbers in currency format or blank.\nPlease correct to continue.");
					setfocus(oPrice);
					return false;
				}
			}
		}

		function ValidateMaxAvailable( oMax )
		{
			// Validate the maximum available amounts
			if (oMax.value != '')
			{
				// Remove any extra spaces
				oMax.value = removeSpaces(oMax.value);
				//Remove commas that would cause problems in validation
				oMax.value = removeCommas(oMax.value);

				rege = /^\d*$/;
				Ok = rege.test(oMax.value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',7);
					oMax.value = "";
					alert("This field must be a positive integer.\nPlease correct to continue.");
					setfocus(oMax);
					return false;
				}
			}
		}

		function ValidateRentalPeriod( oPeriod )
		{
			// Validate the rental periods
			if (oPeriod.value != '')
			{
				// Remove any extra spaces
				oPeriod.value = removeSpaces(oPeriod.value);
				//Remove commas that would cause problems in validation
				oPeriod.value = removeCommas(oPeriod.value);

				rege = /^\d*$/;
				Ok = rege.test(oPeriod.value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',0);
					oPeriod.value = "";
					alert("The rental period must be a positive integer.\nPlease correct to continue.");
					setfocus(oPeriod);
					return false;
				}
			}
		}

		function ValidateNonresidentWait( oPeriod )
		{
			// Validate the Nonresident wait periods
			if (oPeriod.value != '')
			{
				// Remove any extra spaces
				oPeriod.value = removeSpaces(oPeriod.value);
				//Remove commas that would cause problems in validation
				oPeriod.value = removeCommas(oPeriod.value);

				rege = /^\d*$/;
				Ok = rege.test(oPeriod.value);
				if ( ! Ok )
				{
					tabView.set('activeIndex',5);
					oPeriod.value = "";
					alert("The additional days a Nonresident waits must be a positive integer.\nPlease correct to continue.");
					setfocus(oPeriod);
					return false;
				}
			}
		}

		function NewAlertRow()
		{
			$("#maxalertrows").val(parseInt($("#maxalertrows").val()) + 1);
			var tbl = document.getElementById("alerttable");
			var lastRow = tbl.rows.length;
			var newRow = parseInt($("#maxalertrows").val());
			var row = tbl.insertRow(lastRow);

			// Remove Row checkbox
			var cellZero = row.insertCell(0);
			cellZero.className = 'firstcell';
			var e = document.createElement('input');
			e.type = 'checkbox';
			e.name = 'removealert' + newRow;
			e.id = 'removealert' + newRow;
			cellZero.appendChild(e);

			//alert type pick here
			cellZero = row.insertCell(1);
			cellZero.align = 'center';
			var e0 = document.createElement('select');
			e0.name = 'rentalalerttypeid' + newRow;
			e0.id = 'rentalalerttypeid' + newRow;
			cellZero.appendChild(e0);

			// Find the first row that exists
			for (var t = 1; t <= parseInt($("#maxalertrows").val()); t++ )
			{
				if (document.getElementById("rentalalerttypeid" + t))
				{
					break;
				}
			}

			var slength = document.getElementById("rentalalerttypeid" + t).length;
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("rentalalerttypeid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("rentalalerttypeid" + t).options[s].value );
				e0.appendChild(op);
			}

			//notify user pick here
			var cellOne = row.insertCell(2);
			cellOne.align = 'center';
			e1 = document.createElement('select');
			e1.name = 'userid' + newRow;
			e1.id = 'userid' + newRow;
			cellOne.appendChild(e1);
			slength = document.getElementById("userid" + t).length;
			for ( s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("userid" + t).options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("userid" + t).options[s].value );
				e1.appendChild(op);
			}
		}

		function RemoveAlertRows()
		{
			var iRow = 0;
			var tbl = document.getElementById("alerttable");
			// Check the alert rows for any selected for removal
			for (var t = 1; t <= parseInt($("#maxalertrows").val()); t++)
			{
				// See if a row exists for this one
				if ($("#removealert" + t).length)
				{
					// The row exists so increment the row counter
					iRow++;
					// If it is marked for removal, remove it
					if ($("#removealert" + t).is(':checked') == true)
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
							$("#removealert" + t).prop( "checked", false );
							document.getElementById("rentalalerttypeid" + t).options[0].selected = true;
							document.getElementById("userid" + t).options[0].selected = true;
						}
					}
				}
			}
		}


	//-->
	</script>

</head>

<body class="yui-skin-sam" onload="SetUpPage();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%></strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			
			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
				<input type="button" class="button" value="<< Back" onclick="location.href='rentalslist.asp';" /> &nbsp; 
<%				If iRentalId > CLng(0) Then	
					If Not RentalHasReservations( iRentalId ) Then 
%>
						<input type="button" class="button" value="Delete" onclick="DeleteRental();" />

<%					End If	
				End If 
%>
			</td></tr></table>

			<form name="frmRental" action="rentalupdate.asp" method="post">
				<input type="hidden" id="rentalid" name="rentalid" value="<%=iRentalId%>" />
				<input type="hidden" id="maxrentals" name="maxrentals" value="0" />
				<p id="rentalnamecontainer">
					Name: <input type="text" id="rentalname" name="rentalname" size="90" maxlength="90" value="<%=sRentalName%>" />
				</p>

				<div id="demo" class="yui-navset">
					<ul class="yui-nav">
						<li><a href="#tab1"><em>General</em></a></li>
						<li><a href="#tab2"><em>Descriptions</em></a></li>
						<li><a href="#tab3"><em>Images</em></a></li>
						<li><a href="#tab4"><em>Documents</em></a></li>
						<li><a href="#tab5"><em>In Season</em></a></li>
						<li><a href="#tab6"><em>Off Season</em></a></li>
						<li><a href="#tab7"><em>Items</em></a></li>
						<li><a href="#tab8"><em>Fees</em></a></li>
						<li><a href="#tab9"><em>Alerts</em></a></li>
					</ul>            
					<div class="yui-content">

						<div id="tab1"> <!-- General Information -->
							<p>
								<input type="checkbox" id="isdeactivated" name="isdeactivated" <%=sDeactivatedCheck%> /> <strong>This rental is deactivated.</strong> &nbsp; It is not listed and cannot be reserved.
							</p>
							<p>
								<input type="checkbox" id="nocosttorent" name="nocosttorent" <%=sNoCostToRent%> /> <strong>There is no cost to reserve this rental</strong> &nbsp;
							</p>
							<p>
								<input type="checkbox" id="publiccanview" name="publiccanview" <%=sPublicCanView%> /> Public Can View &nbsp;
								<input type="checkbox" id="publiccanreserve" name="publiccanreserve" <%=sPublicCanReserve%> /> Public Can Reserve &nbsp;
								<!--<input type="checkbox" id="needsapproval" name="needsapproval" <%'=sNeedsApproval%> /> Public Reservations Need Approval -->
							</p>
							<p>
								Location: &nbsp; <% ShowLocationPicks iLocationid, false	' In rentalscommonfunctions.asp  %>
							</p>
							<p>
								Categories: <br /> 
								<div id="catagorypicks">
								<% ShowCategoryPicks iRentalId %>
								</div>
							</p>
							<p>
								Dimensions: <input type="text" id="width" name="width" value="<%=sWidth%>" size="20" maxlength="20" /> &nbsp; by &nbsp; 
									<input type="text" id="length" name="length" value="<%=sLength%>" size="20" maxlength="20" />
							</p>
							<p>
								Capacity:<br />
								<textarea id="capacity" name="capacity" maxlength="250" wrap="soft"><%=sCapacity%></textarea>
							</p>
							<p>
								Supervisor: &nbsp; <% ShowRentalSupervisors iSupervisorUserId, "No Supervisor" %>
							</p>
							<p>
								How far out can the public reserve this rental<br />
								Residents: <input type="text" name="residentrentalperiod" id="residentrentalperiod" value="<%=sResidentRentalPeriod%>" size="3" maxlength="3" onchange="ValidateRentalPeriod( this );" /> Months &nbsp; 
								Nonresidents: <input type="text" name="nonresidentrentalperiod" id="nonresidentrentalperiod" value="<%=sNonResidentRentalPeriod%>" size="3" maxlength="3" onchange="ValidateRentalPeriod( this );" /> Months<br />
							</p>
							<p><br />
								Rental Notes (This will display on the receipt, once for each date reserved) (You can use simple HTML to format this):<br />
								<textarea id="receiptnotes" name="receiptnotes" maxlength="2000" wrap="soft"><%=sReceiptNotes%></textarea>
							</p>
							<p><br />
								Terms (The public must agree to this to complete their reservation) (You can use simple HTML to format this):<br />
								<textarea id="terms" name="terms" maxlength="4000" wrap="soft"><%=sTerms%></textarea>
								<br /><br />
							</p>
						</div>
						<div id="tab2"> <!-- Descriptions -->
							<p>Please provide both a long and short description for this rental. These are needed for display on the public side</p>

							<p>
								A Short Description for search results (You can use simple HTML to format this):<br />
								<textarea id="shortdescription" name="shortdescription" maxlength="500" wrap="soft"><%=sShortDescription%></textarea>
								<br />
								<br />
								A Longer Description for the public details page: &nbsp; 
								<input type="checkbox" name="chkUseHTMLonLong" id="chkUseHTMLonLong" <%=sCheckUseHTMLonLong%> /> Includes HTML
								<br />
								<textarea id="description" name="description" maxlength="6000" wrap="soft"><%=sDescription%></textarea>
								<br />
							</p>
						</div>
						<div id="tab3"> <!-- Images -->
							<p>
								These images will appear only for the public to see.<br /><br />
								<span class="mainimagetext">Image* to left of the description:</span><br />
								
								<input type="<% if blnHasWP then %>hidden<%else%>text<%end if%>" class="imageurl" id="iconimageurl" name="iconimageurl" size="100" maxlength="250" value="<%=sIconImageUrl%>" /> 
								<img src="<%=sIconImageUrl%>" id="iconimageurlpic" align="middle" width="240" height="180"  onerror="this.src = '../images/placeholder.png';" />
								<% if blnHasWP then %>
									<input type="button" class="button" value="Change" onclick="showModal('Pick Image',65,80,'iconimageurl');" />
								<% else %>
									<input type="button" class="button" value="Pick" onclick="doImagePicker('iconimageurl');" />
								<% End If %>

								<br />
								The Alt and Title Tags will be set to the rental name for this image.
							</p>
							<hr />
							<p><span class="mainimagetext">Extra Images* (maximum of 6 images)</span></p>
							<table><tr><td>
							<!--<tr><td>
								Place Images: &nbsp; <% 'ShowImagePlacementPicks sImagesPlacement %><br /><br />
							</td></tr>-->
							<tr><td>
								<%	If ExtraImageCount( iRentalId ) < CLng(6) Then 
										sDisabledAdd = ""
									Else
										sDisabledAdd = " disabled=""disabled"" "
									End If 
								%>
								<input id="addanimage" type="button" class="button" value="Add An Image" onclick="AddImageRow();" <%=sDisabledAdd%> /> &nbsp;&nbsp; 
								<input type="button" class="button" value="Remove Selected Images" onclick="RemoveImages();" /><br /><br />
							</td></tr>
							<tr><td>
								<table cellpadding="2" cellspacing="0" border="0" class="rentaltable" id="imagelist">
									<tr>
										<th>Remove</th><th>Image</th><th>Alt Tag</th><th>Order</th></tr>
<%										iMaxImages = ShowImageList( iRentalId )		%>		
								</table>
								<input type="hidden" id="maximages" name="maximages" value="<%=iMaxImages%>" />
								<div class="helpmsg">
									<strong>* Images should be 240px width by 180px height and should be less than 20KB.</strong>
								</div>
							</td></tr></table>
						</div>
						<div id="tab4"> <!-- Documents -->
							<p>
								These are any documents associated with this rental that you want the public to be able to download.
							</p>
							<table><tr><td><br />
								<input type="button" class="button" value="Add A Document" onclick="AddDocument();" /> &nbsp;&nbsp; 
								<input type="button" class="button" value="Remove Selected Documents" onclick="RemoveDocuments();" /><br /><br />
							</td></tr>
							<tr><td>
								<table cellpadding="2" cellspacing="0" border="0" class="rentaltable" id="documentlist">
									<tr>
										<th>Remove</th><th>Document</th><th>Name</th></tr>
<%										iMaxDocuments = ShowDocumentList( iRentalId )		%>		
								</table>
								<input type="hidden" id="maxdocuments" name="maxdocuments" value="<%=iMaxDocuments%>" />
								<div class="helpmsg">
									* Both the Document URL and Name are required.
								</div>
							</td></tr></table>
						</div>
						<div id="tab5"> <!-- In Season -->
<%								If iRentalId = CLng(0) Then		%>
									<p>
									The In Season Schedule cannot be created until the rental has been created.
									</p>
<%								Else	%>
									<p>
										This is the normal operation schedule for this rental. You can manage the rates (for resident, non-resident, or both), hours, minimum rental time, and many more features for each day of the week by clicking on the day.
									</p>
									<input type="button" class="button" value="Copy Schedule Days" onclick="location.href='rentaldaycopy.asp?rentalid=<%=iRentalId%>';" />
									<table id="inseasondays" class="seasondays" cellpadding="2" cellspacing="0" border="0">
										<tr><th>Day Of Week</th><th>Open</th><th>Available to Public</th><th>Hours</th></tr>
<%										ShowSchedule iRentalId, 0	%>
									</table>
<%								End If		%>
						</div>
						<div id="tab6"> <!-- Off Season -->
							<p>
<%								If iRentalId = CLng(0) Then		%>
									The Off Season Schedule cannot be created until the rental has been created.
<%								Else	%>
									<p>
										If the rental is shut down or has a different schedule for part of the year, that is set up here.
									</p>
									<table><tr><td><br />
										<input type="checkbox" id="hasoffseason" name="hasoffseason" <%=sHasOffSeason%> /> &nbsp This Rental Has An Off Season Schedule<br /><br />
										Off Season Starts: <% ShowMonthPicks "offseasonstartmonth", sOffSeasonStartMonth, "offseasonstart", "offseasonstartday" %> &nbsp; 
											<span id="offseasonstart">
												<% ShowDayPicks "offseasonstartday", sOffSeasonStartMonth, sOffSeasonStartDay  %>
											</span> &nbsp;  &nbsp; 
										In Season Resumes: <% ShowMonthPicks "offseasonendmonth", sOffSeasonEndMonth, "offseasonend", "offseasonendday" %>&nbsp;
											<span id="offseasonend">
												<% ShowDayPicks "offseasonendday", sOffSeasonEndMonth, sOffSeasonEndDay  %></span>&nbsp;<% ShowSameNextYearPick "offseasonendyear", sOffSeasonEndYear %>
									</td></tr>
									<tr><td><br />
										<input type="checkbox" id="reservationsduringseason" name="reservationsduringseason" <%=sReservationsDuringSeason%> /> &nbsp Reservations can only be made by the public for the current season.<br /><br />
										<input type="checkbox" id="nonresidentswait" name="nonresidentswait" <%=sNonResidentsWait%> /> &nbsp Additionally, nonresidents must wait an additional 
										<input type="text" id="nonresidentwaitdays" name="nonresidentwaitdays" value="<%=sNonresidentWaitDays%>" size="3" maxlength="3" onchange="ValidateNonresidentWait( this );" />
										days after the season starts, to make reservations. <%=sNonresidentStartDate%>
									</td></tr>
									<tr><td><br />
										<input type="button" class="button" value="Copy Schedule Days" onclick="location.href='rentaldaycopy.asp?rentalid=<%=iRentalId%>';" />
										<table id="offseasondays" class="seasondays" cellpadding="2" cellspacing="0" border="0">
											<tr><th>Day Of Week</th><th>Open</th><th>Available to Public</th><th>Hours</th></tr>
<%											ShowSchedule iRentalId, 1	%>
										</table>
									</td></tr></table>
<%								End If		%>

							</p>
						</div>

						<div id="tab7"> <!-- Items -->
							<p>
								Use this to set up any additional items that are available when this rental is reserved.
							</p>
							<table><tr><td><br />
								<input type="button" class="button" value="Add An Item" onclick="AddItemRow();" /> &nbsp;&nbsp; 
								<input type="button" class="button" value="Remove Selected Items" onclick="RemoveItems();" /><br /><br />
							</td></tr>
							<tr><td>
								<table cellpadding="2" cellspacing="0" border="0" class="rentaltable" id="rentalitemlist">
									<tr><th>Remove</th><th>Item</th><th>
<%										If bOrgHasAccounts Then 										
											response.write "Account"
										Else
											response.write "&nbsp;"
										End If 
%>
										</th><th>Maximum<br />Available</th><th>Rate Each<br />Per Day</th></tr>
<%									iMaxItems =	ShowItemList( iRentalId, bOrgHasAccounts )		%>									
								</table>
								<input type="hidden" id="maxitems" name="maxitems" value="<%=iMaxItems%>" />
								<div class="helpmsg">
									* All fields are required.
								</div>
							</td></tr></table>
						</div>
						<div id="tab8"> <!-- Fees -->
							<br />
							<p>
							These fees are charged once per reservation. To include them, check the box to the left of the fee.<br /><br />
							<table id="rentalfeetable" class="rentaltable" border="0" cellpadding="2" cellspacing="0">
								<tr><th>Type</th><th>Prompt<br />Question</th><th>
<%										If bOrgHasAccounts Then 										
											response.write "Account"
										Else
											response.write "&nbsp;"
										End If 
%>
										</th><th>Fee<br />Amount</th></tr>
<%								iMaxFeeRows = ShowRentalFees( iRentalId, bOrgHasAccounts )		%>
							</table>
							<input type="hidden" id="maxfeerows" name="maxfeerows" value="<%=iMaxFeeRows%>" /><br />
							</p>
						</div>
						<div id="tab9"> <!-- Alerts -->
							<p>
								Set alerts to notify rental supervisors when key events occur. To remove them, check the box and click the Remove Selected button.<br /><br />
								<input type="button" class="button" value="Add Row" id="addalertbutton" onClick="NewAlertRow()" /> &nbsp;&nbsp; 
								<input type="button" class="button" value="Remove Selected" id="removealertbutton" onClick="RemoveAlertRows()" />
								<table id="alerttable" border="0" cellpadding="0" cellspacing="0">
									<tr><th>&nbsp;</th><th>Alert Trigger</th><th>Notify</th></tr>
		<%								iMaxAlertRows = ShowAlertTable( iRentalId ) %>
								</table>
								<input type="hidden" id="maxalertrows" name="maxalertrows" value="<%=iMaxAlertRows%>" />
							</p>
						</div>
					</div>
				</div>

				<p>
					<input type="button" class="button" id="savebutton" value="<%=sButtonValue%>" onclick="validate();" />
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
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' integer iFeeCount = ShowRentalFees( iRentalId, bOrgHasAccounts )
'--------------------------------------------------------------------------------------------------
Function ShowRentalFees( ByVal iRentalId, ByVal bOrgHasAccounts )
	Dim oRs, sSql, iAccountNo, iRateTypeId, sRateAmount, iRowCount, sPrompt

	iRowCount = 0
	sSql = "SELECT pricetypeid, pricetypename, isoptional, needsprompt "
	sSql = sSql & " FROM egov_price_types WHERE isforrentals = 1 AND isrentalflatfee = 1 AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		sSelected = GetFeeInfo( iRentalId, oRs("pricetypeid"), iAccountNo, sRateAmount, sPrompt )

		response.write vbcrlf & "<tr>"

		' Show the Price Type
		response.write "<td class=""type""><input type=""checkbox"" id=""pricetypeid" & iRowCount & """ name=""pricetypeid"" value=""" & oRs("pricetypeid") & """" & sSelected & " /> &nbsp;"
		response.write oRs("pricetypename") 
		If oRs("isoptional") Then
			response.write " (optional)"
'		Else 
'			response.write " (required)"
		End If 
		response.write "</td>"

		response.write "<td align=""center"">"
		If oRs("needsprompt") Then
			response.write "<input type=""text"" id=""prompt" & oRs("pricetypeid") & """ name=""prompt" & oRs("pricetypeid") & """ value=""" & sPrompt & """ size=""30"" maxlength=""70"" />"
		Else
			response.write "&nbsp;<input type=""hidden"" id=""prompt" & oRs("pricetypeid") & """ name=""prompt" & oRs("pricetypeid") & """ value="""" />"
		End If 
		response.write "</td>"
		
		' Show the account
		response.write "<td align=""center"">"
		If bOrgHasAccounts Then 
			ShowAccountPicks "accountid" & oRs("pricetypeid"), iAccountNo, False 
		Else
			response.write "<input type=""hidden"" id=""accountid" & oRs("pricetypeid") & """ name=""accountid" & oRs("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "</td>"

		' Show the rate
		response.write "<td align=""center"">"
		response.write "<input type=""text"" id=""amount" & oRs("pricetypeid") & """ name=""amount" & oRs("pricetypeid") & """ value=""" & sRateAmount & """ size=""7"" maxlength=""7"" onchange=""ValidatePrice(this);"" />"
		response.write "</td>"

		response.write "</tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	ShowRentalFees = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' string sChecked = GetFeeInfo( iRentalId, iPriceTypeId, iAccountNo, sRateAmount, sPrompt )
'--------------------------------------------------------------------------------------------------
Function GetFeeInfo( ByVal iRentalId, ByVal iPriceTypeId, ByRef iAccountNo, ByRef sRateAmount, ByRef sPrompt )
	Dim oRs, sSql

	sSql = "SELECT accountid, ISNULL(amount,0) AS amount, ISNULL(prompt,'') AS prompt "
	sSql = sSql & " FROM egov_rentalfees "
	sSql = sSql & " WHERE rentalid = " & iRentalId & " AND pricetypeid = " & iPriceTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iAccountNo = oRs("accountid")
		If CDbl(oRs("amount")) > CDbl(0.00) Then 
			sRateAmount = FormatNumber(oRs("amount"),2,,,0)
		Else
			sRateAmount = ""
		End If 
		sPrompt = oRs("prompt")
		GetFeeInfo = " checked=""checked"" "
	Else
		iAccountNo = 0
		sRateAmount = ""
		sPrompt = ""
		GetFeeInfo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetRentalValues iRentalId 
'--------------------------------------------------------------------------------------------------
Sub GetRentalValues( ByVal iRentalId )
	Dim sSql, oRs, sMonth, sDay

	sSql = "SELECT rentalname, ISNULL(locationid,0) AS locationid, ISNULL(width,'') AS width, ISNULL(length,'') AS length, isdeactivated, "
	sSql = sSql & "ISNULL(capacity,'') AS capacity, imagesplacement, hasoffseason, offseasonstartmonth, "
	sSql = sSql & "offseasonstartday, offseasonendmonth, offseasonendday, ISNULL(offseasonendyear,0) AS offseasonendyear, publiccanview, publiccanreserve, "
	sSql = sSql & "needsapproval, nocosttorent, ISNULL(supervisoruserid,0) AS supervisoruserid, ISNULL(receiptnotes,'') AS receiptnotes, "
	sSql = sSql & "ISNULL(description,'') AS description, ISNULL(shortdescription,'') AS shortdescription, "
	sSql = sSql & "ISNULL(terms,'') AS terms, ISNULL(residentrentalperiod,0) AS residentrentalperiod, usehtmlonlongdesc, "
	sSql = sSql & "ISNULL(nonresidentrentalperiod,0) AS nonresidentrentalperiod, ISNULL(iconimageurl,'') AS iconimageurl, "
	sSql = sSql & "reservationsduringseason, nonresidentswait, ISNULL(nonresidentwaitdays,0) AS nonresidentwaitdays "
	sSql = sSql & "FROM egov_rentals WHERE orgid = " & session("orgid") & " AND rentalid = " & iRentalId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sRentalName = oRs("rentalname")
		sDescription = oRs("description")
		If oRs("publiccanview") Then 
			sPublicCanView = "checked=""checked"""
		End If 
		If oRs("publiccanreserve") Then 
			sPublicCanReserve = "checked=""checked"""
		End If 
		If oRs("needsapproval") Then 
			sNeedsApproval = "checked=""checked""" 
		End If 
		If oRs("nocosttorent") Then 
			sNoCostToRent = "checked=""checked""" 
		End If 
		If oRs("isdeactivated") Then
			sDeactivatedCheck = "checked=""checked""" 
		End If 
		sIconImageUrl = oRs("iconimageurl")
		iLocationid = oRs("locationid")
		sWidth = oRs("width")
		sLength = oRs("length")
		sCapacity = oRs("capacity")
		iSupervisorUserId = oRs("supervisoruserid")
		sReceiptNotes = oRs("receiptnotes")
		sImagesPlacement = oRs("imagesplacement")
		If oRs("hasoffseason") Then 
			sHasOffSeason = "checked=""checked""" 
		End If 
		sOffSeasonStartMonth = oRs("offseasonstartmonth")
		sOffSeasonStartDay = oRs("offseasonstartday")
		sOffSeasonEndMonth = oRs("offseasonendmonth")
		sOffSeasonEndDay = oRs("offseasonendday")
		sOffSeasonEndYear = oRs("offseasonendyear")
		sShortDescription = oRs("shortdescription")
		sTerms = oRs("terms")
		If clng(oRs("residentrentalperiod")) > clng(0) Then 
			sResidentRentalPeriod = oRs("residentrentalperiod")
		End If 
		If clng(oRs("nonresidentrentalperiod")) > clng(0) Then 
			sNonResidentRentalPeriod = oRs("nonresidentrentalperiod")
		End If
		If oRs("reservationsduringseason") Then 
			sReservationsDuringSeason = " checked=""checked"" "
		Else
			sReservationsDuringSeason = ""
		End If 
		If oRs("nonresidentswait") Then 
			sNonResidentsWait = " checked=""checked"" "
		Else
			sNonResidentsWait = ""
		End If 
		If clng(oRs("nonresidentwaitdays")) > clng(0) Then 
			sNonresidentWaitDays = oRs("nonresidentwaitdays")
			sMonth = Month(DateAdd("d", sNonresidentWaitDays, CDate(sOffSeasonEndMonth & "/" & sOffSeasonEndDay & "/" & Year(Date())) ))
			sDay = Day(DateAdd("d", sNonresidentWaitDays, CDate(sOffSeasonEndMonth & "/" & sOffSeasonEndDay & "/" & Year(Date())) ))
			sNonresidentStartDate = "(" & sMonth & "/" & sDay & " for the In Season)"
		Else
			sNonresidentWaitDays = ""
			sNonresidentStartDate = ""
		End If 
		If oRs("usehtmlonlongdesc") Then
			sCheckUseHTMLonLong = " checked=""checked"" "
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowCategoryPicks iRentalId 
'--------------------------------------------------------------------------------------------------
Sub ShowCategoryPicks( ByVal iRentalId )
	Dim sSql, oRs, iCount

	'iCount = 0
	sSql = "SELECT recreationcategoryid, categorytitle FROM egov_recreation_categories "
	sSql = sSql & "WHERE orgid = " & session("orgid")
	sSql = sSql & " AND isforrentals = 1 ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		'iCount = iCount + 1
		response.write "<input type=""checkbox"" name=""recreationcategoryid"" value=""" & oRs("recreationcategoryid") & """"
		If RentalHasCategory( iRentalId, oRs("recreationcategoryid") ) Then
			response.write " checked=""checked"""
		End If 
		response.write " />&nbsp;" & oRs("categorytitle") & " <br />"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasCategory( iRentalId, iCategoryId )
'--------------------------------------------------------------------------------------------------
Function RentalHasCategory( ByVal iRentalId, ByVal iCategoryId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(recreationcategoryid) AS hits FROM egov_rentals_to_categories "
	sSql = sSql & "WHERE rentalid = " & iRentalId & " AND recreationcategoryid = " & iCategoryId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			RentalHasCategory = True 
		Else
			RentalHasCategory = False
		End If 
	Else
		RentalHasCategory = False 
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer ExtraImageCount( iRentalId )
'--------------------------------------------------------------------------------------------------
Function ExtraImageCount( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(imageid) AS hits FROM egov_rentalimages WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		ExtraImageCount = CLng(oRs("hits")) 
	Else
		ExtraImageCount = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowImageList( iRentalId )
'--------------------------------------------------------------------------------------------------
Function ShowImageList( ByVal iRentalId )
	Dim sSql, oRs, iImageCount, iMaxDisplay, x

	iImageCount = 0
	iMaxDisplay = GetMaxImageDisplayCount( iRentalId )
	sSql = "SELECT imageid, imageurl, alttag FROM egov_rentalimages "
	sSql = sSql & "WHERE rentalid = " & iRentalId & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			iImageCount = iImageCount + 1
			response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removeimage" & iImageCount & """ name=""removeimage" & iImageCount & """ /></td>"
			response.write "<td align=""left""><input type="""
			if blnHasWP then
				response.write "hidden"
			else
				response.write "text"
			end if
			response.write """ id=""imageurl" & iImageCount & """ name=""imageurl" & iImageCount & """ size=""60"" maxlength=""250"" value=""" & oRs("imageurl") & """ />"
			response.write "<img src=""" & oRs("imageurl") & """ id=""imageurl" & iImageCount & "pic"" align=""middle"" width=""240"" height=""180""  onerror=""this.src = '../images/placeholder.png';"" />"
			if blnHasWP then
				response.write " <input type=""button"" class=""button"" value=""Change"" onclick=""showModal('Pick Image', 65, 80, 'imageurl" & iImageCount & "');"" />"
			else
				response.write " <input type=""button"" class=""button"" value=""Change"" onclick=""doImagePicker('imageurl" & iImageCount & "');"" />"
			end if
			response.write "</td>"
			response.write "<td align=""center""><input type=""text"" id=""alttag" & iImageCount & """ name=""alttag" & iImageCount & """ size=""30"" maxlength=""100"" value=""" & oRs("alttag") & """ /></td>"
			response.write "<td align=""center""><select id=""displayorder" & iImageCount & """ name=""displayorder" & iImageCount & """>"
			For x = 1 To iMaxDisplay
				response.write vbcrlf & "<option value=""" & x & """"
				If CLng(x) = CLng(iImageCount) Then 
					response.write " selected=""selected"" "
				End If 
				response.write ">" & x & "</option>"
			Next 
			response.write vbcrlf & "</select>"
			response.write "</td></tr>"
			oRs.MoveNext 
		Loop 
	Else 
		' write out a blank initial row
		response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removeimage1"" name=""removeimage1"" /></td>"
		response.write "<td align=""center""><input type="""
		if blnHasWP then
			response.write "hidden"
		else
			response.write "text"
		end if
		response.write """ id=""imageurl1"" name=""imageurl1"" size=""60"" maxlength=""250"" value="""" />"
		response.write "<img src=""asdfadfs"" id=""imageurl1pic"" width=""240"" height=""180""  onerror=""this.src = '../images/placeholder.png';"" align=""middle"" />"
		if blnHasWP then
			response.write "<input type=""button"" class=""button"" value=""Pick"" onclick=""showModal('Pick Image', 65, 80, 'imageurl1');"" /></td>"
		else
			response.write "<input type=""button"" class=""button"" value=""Pick"" onclick=""doImagePicker('imageurl1');"" /></td>"
		end if
		response.write "<td align=""center""><input type=""text"" id=""alttag1"" name=""alttag1"" size=""30"" maxlength=""100"" value="""" /></td>"
		response.write "<td align=""center""><select id=""displayorder1"" name=""displayorder1"">"
		response.write vbcrlf & "<option value=""1"">1</option>"
		response.write vbcrlf & "</select>"
		response.write "</td></tr>"
		iImageCount = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowImageList = iImageCount
End Function 


'--------------------------------------------------------------------------------------------------
' integer GetMaxImageDisplayCount( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetMaxImageDisplayCount( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(imageid) AS hits FROM egov_rentalimages "
	sSql = sSql & "WHERE rentalid = " & iRentalId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMaxImageDisplayCount = CLng(oRs("hits"))
	Else
		GetMaxImageDisplayCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' void ShowImagePlacementPicks sImagesPlacement 
'--------------------------------------------------------------------------------------------------
Sub ShowImagePlacementPicks( ByVal sImagesPlacement )
	
	response.write vbcrlf & "<select id=""imagesplacement"" name=""imagesplacement"">"
	response.write vbcrlf & "<option value=""none""" 
	If sImagesPlacement = "none" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">No images displayed</option>"
	response.write vbcrlf & "<option value=""bottom"""
	If sImagesPlacement = "bottom" Then 
		response.write " selected=""selected"" "
	End If
	response.write ">Place images along the bottom of the page</option>"
	response.write vbcrlf & "<option value=""right"""
	If sImagesPlacement = "right" Then 
		response.write " selected=""selected"" "
	End If
	response.write ">Place images along the right of the page</option>"
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowDocumentList( iRentalId )
'--------------------------------------------------------------------------------------------------
Function ShowDocumentList( ByVal iRentalId )
	Dim sSql, oRs, iDocCount, iMaxDisplay, x

	iDocCount = 0
	'iMaxDisplay = GetMaxDocumentDisplayCount( iRentalId )
	sSql = "SELECT documentid, documenturl, documenttitle FROM egov_rentaldocuments "
	sSql = sSql & "WHERE rentalid = " & iRentalId & " ORDER BY documenttitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			iDocCount = iDocCount + 1
			response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removedocument" & iDocCount & """ name=""removedocument" & iDocCount & """ /></td>"
			if blnHasWP then
				response.write "<td align=""center""><input type=""hidden"" id=""documenturl" & iDocCount & """ name=""documenturl" & iDocCount & """ size=""60"" maxlength=""250"" value=""" & oRs("documenturl") & """ />"
				response.write "<span id=""documenturl" & iDocCount & "pic""><a href=""" & oRs("documenturl") & """ target=""_newwindow"">View Document</a>&nbsp;&nbsp;</span>"
				response.write "<input type=""button"" class=""button"" value=""Pick"" onclick=""showModal('Pick File', 65, 80, 'documenturl" & iDocCount & "');"" /></td>"
			else
				response.write "<td align=""center""><input type=""text"" id=""documenturl" & iDocCount & """ name=""documenturl" & iDocCount & """ size=""60"" maxlength=""250"" value=""" & oRs("documenturl") & """ />"
				response.write "<span id=""documenturl" & iDocCount & "pic""><a href=""" & oRs("documenturl") & """ target=""_newwindow"">View Document</a>&nbsp;&nbsp;</span>"
				response.write "<input type=""button"" class=""button"" value=""Pick"" onclick=""doDocumentPicker('documenturl" & iDocCount & "');"" /></td>"
			end if
			response.write "<td align=""center""><input type=""text"" id=""documenttitle" & iDocCount & """ name=""documenttitle" & iDocCount & """ size=""40"" maxlength=""250"" value=""" & oRs("documenttitle") & """ /></td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else 
		' write out a blank initial row
		response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removedocument1"" name=""removedocument1"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""documenturl1"" name=""documenturl1"" size=""60"" maxlength=""250"" value="""" />"
		response.write "<input type=""button"" class=""button"" value=""Pick"" onclick=""doDocumentPicker('documenturl1');"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""documenttitle1"" name=""documenttitle1"" size=""40"" maxlength=""250"" value="""" /></td>"
		response.write "</tr>"
		iDocCount = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowDocumentList = iDocCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetMaxDocumentDisplayCount( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetMaxDocumentDisplayCount( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(documentid) AS hits FROM egov_rentaldocuments "
	sSql = sSql & "WHERE rentalid = " & iRentalId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMaxDocumentDisplayCount = CLng(oRs("hits"))
	Else
		GetMaxDocumentDisplayCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowItemList(  iRentalId, bOrgHasAccounts )
'--------------------------------------------------------------------------------------------------
Function ShowItemList( ByVal iRentalId, ByVal bOrgHasAccounts )
	Dim sSql, oRs, iCount, sRateAmount, sMaxAvailable

	iCount = 0

	sSql = "SELECT rentalitem, accountid, maxavailable, amount FROM egov_rentalitems "
	sSql = sSql & "WHERE rentalid = " & iRentalId & " ORDER BY rentalitem"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			iCount = iCount + 1
			response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removeitem" & iCount & """ name=""removeitem" & iCount & """ /></td>"
			response.write "<td align=""center""><input type=""text"" id=""rentalitem" & iCount & """ name=""rentalitem" & iCount & """ value=""" & oRs("rentalitem") & """ size=""30"" maxlength=""30"" /></td>"
			
			response.write "<td align=""center"">"
			If bOrgHasAccounts Then 
				ShowAccountPicks "itemaccountid" & iCount, oRs("accountid"), False 
			Else
				response.write "<input type=""hidden"" id=""itemaccountid" & iCount & """ name=""itemaccountid" & iCount & """ value=""0"" />"
			End If 
			response.write "</td>"

			If clng(oRs("maxavailable")) > clng(0) Then 
				sMaxAvailable = oRs("maxavailable")
			Else
				sMaxAvailable = ""
			End If 
			response.write "<td align=""center""><input type=""text"" id=""maxavailable" & iCount & """ name=""maxavailable" & iCount & """ value=""" & sMaxAvailable & """ size=""5"" maxlength=""5"" onchange=""ValidateMaxAvailable(this);"" /></td>"
			'If CDbl(oRs("amount")) > CDbl(0.00) Then 
				sRateAmount = FormatNumber(oRs("amount"),2,,,0)
			'Else
			'	sRateAmount = ""
			'End If 
			response.write "<td align=""center""><input type=""text"" id=""amount" & iCount & """ name=""amount" & iCount & """ value=""" & sRateAmount & """ size=""7"" maxlength=""7"" onchange=""ValidatePrice(this);"" /></td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else 
		' write out a blank initial row
		response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removeitem1"" name=""removeitem1"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""rentalitem1"" name=""rentalitem1"" value="""" size=""50"" maxlength=""50"" /></td>"
		response.write "<td align=""center"">"
		If bOrgHasAccounts Then
			ShowAccountPicks "itemaccountid1", 0, False 
		Else
			response.write "<input type=""hidden"" id=""itemaccountid1"" name=""itemaccountid1"" value=""0"" />"
		End If 
		response.write "</td>"
		response.write "<td align=""center""><input type=""text"" id=""maxavailable1"" name=""maxavailable1"" value="""" size=""5"" maxlength=""5"" onchange=""ValidateMaxAvailable(this);"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""amount1"" name=""amount1"" value="""" size=""7"" maxlength=""7"" onchange=""ValidatePrice(this);"" /></td>"
		response.write "</tr>"
		iCount = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowItemList = iCount

End Function 


'--------------------------------------------------------------------------------------------------
' integer ShowRentalList( iRentalId )
'--------------------------------------------------------------------------------------------------
Function ShowRentalList( ByVal iRentalId )
	Dim sSql, oRs, iRentalCount, iMaxDisplay, x

	iRentalCount = 0

	sSql = "SELECT R.rentalid, R.rentalname FROM egov_rentals_to_rentals RR, egov_rentals R "
	sSql = sSql & "WHERE RR.rentalid = " & iRentalId & " AND RR.associatedrentalid = R.rentalid ORDER BY R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			iRentalCount = iRentalCount + 1
			response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removerental" & iRentalCount & """ name=""removerental" & iRentalCount & """ /></td>"
			response.write "<td align=""center"">"
			ShowRentalSelections oRs("rentalid"), iRentalCount
			response.write "</td></tr>"
			oRs.MoveNext 
		Loop 
	Else 
		' write out a blank initial row
		response.write vbcrlf & "<tr><td align=""center""><input type=""checkbox"" id=""removerental1"" name=""removerental1"" /></td>"
		response.write "<td align=""center"">"
		ShowRentalSelections 0, 1
		response.write "</td></tr>"
		iRentalCount = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowRentalList = iRentalCount

End Function


'--------------------------------------------------------------------------------------------------
' void ShowRentalSelections iRentalId, iRowNumber 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalSelections( ByVal iRentalId, ByVal iRowNumber )
	Dim sSql, oRs

	sSql = "SELECT rentalid, rentalname FROM egov_rentals "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " ORDER BY R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select id=""associatedrentalid" & iRowNumber & """ name=""associatedrentalid" & iRowNumber & """>"
	response.write vbcrlf & "<option value=""0"">Select a related rental from the list.</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("rentalid") & """"
		If CLng(iRentalId) = CLng(oRs("rentalid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("rentalname") & "</option>"
		oRs.MoveNext
	Loop 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowMonthPicks sMonthPickName, iMonth, sDaySpanId, sDayPickName 
'--------------------------------------------------------------------------------------------------
Sub ShowMonthPicks( sMonthPickName, iMonth, sDaySpanId, sDayPickName )
	Dim x

	response.write "<select id=""" & sMonthPickName & """ name=""" & sMonthPickName & """ onchange=""getDays( '" & sMonthPickName & "', '" & sDaySpanId & "', '" & sDayPickName & "' );"">"
	For x = 1 To 12
		response.write vbcrlf & "<option value=""" & x & """"
		If clng(iMonth) = clng(x) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & MonthName(x) & "</option>"
	Next 
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowDayPicks sDayPickName, iMonth, iDay 
'--------------------------------------------------------------------------------------------------
Sub ShowDayPicks( sDayPickName, iMonth, iDay )
	Dim x, dStartDate, iEnd, dEndDate

	' We will try using the 2009 dates since this is not a leap year, or any other funny thing
	dStartDate = CDate( iMonth & "/1/2009")
	dEndDate = DateAdd("m", 1, dStartDate)
	dEndDate = DateAdd("d", -1, dEndDate)
	iEnd = Day(dEndDate)

	response.write "<select id=""" & sDayPickName & """ name=""" & sDayPickName & """>"
	For x = 1 To iEnd
		response.write vbcrlf & "<option value=""" & x & """"
		If clng(iDay) = clng(x) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & x & "</option>"
	Next
	response.write vbcrlf & "</select>"

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowSchedule iRentalId, iIsOffSeason 
'--------------------------------------------------------------------------------------------------
Sub ShowSchedule( ByVal iRentalId, ByVal iIsOffSeason )
	Dim sSql, oRs, iRowCount, sPrefix

	If iIsOffSeason = 1 Then
		sPrefix = "offseason"
	Else
		sPrefix = "inseason"
	End If 

	iRowCount = 0
	sSql = "SELECT dayid, weekdayname, isopen, isavailabletopublic, openinghour, openingminute, openingampm, "
	sSql = sSql & "closinghour, closingminute, closingampm FROM egov_rentaldays "
	sSql = sSql & "WHERE rentalid = " & iRentalId & " AND isoffseason = " & iIsOffSeason & " ORDER BY dayofweek"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr id=""" & iRowCount & """"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

'		response.write "<td align=""center"">"
'		response.write "<input type=""hidden"" name=""" & sPrefix & "dayid"" value=""" & oRs("dayid") & """ />"
'		response.write "<input type=""checkbox"" name=""" & sPrefix & "dayid"" value=""" & oRs("dayid") & """"
'		If oRs("isopen") Then 
'			response.write " checked=""checked"" "
'		End If 
'		response.write " /></td>"
'		response.write "<td align=""center"">"
'		response.write "<input type=""checkbox"" name=""" & sPrefix & "dayid"" value=""" & oRs("dayid") & """"
'		If oRs("isavailabletopublic") Then 
'			response.write " checked=""checked"" "
'		End If 
'		response.write " /></td>"

		response.write "<td align=""left"" class=""rentaldowname"" title=""click to edit"" onClick=""location.href='rentaldayedit.asp?dayid=" & oRs("dayid") & "';""><strong>" & oRs("weekdayname") & "</strong></td>"

		response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='rentaldayedit.asp?dayid=" & oRs("dayid") & "';"">"
		If oRs("isopen") Then
			response.write "YES"
		Else
			response.write "NO"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='rentaldayedit.asp?dayid=" & oRs("dayid") & "';"">"
		If oRs("isavailabletopublic") Then
			response.write "YES"
		Else
			response.write "NO"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='rentaldayedit.asp?dayid=" & oRs("dayid") & "';"">"
		If oRs("isopen") Then 
			If oRs("openinghour") <> "" Then
				response.write oRs("openinghour") & ":"
				If clng(oRs("openingminute")) < clng(10) Then 
					response.write "0"
				End If 
				response.write oRs("openingminute") & "&nbsp;" & UCase(oRs("openingampm")) & " &ndash; "
				response.write oRs("closinghour") & ":"
				If clng(oRs("closingminute")) < clng(10) Then 
					response.write "0"
				End If 
				response.write oRs("closingminute") & "&nbsp;" & UCase(oRs("closingampm"))
			Else
				response.write "No Hours"
			End If 
		Else
			response.write "Closed"
		End If 
		response.write "</td>"
		response.write vbcrlf & "</tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' integer ShowAlertTable( iRentalId )
'--------------------------------------------------------------------------------------------------
Function ShowAlertTable( ByVal iRentalId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT rentalalerttypeid, userid FROM egov_rentalalerts WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then 
				sRowClass = ""
			Else
				sRowClass = " class=""altrow"" "
			End If 
			response.write vbcrlf & "<tr" & sRowClass & "><td class=""firstcell"">"
			response.write "<input type=""checkbox"" id=""removealert" & iRowCount & """ name=""removealert" & iRowCount & """ />"
			response.write "</td>"
			response.write "<td align=""center"">"
			ShowRentalAlertTypePicks oRs("rentalalerttypeid"), iRowCount
			response.write "</td>"
			response.write "<td align=""center"">"
			ShowUserPicks oRs("userid"), iRowCount
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		' put in a starter row.
		iRowCount = 1
		response.write vbcrlf & "<tr><td class=""firstcell"">"
		response.write "<input type=""checkbox"" id=""removealert" & iRowCount & """ name=""removealert" & iRowCount & """ /></td>"
		response.write "<td align=""center"">"
		ShowRentalAlertTypePicks 0, iRowCount
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowUserPicks 0, iRowCount
		response.write "</td>"
		response.write "</tr>"
	End If 
	
	oRs.Close 
	Set oRs = Nothing 

	ShowAlertTable = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowRentalAlertTypePicks iRentalAlertTypeIdiRentalAlertTypeId, iRowCount 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalAlertTypePicks( ByVal iRentalAlertTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSQL = "SELECT rentalalerttypeid, rentalalerttype FROM egov_rentalalerttypes "
	sSql = sSql & " ORDER BY rentalalerttype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select id=""rentalalerttypeid" & iRowCount & """ name=""rentalalerttypeid" & iRowCount & """>"
		'response.write vbcrfl & "<option value=""0"">Select an Alert Type</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("rentalalerttypeid") & """ "  
			If CLng(iRentalAlertTypeId) = CLng(oRs("rentalalerttypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("rentalalerttype")
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' void ShowUserPicks iNotifyUserId, iRowCount 
'--------------------------------------------------------------------------------------------------
Sub ShowUserPicks( ByVal iNotifyUserId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users "
	sSql = sSql & "WHERE isrentalsupervisor = 1 AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY lastname, firstname"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	response.write vbcrlf & "<select id=""userid" & iRowCount & """ name=""userid" & iRowCount & """>"
	response.write vbcrlf & "<option value=""0"">No One Is Alerted</option>"

	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """ "  
		If CLng(iNotifyUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("firstname") & " " & oRs("lastname")
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


%>



