<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Reservation Edit</title>
    
    <style type="text/css">
        div#content {
            margin-left: 2em;
           }   
           
        div.tab {
            border: 1px solid black;
            margin-top: 2em;
            width: 850px;
            padding-bottom: 2em;
            padding-left: 1em;
            padding-right: 1em;
            }
        
        table.resultstable {
            width: 822px;
            border: 1px solid #336699;
            }
        
        table.resultstable th {
            height: 26px;
            color: #003366;
            border-bottom: 1px solid #336699;
            border-right: 1px solid #336699;
            background-color: #93bee1;
	        }
	            
        table.resultstable td {
	        height: 20px;
	        }
	        
	    tr.altrow {
	        background-color: #ececec;
	        }
	        
	    tr.first td {
	        border-top: 1px solid #336699;
	        }
	        
	    input.button {
	        cursor: pointer;
	        }
	        
	    td.keyrow {
	        border-top: 1px solid black;
	        }
    </style>             
</head>
<body>
    <form id="form1" runat="server">
    <p>
		<font size="+1"><strong>Edit Reservation</strong></font><br />
	</p>
    <p>
        <b>Rental:</b> Adobe Picnic Shelter, Swaim Park
    </p>
    <p>
        <b>Reservation For:</b> Steve Loar
    </p>
    <p>
        <b>Status:</b> Reserved
    </p>
    <p>
        <input id="Button7" type="button" value="Save Changes" />&nbsp;&nbsp;
        <input id="Button8" type="button" value="Cancel This Reservation" />
    </p>
    <div class="tab">
        <b>Venue Info. Tab</b><br /><br />
        <p>
            Dimensions: 40ft X 35ft<br />
            Capacity: Tables for 50 but area can hold 100<br />
            Description: This is a picnic area with a shelter (40 X 35) it has 4 picnic tables and a charcoal grill. Capacity is limited to 100.<br />
        </p>
        <p>
            Documents:<br />
            <a href="#">Medical Release Form</a><br />
            <a href="#">General Release Form</a><br />
        </p>
    </div>
    <div class="tab">
        <b>Event Info. Tab</b><br /><br />
        <p>
            Organization: <input type="text" size="40" />
        </p>
        <p>
            Point of Contact: <input type="text" size="40" />
        </p>
        <p>
            Number Attending: <input type="text" size="5" />
        </p>
        <p>
            Purpose: <input type="text" size="50" />
        </p>
        <p>
            Invoice/Receipt Notes:<br />
            <textarea cols="100" rows="6"></textarea>
        </p>
        <p>
            Private Notes:<br />
            <textarea cols="100" rows="6"></textarea>
        </p>
    </div>
    
    <div class="tab">
        <b>Rates & Fees Tab</b><br /><br />
        <p>
            <input id="Button5" type="button" value="Add A Date" />&nbsp;&nbsp;
            <input id="Button6" type="button" value="Cancel Selected Dates" />&nbsp;&nbsp;
            <input id="Button10" type="button" value="Select All Dates" />&nbsp;&nbsp;
            <input id="Button11" type="button" value="Deselect All Dates" />
        </p>
        <table border="0" cellpadding="2" cellspacing="0" class="resultstable">
            <tr><th>Date</th><th>Status</th><th>Fees</th></tr>
            <tr><td align="left" valign="top"><input id="Checkbox22" type="checkbox" />&nbsp;
                    <b>10/01/2009</b><br />
                    7:00PM to 12:00AM (5 hours)
                    <br /><br />
                    Arrival:&nbsp;<input type="text" value="7:00PM" size="15" />
                    <br />
                    Departure:&nbsp;<input type="text" value="12:00AM" size="15" />
                    <br /><br />
                    Open: 8:00 AM to 1:00AM The Next Day<br />
                    Last Reservation at: 7:00 PM That Day<br />
                    Minium Rental Time: 3 Hours
                </td>
                <td align="center" valign="top">Reserved</td>
                <td align="right">
                    <table cellpadding="2" cellspacing="0" border="0" class="none">
                    <tr><td colspan="3" align="left"><b>Rates</b></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Resident (20.00 per Hour)</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Nonresident (5.00 per Hour)</td><td align="center"><input type="text" value="25.00" size="5" /></td></tr>
                    <tr><td colspan="3" align="left"><b>Items</b></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Chairs (5.00 each, Max 25)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Large Round Table (25.00 each, Max 5)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Small Cocktail Tables (18.00 each, Max 3)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Tablecloths (5.75 each, Max 30)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td colspan="2" align="right"><b>Subtotal</b></td><td align="center"><b>125.00</b></td></tr>
                    </table>
                </td>
            </tr>
            <tr class="altrow"><td align="left" valign="top"><input id="Checkbox1" type="checkbox" />&nbsp;
                    <b>10/02/2009</b><br />
                    7:00PM to 12:00AM (5 hours)
                    <br /><br />
                    Arrival:&nbsp;<input type="text" value="7:00PM" size="15" />
                    <br />
                    Departure:&nbsp;<input type="text" value="12:00AM" size="15" />
                    <br /><br />
                    Open: 8:00 AM to 1:00AM The Next Day<br />
                    Last Reservation at: 7:00 PM That Day<br />
                    Minium Rental Time: 3 Hours
                </td>
                <td align="center" valign="top">Reserved</td>
                <td align="right">
                    <table cellpadding="2" cellspacing="0" border="0" class="none">
                    <tr><td colspan="3" align="left"><b>Rates</b></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Resident (20.00 per Hour)</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Nonresident (5.00 per Hour)</td><td align="center"><input type="text" value="25.00" size="5" /></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Weekend Surcharge (3.00 per Hour from 7:00 PM)</td><td align="center"><input type="text" value="15.00" size="5" /></td></tr>
                    <tr><td colspan="3" align="left"><b>Items</b></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Chairs (5.00 each, Max 25)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Large Round Table (25.00 each, Max 5)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Small Cocktail Tables (18.00 each, Max 3)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Tablecloths (5.75 each, Max 30)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td colspan="2" align="right"><b>Subtotal</b></td><td align="center"><b>140.00</b></td></tr>
                    </table>
                </td>
            </tr>
            <tr><td align="left" valign="top"><input id="Checkbox2" type="checkbox" />&nbsp;
                    <b>10/03/2009</b><br />
                    7:00PM to 12:00AM (5 hours)
                    <br /><br />
                    Arrival:&nbsp;<input type="text" value="7:00PM" size="15" />
                    <br />
                    Departure:&nbsp;<input type="text" value="12:00AM" size="15" />
                    <br /><br />
                    Open: 8:00 AM to 1:00AM The Next Day<br />
                    Last Reservation at: 7:00 PM That Day<br />
                    Minium Rental Time: 3 Hours
                </td>
                <td align="center" valign="top">Reserved</td>
                <td align="right">
                    <table cellpadding="2" cellspacing="0" border="0" class="none">
                    <tr><td colspan="3" align="left"><b>Rates</b></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Resident (20.00 per Hour)</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Nonresident (5.00 per Hour)</td><td align="center"><input type="text" value="25.00" size="5" /></td></tr>
                    <tr><td align="center"><input type="checkbox" checked="checked" /></td><td>Weekend Surcharge (3.00 per Hour from 7:00 PM)</td><td align="center"><input type="text" value="15.00" size="5" /></td></tr>
                    <tr><td colspan="3" align="left"><b>Items</b></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Chairs (5.00 each, Max 25)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Large Round Table (25.00 each, Max 5)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Small Cocktail Tables (18.00 each, Max 3)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td align="center"><input type="text" value="0" size="5" /></td><td>Tablecloths (5.75 each, Max 30)</td><td align="center"><input type="text" value="" size="5" /></td></tr>
                    <tr><td colspan="2" align="right"><b>Subtotal</b></td><td align="center"><b>140.00</b></td></tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="keyrow" colspan="2">&nbsp;</td>
                <td class="keyrow" align="right">
                    <input type="checkbox" checked="checked" /> &nbsp; <b>Deposit Fee</b> (35.00)&nbsp;
                    <input type="text" value="35.00" size="5" />
                </td>
            </tr>
            <tr>
                <td colspan="2">&nbsp;</td>
                <td align="right">
                    <input type="checkbox" /> &nbsp; <b>Alcohol Surcharge</b> (50.00)&nbsp;
                    <input type="text" value="" size="5" />
                </td>
            </tr>
            <tr class="altrow">
                <td class="keyrow" colspan="2"><b>Total</b></td>
                <td class="keyrow" align="right"><b>440.00</b></td>
            </tr>
            
        </table>
    </div>
    <div class="tab">
        <b>Invoices & Receipts Tab</b><br /><br />
        <table class="resultstable" cellpadding="0" cellspacing="0" border="0">
            <tr><th>Total Fees</th><th>Non-Invoiced Fees</th><th>Invoiced Fees</th><th>Total Paid</th><th>Total Due</th></tr>
            <tr>
                <td align="center">440.00</td>
                <td align="center">405.00</td>
                <td align="center">35.00</td>
                <td align="center">0.00</td>
                <td align="center">440.00</td>
            </tr>
        </table><br />
        <input id="Button2" type="button" value="New Invoice" />&nbsp;&nbsp;
        <input id="Button1" type="button" value="Pay Invoices" />&nbsp;&nbsp;
        <input id="Button4" type="button" value="Void Invoices" />&nbsp;&nbsp;
        <input id="Button9" type="button" value="Refund Payments" />&nbsp;&nbsp;
        <input id="Button3" type="button" value="View Combined Invoice/Receipt" /><br /><br />
        
        <table class="resultstable" cellpadding="0" cellspacing="0" border="0">
            <tr><th>Invoice</th><th>Date</th><th>Invoice Total</th><th>Amount Paid</th><th>Status</th></tr>
            <tr>
                <td align="center">4795</td>
                <td align="center">10/01/2009</td>
                <td align="center">35.00</td>
                <td align="center">0.00</td>
                <td align="center">Due</td>
            </tr>
        </table>
        
    </div>
    </form>
</body>
</html>
