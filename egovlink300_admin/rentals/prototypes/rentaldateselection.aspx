<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Rental Date Selection</title>
    <style type="text/css">
        body {
            font-family:Verdana;
            font-size: 10px;
            }
            
        #searchdiv {
            width: 800px;
            }
            
        table.resultstable {
            width: 822px;
            border: 1px solid #336699;
            }
        
        table.resultstable th {
            height: 26px;
            font-size: 10px;
            color: #003366;
            border-bottom: 1px solid #336699;
            border-right: 1px solid #336699;
            background-color: #93bee1;
	        }
	        
	    table.resultstable td {
	        cursor: pointer;
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
        
    </style>
    <script language="javascript">
        <!--
        
        function Continue()
        {
            location.href='reservationedit.aspx';
        }
        
        //-->
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <p>
		<font size="+1"><strong>Rental Date Selection</strong></font><br />
	</p>
	<p>
        <b>Reservation Type:</b> Reservation
    </p>
	<p>
        <b>Reservation For:</b> Steve Loar
    </p>
    
    <div>
        <b>Location:</b> Adobe Picnic Shelter, Swaim Park
    </div><br /><br />
    
    <p><font size="+1"><b>Edit Dates and Times</b></font><br /><br />
        <p>
            <input type="button" value="Add A Date" class="button" /> &nbsp; <input type="button" value="Remove Selected Dates" class="button" />
        </p>
        <table border="0" cellpadding="2" cellspacing="0" class="resultstable">
            <tr><th>Date</th><th>Start Time</th><th>End Time</th><th>Available</th></tr>
            <tr><td>
                <input type="checkbox" /> &nbsp;
                <input type="text" value="10/01/2009" style="width: 90px" /><br /><br />
                    
                </td>
                <td align="center"><input type="text" value="7" size="2" style="width: 24px" />:
                    <input type="text" value="00" size="2" style="width: 24px" />&nbsp;
                    <select>
                        <option value="AM">AM</option>
                        <option value="PM" selected="selected">PM</option>
                    </select>
                </td>
                <td align="center"><input type="text" value="12" size="2" style="width: 24px" />:
                    <input type="text" value="00" size="2" style="width: 24px" />&nbsp;
                    <select>
                        <option value="AM" selected="selected">AM</option>
                        <option value="PM">PM</option>
                    </select>
                </td>
                <td align="center"><b>YES</b></td>
            </tr>
            <tr>
                <td colspan="3" valign="top">Open: 8:00 AM to 1:00AM The Next Day<br />
                    Last Reservation at: 7:00 PM That Day<br />
                    Minium Rental Time: 3 Hours
                </td>
                <td valign="top">Other Reservations<br />8:00 AM - 9:00AM<br />1:00PM - 3:00PM</td>
            </tr>
        </table><br /><br />
        <table border="0" cellpadding="2" cellspacing="0" class="resultstable">
            <tr><th>Date</th><th>Start Time</th><th>End Time</th><th>Available</th></tr>
            <tr><td>
                <input type="checkbox" /> &nbsp;
                <input type="text" value="10/02/2009" style="width: 90px" /></td>
                <td align="center"><input type="text" value="7" size="2" style="width: 24px" />:
                    <input type="text" value="00" size="2" style="width: 24px" />&nbsp;
                    <select>
                        <option value="AM">AM</option>
                        <option value="PM" selected="selected">PM</option>
                    </select>
                </td>
                <td align="center"><input type="text" value="12" size="2" style="width: 24px" />:
                    <input type="text" value="00" size="2" style="width: 24px" />&nbsp;
                    <select>
                        <option value="AM" selected="selected">AM</option>
                        <option value="PM">PM</option>
                    </select>
                </td>
                <td align="center"><b>YES</b></td>
            </tr>
            <tr>
                <td colspan="3" valign="top">Open: 8:00 AM to 1:00AM The Next Day<br />
                    Last Reservation at: 7:00 PM That Day<br />
                    Minium Rental Time: 3 Hours
                </td>
                <td valign="top">Other Reservations<br />8:00 AM - 9:00AM<br />1:00PM - 3:00PM</td>
            </tr>
        </table><br /><br />
        <table border="0" cellpadding="2" cellspacing="0" class="resultstable">
            <tr><th>Date</th><th>Start Time</th><th>End Time</th><th>Available</th></tr>
            <tr><td>
                <input type="checkbox" /> &nbsp;
                <input type="text" value="10/03/2009" style="width: 90px" /></td>
                <td align="center"><input type="text" value="7" size="2" style="width: 24px" />:
                    <input type="text" value="00" size="2" style="width: 24px" />&nbsp;
                    <select>
                        <option value="AM">AM</option>
                        <option value="PM" selected="selected">PM</option>
                    </select>
                </td>
                <td align="center"><input type="text" value="12" size="2" style="width: 24px" />:
                    <input type="text" value="00" size="2" style="width: 24px" />&nbsp;
                    <select>
                        <option value="AM" selected="selected">AM</option>
                        <option value="PM">PM</option>
                    </select>
                </td>
                <td align="center"><b>YES</b></td>
            </tr>
            <tr>
                <td colspan="3" valign="top">Open: 8:00 AM to 1:00AM The Next Day<br />
                    Last Reservation at: 7:00 PM That Day<br />
                    Minium Rental Time: 3 Hours
                </td>
                <td valign="top">Other Reservations<br />8:00 AM - 9:00AM<br />1:00PM - 3:00PM</td>
            </tr>
        </table><br /><br />
    </p>
    <p><input type="button" value="Check and Continue" onclick="Continue();" class="button" /></p>
    </form>
</body>
</html>
