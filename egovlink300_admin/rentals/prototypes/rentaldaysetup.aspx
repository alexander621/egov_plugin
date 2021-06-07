<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Facility Day Setup</title>
    <style type="text/css">
        div#content {
            margin-left: 2em;
           }
           
        th {
            background-color: Silver;
            border-bottom: 1px solid black;
            }
        
        table.ratetable {
            width:840px; 
            border:1px solid black;
            }
            
        .style2
        {
            width: 209px;
            text-align: center;
        }
        .style4
        {
            background-color: Silver;
            border-bottom: 1px solid black;
            width: 170px;
        }
        .style5
        {
            width: 298px;
        }
        #Text6
        {
            width: 57px;
        }
        .style6
        {
            background-color: Silver;
            border-bottom: 1px solid black;
            width: 209px;
        }
        .style7
        {
            background-color: Silver;
            border-bottom: 1px solid black;
            width: 298px;
        }
        .style8
        {
            background-color: Silver;
            border-bottom: 1px solid black;
            width: 64px;
        }
        .style9
        {
            width: 64px;
        }
        .style10
        {
            width: 170px;
        }
        .style11
        {
            background-color: Silver;
            border-bottom: 1px solid black;
            width: 94px;
        }
        .style12
        {
            width: 94px;
        }
        #Text8
        {
            width: 36px;
        }
        #Text9
        {
            width: 30px;
        }
        #Text10
        {
            width: 35px;
        }
        #Text11
        {
            width: 30px;
        }
    </style>
</head>
<body>
<div id="content">
    <p>
        <strong>Burgess Recreation Center Room 105</strong></p>
    <p>
        <strong>In Season - Monday</strong></p>
<p>
        <input id="Button1" type="button" value="Copy Setup From" />&nbsp;&nbsp;
        <select id="Select11" name="D7">
            <option value="tu">Tuesday</option>
            <option value="we">Wednesday</option>
            <option value="th">Thursday</option>
            <option value="fr">Friday</option>
            <option value="sa">Saturday</option>
            <option value="su">Sunday</option>
        </select> (this might be only on new day)</p>
    <form id="form1" runat="server">
    <div>
    
        <p>
            <input id="Checkbox7" type="checkbox" />&nbsp; Available for public reservations on this day 
            of the week.
        </p>

        <p>
            Opens at&nbsp;&nbsp;&nbsp;
            <input id="Text8" type="text" size="2" maxlength="2" />:<input id="Text9" type="text" size="2" maxlength="2" />&nbsp;
            <select id="Select7" name="D3">
                <option value="am">AM</option>
                <option value="pm">PM</option>
            </select><br />
            Closes at&nbsp;&nbsp;&nbsp;
            <input id="Text10" type="text" size="2" maxlength="2" />:<input id="Text11" type="text" size="2" maxlength="2" />&nbsp;
            <select id="Select9" name="D5">
                <option value="am">AM</option>
                <option value="pm">PM</option>
            </select>&nbsp;
            <select id="Select10" name="D6">
                <option value="t">That day</option>
                <option value="n">Next day</option>
            </select><br /> 
            Latest Reservation Start Time:&nbsp;&nbsp;&nbsp;
            <input id="Text7" type="text" size="2" maxlength="2" />:<input id="Text15" type="text" size="2" maxlength="2" />&nbsp;
            <select id="Select19" name="D5">
                <option value="am">AM</option>
                <option value="pm">PM</option>
            </select>&nbsp;
            &nbsp; 
           
        </p>
        
        <p>
            Time Buffer After a Reservation: <input id="Text16" type="text" size="5" maxlength="5" /> &nbsp;
            <select id="Select21" name="D6">
                <option value="t">Minutes</option>
                <option value="n">Hours</option>
            </select>
        </p>

        <p>
            Minium Rental Time &nbsp; 
            <input id="Text6" type="text" />&nbsp;
            <select id="Select8" name="D6">
                <option value="t">Minutes</option>
                <option value="n">Hours</option>
            </select>
        </p>
    
        <table class="ratetable" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <th class="style7">Type</th><th class="style6">Account</th>
                <th class="style11">Rate<br />Type</th><th class="style8">Rate</th><th class="style4">Starts At</th>
            </tr>
            <tr>
                <td class="style5">
                    <input id="Checkbox1" type="checkbox" />Resident (base)</td>
                <td class="style2">
                    <select id="Select1" name="D1">
                        <option value="1">Account 1</option>
                        <option value="2">Account 2</option>
                        <option value="3">Account 3</option>
                    </select></td>
                <td class="style12" align="center">
                    <select id="Select6" name="D1">
                        <option value="1">Hourly</option>
                        <option value="2">Flat Fee</option>
                    </select>
                </td>
                <td class="style9" align="center">
                    <input id="Text2" type="text" size="5" maxlength="5" />
                </td>
                <td class="style10">&nbsp;</td>
            </tr>
            <tr>
                <td class="style5">
                    <input id="Checkbox2" type="checkbox" />NonResident (+)</td>
                <td class="style2">
                    <select id="Select2" name="D1">
                        <option value="1">Account 1</option>
                        <option value="2">Account 2</option>
                        <option value="3">Account 3</option>
                        <option value="4">Account 4</option>
                    </select>
                </td>
                <td class="style12" align="center">
                    <select id="Select14" name="D1">
                        <option value="1">Hourly</option>
                        <option value="2">Flat Fee</option>
                    </select>
                </td>
                <td class="style9" align="center">
                    <input id="Text1" type="text" size="5" maxlength="5" />
                </td>
                <td class="style10">&nbsp;</td>
            </tr>
            <tr>
                <td class="style5">
                    <input id="Checkbox3" type="checkbox" />Everyone (base)</td>
                <td class="style2">
                    <select id="Select3" name="D1">
                        <option value="1">Account 1</option>
                        <option value="2">Account 2</option>
                        <option value="3">Account 3</option>
                        <option value="4">Account 4</option>
                    </select>
                </td>
                <td class="style12" align="center">
                    <select id="Select15" name="D1">
                        <option value="1">Hourly</option>
                        <option value="2">Flat Fee</option>
                    </select>
                </td>
                <td class="style9" align="center">
                    <input id="Text3" type="text" size="5" maxlength="5" />
                </td>
                <td class="style10">&nbsp;</td>
            </tr>
            <tr>
                <td class="style5">
                    <input id="Checkbox4" type="checkbox" />Building Fee (fee)</td>
                <td class="style2">
                    <select id="Select4" name="D1">
                        <option value="1">Account 1</option>
                        <option value="2">Account 2</option>
                        <option value="3">Account 3</option>
                        <option value="4">Account 4</option>
                    </select>
                </td>
                <td class="style12" align="center">
                    <select id="Select16" name="D1">
                        <option value="1">Hourly</option>
                        <option value="2">Flat Fee</option>
                    </select>
                </td>
                <td class="style9" align="center">
                    <input id="Text4" type="text" size="5" maxlength="5" />
                </td>
                <td class="style10">&nbsp;</td>
            </tr>
            <tr>
                <td class="style5">
                    <input id="Checkbox5" type="checkbox" />Deposit Fee (fee)</td>
                <td class="style2">
                    <select id="Select5" name="D1">
                        <option value="1">Account 1</option>
                        <option value="2">Account 2</option>
                        <option value="3">Account 3</option>
                        <option value="4">Account 4</option>
                    </select>
                </td>
                <td class="style12" align="center">
                    <select id="Select17" name="D1">
                        <option value="1">Hourly</option>
                        <option value="2">Flat Fee</option>
                    </select>
                </td>
                <td class="style9" align="center">
                    <input id="Text5" type="text" size="5" maxlength="5" />
                </td>
                <td class="style10">&nbsp;</td>
            </tr>
            <tr>
                <td class="style5">
                    <input id="Checkbox6" type="checkbox" />Weekend Night Surcharge (fee)</td>
                <td class="style2">
                    <select id="Select12" name="D1">
                        <option value="1">Account 1</option>
                        <option value="2">Account 2</option>
                        <option value="3">Account 3</option>
                        <option value="4">Account 4</option>
                    </select>
                </td>
                <td class="style12" align="center">
                    <select id="Select18" name="D1">
                        <option value="1">Hourly</option>
                        <option value="2">Flat Fee</option>
                    </select>
                </td>
                <td class="style9" align="center">
                    <input id="Text12" type="text" size="5" maxlength="5" />
                </td>
                <td align="center" class="style10">
                    <input id="Text13" type="text" size="2" maxlength="2" />:<input id="Text14" type="text" size="2" maxlength="2" />&nbsp;
                    <select id="Select13" name="D3">
                        <option value="am">AM</option>
                        <option value="pm">PM</option>
                    </select>
                </td>
            </tr>
        </table>
    
    </div>
    </form>
<p>
    <input id="Button2" type="button" value="Save Changes" /></p>
</div>    
</body>
</html>
