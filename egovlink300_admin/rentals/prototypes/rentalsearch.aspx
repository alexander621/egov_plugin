<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    protected void DropDownList2_SelectedIndexChanged(object sender, EventArgs e)
    {

    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Rental Search</title>
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
        
        .style1
        {
            width: 439px;
        }
        .style2
        {
            width: 432px;
        }
        
    </style>
    <script language="javascript">
        <!--
        
        function GoToDatePicks()
        {
            location.href='rentaldateselection.aspx';
        }
        
        //-->
    </script>
</head>
<body>
    <p>
		<font size="+1"><strong>Rental Search</strong></font><br />
	</p>
    
    <div id="searchdiv" style="border: 1px solid black;padding-left: .5em; padding-top: 1em; padding-bottom: 1em;">
    
    <form id="form1" runat="server">
    <table cellpadding="0" cellspacing="0" border="0" width="100%">
    <tr><td valign="top">
        <p>Reservation Type: 
            <asp:DropDownList ID="DropDownList8" runat="server">
                <asp:ListItem Selected="True">Reservation</asp:ListItem>
                <asp:ListItem>On Hold</asp:ListItem>
            </asp:DropDownList>
        </p>
        <p>Registered User: 
            <asp:DropDownList ID="DropDownList7" runat="server">
                <asp:ListItem>None (For On Hold)</asp:ListItem>
                <asp:ListItem Selected="True">Steve Loar</asp:ListItem>
                <asp:ListItem>Peter Selden</asp:ListItem>
            </asp:DropDownList>
        </p>
        <p>
            Is looking for: 
        <asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem Selected="True">Picnic Shelter</asp:ListItem>
            <asp:ListItem>Meeting Room</asp:ListItem>
        </asp:DropDownList>
        </p>
        <p>
            Located at:
            <asp:DropDownList ID="DropDownList2" runat="server" 
                onselectedindexchanged="DropDownList2_SelectedIndexChanged">
                <asp:ListItem Selected="True">Any Location</asp:ListItem>
                <asp:ListItem>Burgess Rec Center</asp:ListItem>
                <asp:ListItem>Gymnastics Center</asp:ListItem>
                <asp:ListItem>Ice Oasis</asp:ListItem>
            </asp:DropDownList>
        </p>
        <p>
            With A Name Like: <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        </p>
        <p>
            Start Date:
            <asp:TextBox ID="TextBox2" runat="server" Text="10/1/2009"></asp:TextBox>
        &nbsp;
        </p>
        <p>
            Order By:
            <asp:DropDownList ID="DropDownList10" runat="server" 
                onselectedindexchanged="DropDownList2_SelectedIndexChanged">
                <asp:ListItem Selected="True">Location, Name</asp:ListItem>
                <asp:ListItem>Name, Location</asp:ListItem>
            </asp:DropDownList>
        </p>
        
    </td>    
    <td valign="top" class="style2">  
        <p>
            For:&nbsp;
            <asp:DropDownList ID="DropDownList9" runat="server">
                <asp:ListItem>All Day</asp:ListItem>
                <asp:ListItem Selected="True">Selected Time Period</asp:ListItem>
            </asp:DropDownList>
        </p>
        <p>
            Start Time:
            <asp:TextBox ID="TextBox5" runat="server" Width="24px" Text="7"></asp:TextBox>
    &nbsp;<b>:</b>
            <asp:TextBox ID="TextBox6" runat="server" Width="26px" Text="00"></asp:TextBox>
    &nbsp;<asp:DropDownList ID="DropDownList6" runat="server">
                <asp:ListItem>AM</asp:ListItem>
                <asp:ListItem Selected="True">PM</asp:ListItem>
            </asp:DropDownList>
        </p> 
        <p>
            End Time:
            <asp:TextBox ID="TextBox4" runat="server" Width="24px" Text="12"></asp:TextBox>
    &nbsp;<b>:</b>
            <asp:TextBox ID="TextBox7" runat="server" Width="26px" Text="00"></asp:TextBox>
    &nbsp;<asp:DropDownList ID="DropDownList5" runat="server">
                <asp:ListItem Selected="True">AM</asp:ListItem>
                <asp:ListItem>PM</asp:ListItem>
            </asp:DropDownList>
        </p> 
        <div>
            Ocurring: <br />
            <asp:RadioButton ID="RadioButton1" GroupName="ocurring" runat="server" Text="Just this once" /><br />
            <asp:RadioButton ID="RadioButton2" GroupName="ocurring" runat="server" Text="Daily" Checked="true" /><br />
            <asp:RadioButton ID="RadioButton3" GroupName="ocurring" runat="server" Text="Weekly On These Days :" /><br />
            &nbsp;&nbsp;&nbsp;&nbsp; <asp:CheckBox ID="CheckBox1" runat="server" Text="Su" /> 
            &nbsp;<asp:CheckBox ID="CheckBox2" runat="server" Text="Mo" /> 
            &nbsp;<asp:CheckBox ID="CheckBox3" runat="server" Text="Tu" /> 
            &nbsp;<asp:CheckBox ID="CheckBox4" runat="server" Text="We" /> 
            &nbsp;<asp:CheckBox ID="CheckBox5" runat="server" Text="Th" /> 
            &nbsp;<asp:CheckBox ID="CheckBox6" runat="server" Text="Fr" /> 
            &nbsp;<asp:CheckBox ID="CheckBox7" runat="server" Text="Sa" /> 
            <br /><asp:RadioButton ID="RadioButton4" GroupName="ocurring" runat="server" Text="Monthly on the" />
        
        &nbsp;
            <asp:DropDownList ID="DropDownList3" runat="server">
                <asp:ListItem>First</asp:ListItem>
                <asp:ListItem>Second</asp:ListItem>
                <asp:ListItem>Third</asp:ListItem>
                <asp:ListItem>Fourth</asp:ListItem>
                <asp:ListItem>Last</asp:ListItem>
            </asp:DropDownList>
    &nbsp;<asp:DropDownList ID="DropDownList4" runat="server">
                <asp:ListItem>Sunday</asp:ListItem>
                <asp:ListItem>Monday</asp:ListItem>
                <asp:ListItem>Tuesday</asp:ListItem>
                <asp:ListItem>Wednesday</asp:ListItem>
                <asp:ListItem>Thursday</asp:ListItem>
                <asp:ListItem>Friday</asp:ListItem>
                <asp:ListItem>Saturday</asp:ListItem>
            </asp:DropDownList>
        
        </div>
        <p>
            End Date:
            <asp:TextBox ID="TextBox3" runat="server" Text="10/3/2009"></asp:TextBox>
        </p>
        
        
    </td></tr></table>
    
    <p>
        <asp:Button ID="Button1" runat="server" Text="Search" class="button" />
    </p>
    </form>
    </div><br /><br />
    <div style="width: 832px; height: 400px; border: 1px solid red;">
        <p><font size="+1"><b>Results</b></font><br /><br />
            <table border="0" cellpadding="2" cellspacing="0" class="resultstable">
            <tr><th>Rental</th><th>Location</th><th>Dimensions</th><th class="style1">Capacity</th><th>Available</th></tr>
            <tr><td onclick="GoToDatePicks();">Adobe Picnic Shelter (click here)</td><td align="center">Swaim Park</td><td align="center">40ft X 35ft</td>
                <td>Tables for 50 but area can hold 100</td><td align="center">Yes</td></tr>
            <tr><td colspan="5">This is a picnic area with a shelter (40 X 35) it has 4 picnic tables and a charcoal grill. Capacity is limited to 100.</td></tr>

            <tr class="altrow first"><td>Happy Picnic Shelter</td><td align="center">Swaim Park</td><td align="center">4ft X 3ft</td>
                <td>Table for 2</td><td align="center">Yes</td></tr>
            <tr class="altrow"><td colspan="5">This is a small picnic area .</td></tr>
            
            <tr class="first"><td>Zoro Picnic Shelter</td><td align="center">Swaim Park</td><td align="center">4ft X 3ft</td>
                <td>Table for 2</td><td align="center">No</td></tr>
            <tr><td colspan="5">This is a small picnic area .</td></tr>
            </table>
        </p>
    </div>

</body>
</html>
