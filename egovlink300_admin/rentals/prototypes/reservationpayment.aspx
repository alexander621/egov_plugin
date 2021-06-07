<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Reservation Payment</title>
</head>
<body>
    <form id="form1" runat="server">
    <p>
		<font size="+1"><strong>Reservation Payment</strong></font><br />
	</p>
    <div>
    <fieldset>
    <table cellpadding="2" cellspacing="0" border="0" class="none">
        <tr><td colspan="3" align="left"><b>Adobe Picnic Shelter, Swaim Park 11/13/2009 7:00PM to 12:00AM</b></td></tr>
        <tr><td>Resident</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
        <tr><td>Nonresident</td><td align="center"><input type="text" value="25.00" size="5" /></td></tr>
        <tr><td>Chairs</td><td align="center"><input type="text" value="20.00" size="5" /></td></tr>
        <tr><td>Large Round Table</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
        <tr><td>Small Cocktail Tables</td><td align="center"><input type="text" value="45.00" size="5" /></td></tr>
        <tr><td>Tablecloths</td><td align="center"><input type="text" value="40.00" size="5" /></td></tr>
    </table>
    </fieldset>
    <fieldset>
    <table cellpadding="2" cellspacing="0" border="0" class="none">
        <tr><td colspan="2" align="left"><b>Adobe Picnic Shelter, Swaim Park 11/14/2009 7:00PM to 12:00AM</b></td></tr>
        <tr><td>Resident</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
        <tr><td>Nonresident</td><td align="center"><input type="text" value="25.00" size="5" /></td></tr>
        <tr><td>Chairs</td><td align="center"><input type="text" value="20.00" size="5" /></td></tr>
        <tr><td>Large Round Table</td><td align="center"><input type="text" value="100.00" size="5" /></td></tr>
        <tr><td>Small Cocktail Tables</td><td align="center"><input type="text" value="45.00" size="5" /></td></tr>
        <tr><td>Tablecloths</td><td align="center"><input type="text" value="40.00" size="5" /></td></tr>
    </table>
    </fieldset>
    
    <fieldset>
    <table cellpadding="2" cellspacing="0" border="0" class="none">
    <tr><td>Deposit</td><td align="center"><input type="text" value="50.00" size="5" /></td></tr>
    <tr><td>Alcohol Surcharge</td><td align="center"><input type="text" value="75.00" size="5" /></td></tr>
    </table>
    </fieldset>
    
    <fieldset><legend><strong>Payment&nbsp;</strong></legend><br />
        <input type="hidden" value="0.00" name="amount" />
        <table border="0" cellpadding="3" cellspacing="0" width="50%"><tr><td class="label" align="right" nowrap="nowrap">Citizen Location:</td><td>
        <select name="PaymentLocationId">
        <option value="1">Walk In</option>
        <option value="2">Phone</option>
        </select></td></tr><tr><td class="label" align="right" nowrap="nowrap">Credit Card Scan: </td><td><input type="text" value="" name="amount1" size="10" maxlength="9" onblur="addTotal()" /></td></tr><tr><td class="label" align="right" nowrap="nowrap">Check : </td><td><input type="text" value="" name="amount2" size="10" maxlength="9" onblur="addTotal()" />&nbsp;<strong>Check #: </strong><input type="text" value="" name="checkno" size="8" maxlength="8" /></td></tr><tr><td class="label" align="right" nowrap="nowrap">Cash: </td><td><input type="text" value="" name="amount3" size="10" maxlength="9" onblur="addTotal()" /></td></tr><tr><td class="label" align="right" nowrap="nowrap">Other: </td><td><input type="text" value="" name="amount8" size="10" maxlength="9" onblur="addTotal()" /></td></tr><tr><td class="label" align="right" nowrap="nowrap">Citizen Accounts: </td><td><input type="text" value="" name="amount4" size="10" maxlength="9" onblur="addTotal()" />&nbsp; <strong>From:</strong>
        <select name="accountid">
        <option value="19221">Ralph Cramdon (0.00) </option>
        <option value="19222">Heidi Tester (0.00) </option>
        <option value="19223">Joey Tester (0.00) </option>
        <option value="19220">Sue Tester (0.00) </option>
        <option value="5158">Tommy Tester (0.00) </option>
        </select></td></tr><tr><td class="label" align="right" nowrap="nowrap">Payment Total:</td><td><span id="total">0.00</span></td></tr><tr><td class="label" align="right" nowrap="nowrap">Balance Due:</td><td><span id="balancedue">930.00</span></td></tr><tr><td class="label" align="right" nowrap="nowrap">Notes:</td><td><textarea name="notes" class="purchasenotes"></textarea></td></tr></table><br /><br />
        <input type="button" class="button" name="complete" value="Complete Purchase" onClick="validate()" />
    </fieldset>
    </div>
    </form>
</body>
</html>
