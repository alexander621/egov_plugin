<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
<title>PayPal - PayPal Website Payments - Confirm Your Payment</title>
<META http-equiv="DESCRIPTION" content="PayPal lets you send money to anyone with email. PayPal is free for consumers and works seamlessly with your existing credit card and checking account. You can settle debts, borrow cash, divide bills or split expenses with friends all without going to an ATM or looking for your checkbook."><META http-equiv="KEYWORDS" content="Send, money, payments, credit, credit card, instant, money, financial services, mobile, wireless, WAP, cell phones, two-way pagers, Windows CE">

<link rel="stylesheet" type="text/css" href="pp_styles_111402.css">

	<style type="text/css">
	
		/* these apply to all cowp styles*/
		body 						{background-color: #FFFFFF; color: #000000;}
		A 							{color: #0033CC;}
		.pptext 						{color: #000000;}
		.ppsmalltext				{color: #000000;}
		
		/* these are styles that always have to override no matter what the body bg color. They deal with
		the modules on the site that have colored backgrounds that we didn't want to remove (e.g. the pie modules) */
		.ppsmalltextblack 			{color: #000000; font-size: 11px; font-family: verdana,arial,helvetica,sans-serif; font-weight: 400;}
		.ppdashheaderblack			{color: #000000; font-size: 11px; font-family: verdana,arial,helvetica,sans-serif; font-weight: 700;}
		.ppsmalltextbluelink		{color: #0033CC; font-size: 11px; font-family: verdana,arial,helvetica,sans-serif; font-weight: 400; text-decoration: underline;}
		.ppwaxloginbg				{background-color: #FFFFFF;}
	    	.ppwaxloginborder			{background-color: #000000;}	
		.ppmessage					{color: #000000; font-size: 13px; font-family: verdana,arial,helvetica,sans-serif; font-weight: 700;}  
		
		/* these overrides apply to all non-white bg cowp styles */
			
	</style>
<link rel="stylesheet" type="text/css" href="pp_table_styles.css">

<script>
	function demo_notice(){
	alert("Hyperlinks and button clicks - with the exception of\nthe CONTINUE button - that take you outside the\npayment process have been disabled for the demo.\nPlease press the CONTINUE button on the bottom\nof the page to move forward thru the demo.");
}
</script>


</head>


<body>
		

	<table class="ppwaxtablewidth" cellpadding="0" cellspacing="0" border="0" align=center>
		<tr>
			<td colspan=2><img src="images/pixel.gif" width=1 height=5></td>
		</tr>
		<tr valign=top>
			<td align=left width=18%>
				<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><img src="images/logo-xclickBox.gif" border="0" align=bottom></td>
					</tr>
				</table>
			</td>
			<td width="100%" align="center" class="pptext"></td>
		</tr>
		<tr>
			<td colspan=2><img src="images/pixel.gif" width=1 height=5></td>
		</tr>
	</table>


<table class=pscowpimage cellspacing="0" border="0" align=center
			cellpadding="2" bgcolor=FFFFFF	>

			<tr>
			<td>
				<table class=pscowpimage cellpadding=0 cellspacing=0 border=0 align=left bgcolor=FFFFFF>
					<tr><td>** DEMONSTRATION ONLY - NO PAYMENT INFORMATION PROCESSED ** </td></tr>
				</table>
			</td>	
		</tr>
</table>

<br class="h5">
<table bgcolor="#336699" width="100%" cellspacing="0" cellpadding="0" border="0" align="center">
<tr>
	<td  width="100%"><img src="images/pixel.gif" width=1 height="27" border="0"></td>
</tr>
</table>
<br class="h5">

<img src="images/pixel.gif" width=1 height=10><br>
  




<table class="main" align="center" cellpadding="0" cellspacing="0" border="0">
 <tr>
  <td>

<table class="main" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr>
	  	<td width="100%" class="ppheading">Confirm Your Payment</td>
	  	  	<td nowrap class="ppsmalltext">Secure Transaction&nbsp;</td>
	  	<td><img src="images/secure_lock_2.gif" border="0"></td>
	</tr>
	<tr>
		<td colspan="3"><img src="images/pixel.gif" width="2" height="2"></td>
	</tr>
</table>

<table class="main" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr>
		<td><img src="images/pixel.gif" width=6 height=6></td>
	</tr>
	<tr>
		<td bgcolor=#000000 width=100%><img src="images/pixel.gif" width=1 height=2></td>
	</tr>
	<tr>
		<td><img src="images/pixel.gif" width=6 height=6></td>
	</tr>
</table>
  </td>
</tr>
<tr>
  <td>

<table class="main" cellspacing="0" cellpadding="0" border="0" align="center">
	<tr>
		<td class="pptext">
			Review the payment details below and click <span class="ppem106">Pay</span> to complete your secure payment.
		</td>
	</tr>
	<!-- added for 4419 -->
		
		
</table>
<br></td></tr></table>

	<table class="main" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td>

<table width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
	<td width="150"><img src="images/pixel.gif" width="150" height="1"></td>
	<td width="6"><img src="images/pixel.gif" width="6" height="1"></td>
	<td width="100%"><img src="images/pixel.gif" width="432" height="1"></td>
</tr>
			<tr valign="top">
			<td align="right" class="pplabel">Pay To:</td>
			<td><br class="text_spacer"></td>
			<td class="pptext">City of Selden&nbsp;</td>
		</tr>
					<tr valign="top">
				<td align="right" class="pplabel">User Status:</td>
				<td><br class="text_spacer"></td>
				<td class="pptext">	<span class="pptext">
	

&nbsp;</td>
			</tr>
			
		
	<tr valign="top">
		<td align="right" class="pplabel">		   Payment For:
						</td>
		<td><br class="text_spacer"></td>
		<td>
						<span class="pptext">
															<%=request("item_name")%>														&nbsp;
			</span>
							<br>
	<span class="ppsmalltext">
		<%
			' GET CUSTOM PAYMENT FORM FIELDS
			For each oField in Request.Querystring
				If UCASE(left(oField,7)) = "CUSTOM_" AND instr(oField,"ef:") = 0 Then
					response.write  UCASE(replace(oField,"custom_","")) & ": " & request(oField) & "<br>"
				End If 
			Next 
		%>
		</span>
		</td>
	</tr>
							    				<tr valign="top">
					<td align="right" class="pplabel">Quantity:</td>
					<td><br class="text_spacer"></td>
					<td class="pptext">1&nbsp;</td>
				</tr>
    								
    	
	 	
	<tr valign="top">
		<td align="right" class="pplabel"><label for="currency">Currency:</td>
		<td><br class="text_spacer"></td>
				<td>
			<span class="pptext">U.S. Dollars</span>
			<span class="ppsmalltext"></span>
		</td>
			</tr>

		<tr valign="top">
		<td align="right" class="pplabel">						
												Amount:									
					</td>
		<td><br class="text_spacer"></td>
		<td>
												<span class="pptext">$<%=request("custom_paymentamount")%> USD</span>
											
		</td>
	</tr>
		
	
	
	
				
						
												
		</table>
	</td></tr></table>
<table class="main" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td>
		
		<hr class="dotted">
		<table class="title" align=center cellpadding=0 cellspacing=0 border=0>
		<tr><td class="pptext">
							
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="pptablesubheading">
	<tr>
		<td><img src="images/pixel.gif" width="6" height="6"></td>
	</tr>
	<tr valign=bottom>
		<td nowrap class="ppsubheading">Source of Funds</td>
	</tr>
	<tr>
		<td><img src="images/pixel.gif" width="6" height="6"></td>
	</tr>
</table>		
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td valign="top" width="55%">	<table width="100%" cellpadding="0" cellspacing="0" border="0">
		
			
			<tr>
			<td width="150"><img src="images/pixel.gif" width="150" height="1"></td>
			<td width="6"><img src="images/pixel.gif" width="6" height="1"></td>
			<td width="100%">
									<img src="images/pixel.gif" width="432" height="1">
							</td>
		</tr>
	
	
		
			
		
		
		
		
								<tr valign=top>
									<td align=right class="pplabel">Credit Card:</td>
			 						<td><br class="text_spacer"></td>
										<td class="pptext">$<%=request("amount")%> USD from MasterCard XXXX-XXXX-XXXX-4444 
																</td>
				</tr>
									
			
					<tr valign=top>
				<td colspan="3">
	

						
			<br>This credit card transaction will appear on your bill as "CITYSELDEN".
																			
				</td>
			</tr>
			
	  	
	
	</table>
</td>

	</tr>
</table>
		</td>
		</tr></table>
	
	<table class="main" align="center" cellpadding="0" cellspacing="0" border="0">
		<tr><td>
		

		</td></tr>
	</table>
    
	
	<hr class="dotted">
</td></tr></table>

<table class="main" align="center" cellpadding="0" cellspacing="0" border="0">
 <tr>
  <td>
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="pptablesubheading">
	<tr>
		<td><img src="/en_US/i/scr/pixel.gif" width="6" height="6"></td>
	</tr>
	<tr valign=bottom>
		<td nowrap class="ppsubheading">Shipping Information</td>
	</tr>
	<tr>
		<td><img src="/en_US/i/scr/pixel.gif" width="6" height="6"></td>
	</tr>
</table>

<!--buyer credit flag-->
<table width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr valign="top">
		<td></td>
		<td nowrap class="ppsmalltextbold"><img src="/images/pixel.gif" width=1 height=4><br>Ship to</td>
		<td><!--<img src="/images/pixel.gif" width="10" height="20">--></td>
		<td width="100%"><SELECT name="shipping_address_id" ><OPTION value="ZV04RsVwXymjFPrZSvAGR4ACEjswFBryxj61j1iWrTN4eUo8qYiCF07V9T4" SELECTED>4303 Hamilton, Cincinnati, OH 45223, United States (Confirmed)	</SELECT></td>
	</tr>
	<tr valign="top">
		<td>
			<input type=radio name="shipping_address_present" value="0" checked >
		</td>
		<td colspan="3" class="ppsmalltextbold">
			<img src="/en_US/i/scr/pixel.gif" width="1" height="4"><br>No shipping address required
		</TD>
		</tr>
</table>

  </td>
 </tr>
 <tr>
  <td>


											
		


</td></tr>
</table>
<table class="main" cellpadding="0" cellspacing="0" border="0" align="center">
	<td class="maintd" align="left">
		<input type="hidden" name="paysubmit" value="0">

																	<table class="main" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr>
		<td><img src="/en_US/i/scr/pixel.gif" width=6 height=6></td>
	</tr>
	<tr>
		<td bgcolor=#000000 width=100%><img src="/en_US/i/scr/pixel.gif" width=1 height=2></td>
	</tr>
	<tr>
		<td><img src="/en_US/i/scr/pixel.gif" width=6 height=6></td>
	</tr>
</table><table class="main" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr>
		<td colspan="6"><img src="/en_US/i/scr/pixel.gif" width="4" height="4"></td>
	</tr>
	<tr>
		<td width="6"><img src="/en_US/i/scr/pixel.gif" width="6" height="6"></td>
		<td width="100%" class="ppsmalltext" align=right ><img src="../images/flashing_foward.gif" align=right></td>
		<td>
			
			<input type="button" onClick="location.href='paypal_demo_page3.asp?<%= REQUEST.QUERYSTRING %>';" name="submit.x" value="&nbsp;&nbsp;&nbsp;&nbsp;Pay&nbsp;&nbsp;&nbsp;&nbsp;" class="ppbuttonhot"></td>
		<td width="6"><img src="/en_US/i/scr/pixel.gif" width="6" height="6"></td>
		<td><input type="button" name="cancel.x" value="Cancel" onclick="location.href='http://www.egovlink.com/demo';"  class="ppbutton"></td>
		<td width="6"><img src="/en_US/i/scr/pixel.gif" width="6" height="6"></td>
	</tr>
	<tr>
		<td colspan="6"><img src="/en_US/i/scr/pixel.gif" width="6" height="6"></td>
	</tr>
</table>	




<!-- Begin Context -->	<table class="main" cellpadding=0 cellspacing=0 border=0 align=center>
		<tr>
			<td align=center valign=top class="ppsmallnote">
				<style type="text/css">
					.links { 
							color: #999999;
							text-decoration: none;
					}
				</style>
				&nbsp;
				PayPal protects your privacy and security.<br> For more information, see our <a href="#" onClick="demo_notice();" >Privacy Policy</a> and <a href="#" onClick="demo_notice();" >User Agreement</a>.<br>
				<br>&nbsp;</td>
		</tr>
	</table>
<br />
</body>
</html></form>

  </td>
 </tr>
</table>

</body>
</html>