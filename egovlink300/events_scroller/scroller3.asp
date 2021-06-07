<!-- #include file="../includes/common.asp" //-->
<!--#Include file="../include_top_functions.asp"-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: scroller.asp
' AUTHOR: Steve Loar
' CREATED: 10/30/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Display the news item scroller
'
' MODIFICATION HISTORY
' 1.0   03/07/06   Steve Loar - INITIAL VERSION 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

	Dim sScroller, oNewItems, sSQL

	sSQL = "SELECT newsitemid, itemtitle, itemdate, itemtext, itemlinkurl "
	sSQL = sSQL & " FROM egov_news_items WHERE itemdisplay = 1 and orgid = " & iorgid & " ORDER BY itemorder "

	Set oNewItems = Server.CreateObject("ADODB.Recordset")
	oNewItems.Open sSQL, Application("DSN"), 3, 1

	sScroller = ""
	Do While Not oNewItems.EOF
		sScroller = sScroller & vbcrlf & " myScroller1.addItem(""<div class='scrollertitle'>" & oNewItems("itemdate") & " " & JavaScriptSafe(oNewItems("itemtitle"))
		sScroller = sScroller & "</div><div class='scrollertext'>" & JavaScriptSafe(oNewItems("itemtext"))
		If oNewItems("itemlinkurl") <> "" then
			sScroller = sScroller & "<a href='" & oNewItems("itemlinkurl") & "' target='_top'> More &gt;&gt;</a>"
		End If 
		sScroller = sScroller & "</div>"");" 
		oNewItems.MoveNext
	Loop 

	oNewItems.close
	Set oNewItems = Nothing 
%>
<html>
<head>
<title></title>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
<link rel='stylesheet' href='megamillions.css' type='text/css'>
</head>
<body bgcolor='#FFFFFF' text='#000000' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>

<div id='tempholder' onmouseover="myScroller1.stop();" onmouseout="myScroller1.start();"></div>

<script language='JavaScript' src='dhtmllib.js'></script>
<script language='JavaScript' src='scroller.js'></script>

<script language='JavaScript'>

var myScroller1 = new Scroller(0, 0, 180, 180, 0, 0);

myScroller1.addItem("<div align='center'><a href='aboutus/press_release_detail.asp?artid=1769' target='_top'><div class='scrollertitle'>JACKPOT GROWS TO ESTIMATED $15 MILLION IN MEGA MILLIONS;</div> Tuesday`s Drawing Produces More Than 310,000 Winning Tickets</a></div><font color='#000000'>The jackpot in Mega Millions is growing – and there is no shortage of winners! A total of 313,680 tickets from coast to coast won prizes in the Tuesday, December 5, 2006 drawing.  That includes three ...</font><a href='aboutus/press_release_detail.asp?artid=1769' target='_top'> More &gt;&gt;</a>")
myScroller1.addItem("<div align='center'><a href='aboutus/press_release_detail.asp?artid=1766' target='_top'><div class='scrollertitle'>ONE MEGA MILLIONS TICKET WINS ESTIMATED $40 MILLION JACKPOT;</div> Ticket Sold In Ohio Matches All Six Numbers To Win</a></div><font color='#000000'>One ticket matched all six Mega Millions numbers in the Friday, December 1, 2006 drawing, and now players are checking their tickets to see who won the estimated $40 million jackpot.  The jackpot-winn...</font><a href='aboutus/press_release_detail.asp?artid=1766' target='_top'> More &gt;&gt;</a>")
myScroller1.addItem("<div align='center'><a href='aboutus/press_release_detail.asp?artid=1762' target='_top'><div class='scrollertitle'>MEGA MILLIONS JACKPOT GROWS TO ESTIMATED $40 MILLION;</div> Jackpot Is On The Move As Nearly 440,000 Tickets Win Prizes</a></div><font color='#000000'>Countless Mega Millions players from coast to coast had their eyes on the prize in the Tuesday, November 28, 2006 drawing.  A total of 439,899 tickets won prizes, including three tickets that each won...</font><a href='aboutus/press_release_detail.asp?artid=1762' target='_top'> More &gt;&gt;</a>")
myScroller1.addItem("<div align='center'><a href='aboutus/press_release_detail.asp?artid=1758' target='_top'><div class='scrollertitle'>MORE THAN 320,000 TICKETS WIN IN MEGA MILLIONS DRAWING;</div> Jackpot Grows To Estimated $31 Million</a></div><font color='#000000'>The Friday, November 24 Mega Millions drawing really brought out the winners!  A total of 321,032 tickets from coast to coast won Mega Millions prizes, including three tickets that nearly hit the jack...</font><a href='aboutus/press_release_detail.asp?artid=1758' target='_top'> More &gt;&gt;</a>")
myScroller1.addItem("<div align='center'><a href='aboutus/press_release_detail.asp?artid=1755' target='_top'><div class='scrollertitle'>MEGA MILLIONS ESTIMATED JACKPOT GROWS TO $23 MILLION;</div> Nearly 370,000 Tickets Win Prizes In Tuesday`s Drawing</a></div><font color='#000000'>Even though there was no jackpot winner in Tuesday's Mega Millions drawing, a lot of players still walked away winners. With nine different ways to win, a total of 369,032 tickets won prizes in the No...</font><a href='aboutus/press_release_detail.asp?artid=1755' target='_top'> More &gt;&gt;</a>")
myScroller1.addItem("<div align='center'><a href='aboutus/press_release_detail.asp?artid=1191' target='_top'><div class='scrollertitle'>JACKPOT HITS ESTIMATED $53 MILLION IN MEGA MILLIONS;</div> EXCITEMENT GROWS FROM COAST TO COAST!</a></div><font color='#000000'>Mega Millions players across the country are feeling the excitement as the jackpot continues to grow. When Tuesday's winning numbers were drawn, they turned 342,708 tickets into Mega Millions winners....</font><a href='aboutus/press_release_detail.asp?artid=1191' target='_top'> More &gt;&gt;</a>")

myScroller1.setPause(1000);
myScroller1.setSpeed(35);

function runmikescroll() 
{
  var layer;
  var mikex, mikey;
  layer = getLayer("placeholder");
  mikex = getPageLeft(layer);
  mikey = getPageTop(layer);
  myScroller1.create();
  myScroller1.hide();
  myScroller1.moveTo(mikex, mikey);
  myScroller1.setzIndex(100);
  myScroller1.show();
}

window.onload=runmikescroll;

</script>

<div id='placeholder' style='position:relative; width:175; height:300px;' onmouseover="myScroller1.stop();" onmouseout="myScroller1.start();"></div>

</body>

</html>


<%
'------------------------------------------------------------------------------------------------------------
' FUNCTION JavaScriptSafe( sString )
'------------------------------------------------------------------------------------------------------------
Function JavaScriptSafe( sString )
	If Not VarType( sString ) = vbString Then JavaScriptSafe = sString : Exit Function
	JavaScriptSafe = Replace( sString, Chr(34), "'" )
End Function
%>

