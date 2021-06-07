<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!--#Include file="../include_top_functions.asp"-->
<style type="text/css">
div.scrollertitle2 {
  background-color: #F3F4F6;
  width:            180px;
  color:            #0066c1;
  font-family:      'Lucida Sans Unicode','Lucida Grande',Arial,sans-serif;
  font-size:        13px;
  font-style:       normal;
  font-weight:      bold;
  line-height:      normal;
  margin:           0 0 0 0px;
  text-decoration:  none;
}

div.scrollertext2 {
  background-color: #F3F4F6;
  color:            #0066c1;
  width:            180px;
  font-family:      'Lucida Sans Unicode','Lucida Grande',Arial,sans-serif;
  font-size:        11px;
  font-style:       normal;
  font-weight:      normal;
  line-height:      normal;
  margin:           0 0 0 0px;
  text-decoration:  none;
}
</style>
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

	Dim sScroller, oNewItems, sSQL, x

	sSQL = "SELECT newsitemid, itemtitle, itemdate, itemtext, itemlinkurl "
	sSQL = sSQL & " FROM egov_news_items WHERE itemdisplay = 1 and orgid = " & iorgid 
	sSQL = sSQL & " AND (publicationstart is null OR publicationstart <= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
	sSQL = sSQL & " AND (publicationend is null OR publicationend >= cast(cast(datepart(mm,getdate()) as varchar) + '/' + cast(datepart(dd,getdate()) as varchar) +'/' + cast(datepart(yyyy,getdate()) as varchar) + ' 00:00:000' as datetime) ) "
	sSQL = sSQL & " ORDER BY itemorder "

	Set oNewItems = Server.CreateObject("ADODB.Recordset")
	oNewItems.Open sSQL, Application("DSN"), 3, 1

	sScroller = ""
	x = 0
	do while not oNewItems.EOF
		  sScroller = sScroller & vbcrlf & " pausecontent[" & x & "]= ""<div class='scrollertitle2'>" & oNewItems("itemdate") & " " & JavaScriptSafe(oNewItems("itemtitle"))
		  sScroller = sScroller & "</div><div class='scrollertext2'>" & JavaScriptSafe(oNewItems("itemtext"))

  		if oNewItems("itemlinkurl") <> "" then
    			sScroller = sScroller & "<br /><a href='" & oNewItems("itemlinkurl") & "' target='_top'> More &gt;&gt;</a>"
    end if

  		sScroller = sScroller & "</div>"";" & vbcrlf

  		oNewItems.MoveNext
   	x = x + 1
	loop

	oNewItems.close
	Set oNewItems = Nothing 
%>

<html>
<head>

	<title></title>
	
	<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script type="text/javascript">

		var pausecontent = new Array();
		<% =sScroller %>

	</script>

<script type="text/javascript">

/***********************************************
* Pausing up-down scroller- © Dynamic Drive (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit http://www.dynamicdrive.com/ for this script and 100s more.
***********************************************/

function pausescroller(content, divId, divClass, delay)
{
	this.content=content //message array content
	this.tickerid=divId //ID of ticker div to display information
	this.delay=delay //Delay between msg change, in miliseconds.
	this.mouseoverBol=0 //Boolean to indicate whether mouse is currently over scroller (and pause it if it is)
	this.hiddendivpointer=1 //index of message array for hidden div
	<% if x > 1 then %>
		document.write('<div id="'+divId+'" class="'+divClass+'" style="position: relative; overflow: hidden"><div class="innerDiv" style="position: absolute; width: 100%" id="'+divId+'1">'+content[0]+'</div><div class="innerDiv" style="position: absolute; width: 100%; visibility: hidden" id="'+divId+'2">'+content[1]+'</div></div>')
	<% else %>
		document.write('<div id="'+divId+'" class="'+divClass+'" style="position: relative; overflow: hidden"><div class="innerDiv" style="position: absolute; width: 100%" id="'+divId+'1">'+content[0]+'</div><div class="innerDiv" style="position: absolute; width: 100%; visibility: hidden" id="'+divId+'2">'+content[0]+'</div></div>')
	<% end if %>
	var scrollerinstance=this
	if (window.addEventListener) //run onload in DOM2 browsers
		window.addEventListener("load", function(){scrollerinstance.initialize()}, false)
	else if (window.attachEvent) //run onload in IE5.5+
		window.attachEvent("onload", function(){scrollerinstance.initialize()})
	else if (document.getElementById) //if legacy DOM browsers, just start scroller after 0.5 sec
		setTimeout(function(){scrollerinstance.initialize()}, 500)
}

// -------------------------------------------------------------------
// initialize()- Initialize scroller method.
// -Get div objects, set initial positions, start up down animation
// -------------------------------------------------------------------

pausescroller.prototype.initialize=function()
{
	this.tickerdiv=document.getElementById(this.tickerid)
	this.visiblediv=document.getElementById(this.tickerid+"1")
	this.hiddendiv=document.getElementById(this.tickerid+"2")
	this.visibledivtop=parseInt(pausescroller.getCSSpadding(this.tickerdiv))
	//set width of inner DIVs to outer DIV's width minus padding (padding assumed to be top padding x 2)
	this.visiblediv.style.width=this.hiddendiv.style.width=this.tickerdiv.offsetWidth-(this.visibledivtop*2)+"px"
	this.getinline(this.visiblediv, this.hiddendiv)
	this.hiddendiv.style.visibility="visible"
	var scrollerinstance=this
	document.getElementById(this.tickerid).onmouseover=function(){scrollerinstance.mouseoverBol=1}
	document.getElementById(this.tickerid).onmouseout=function(){scrollerinstance.mouseoverBol=0}
	if (window.attachEvent) //Clean up loose references in IE
		window.attachEvent("onunload", function(){scrollerinstance.tickerdiv.onmouseover=scrollerinstance.tickerdiv.onmouseout=null})
	setTimeout(function(){scrollerinstance.animateup()}, this.delay)
}


// -------------------------------------------------------------------
// animateup()- Move the two inner divs of the scroller up and in sync
// -------------------------------------------------------------------

pausescroller.prototype.animateup=function()
{
	var scrollerinstance=this
	if (parseInt(this.hiddendiv.style.top)>(this.visibledivtop+5))
	{
		this.visiblediv.style.top=parseInt(this.visiblediv.style.top)-5+"px"
		this.hiddendiv.style.top=parseInt(this.hiddendiv.style.top)-5+"px"
		setTimeout(function(){scrollerinstance.animateup()}, 50)
	}
	else
	{
		this.getinline(this.hiddendiv, this.visiblediv)
		this.swapdivs()
		setTimeout(function(){scrollerinstance.setmessage()}, this.delay)
	}
}

// -------------------------------------------------------------------
// swapdivs()- Swap between which is the visible and which is the hidden div
// -------------------------------------------------------------------

pausescroller.prototype.swapdivs=function()
{
	var tempcontainer=this.visiblediv
	this.visiblediv=this.hiddendiv
	this.hiddendiv=tempcontainer
}

pausescroller.prototype.getinline=function(div1, div2)
{
	div1.style.top=this.visibledivtop+"px"
	div2.style.top=Math.max(div1.parentNode.offsetHeight, div1.offsetHeight)+"px"
}

// -------------------------------------------------------------------
// setmessage()- Populate the hidden div with the next message before it's visible
// -------------------------------------------------------------------

pausescroller.prototype.setmessage=function()
{
	var scrollerinstance=this
	if (this.mouseoverBol==1) //if mouse is currently over scoller, do nothing (pause it)
		setTimeout(function(){scrollerinstance.setmessage()}, 100)
	else
	{
		var i=this.hiddendivpointer
		var ceiling=this.content.length
		this.hiddendivpointer=(i+1>ceiling-1)? 0 : i+1
		this.hiddendiv.innerHTML=this.content[this.hiddendivpointer]
		this.animateup()
	}
}

pausescroller.getCSSpadding=function(tickerobj)
{ //get CSS padding value, if any
	if (tickerobj.currentStyle)
		return tickerobj.currentStyle["paddingTop"]
	else if (window.getComputedStyle) //if DOM2
		return window.getComputedStyle(tickerobj, "").getPropertyValue("padding-top")
	else
		return 0
}

</script>

</head>

<body style="background-color: #F3F4F6" id="scrollerbody2">

	<script type="text/javascript">

		new pausescroller(pausecontent, "pscroller1", "someclass", 4000)

	</script>

</body>
</html>

<%
'----------------------------------------------------------------------------------------
' FUNCTION JavaScriptSafe( sString )
'----------------------------------------------------------------------------------------
Function JavaScriptSafe( sString )
	If Not VarType( sString ) = vbString Then JavaScriptSafe = sString : Exit Function
	JavaScriptSafe = Replace( sString, Chr(34), "'" )
End Function
%>