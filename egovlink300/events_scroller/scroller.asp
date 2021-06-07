<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: scroller.asp
' AUTHOR: David Boyer
' CREATED: 04/12/2012
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Display the events scroller.
'
' MODIFICATION HISTORY
' 1.0 04/12/2012 David Boyer - INITIAL VERSION (Copy of News Scroller)
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim sScroller, oEventItems, sSQL, x

 lcl_showDaysLimit = getEventsDaysLimit(iOrgID)

	sSQL = "SELECT e.eventid, "
 sSQL = sSQL & " e.eventdate, "
 sSQL = sSQL & " e.eventduration, "
 sSQL = sSQL & " e.subject, "
 sSQL = sSQL & " e.message, "
 sSQL = sSQL & " e.categoryid, "
 sSQL = sSQL & " ec.categoryname "
	sSQL = sSQL & " FROM events e "
 sSQL = sSQL &   " LEFT OUTER JOIN eventcategories ec ON e.categoryid = ec.categoryid "
 sSQL = sSQL & " WHERE e.orgid = " & iorgid
 sSQL = sSQL & " AND cast(e.eventdate as date) >= cast(getdate() as date) "
 sSQL = sSQL & " AND cast(e.eventdate as date) <= cast(dateadd(day," & lcl_showDaysLimit & ",getdate()) as date) "
	sSQL = sSQL & " ORDER BY e.eventdate "

	set oEventItems = Server.CreateObject("ADODB.Recordset")
	oEventItems.Open sSQL, Application("DSN"), 3, 1

	sScroller = ""
	x = 0

 if not oEventItems.eof then
   	do while not oEventItems.eof
 						dEnd           = ""
 						'sCategory      = ""
	   			lcl_event_date = oEventItems("eventdate")

   				if oEventItems("EventDuration") > 0 then
     					dEnd = DateAdd("n",oEventItems("EventDuration"),oEventItems("EventDate"))

     					if DateDiff("d",dEnd,oEventItems("EventDate")) = 0 then
       						dEnd = FormatDateTime(dEnd,vbLongTime)
						    end if

   						 dEnd = " - " & dEnd
				   end if

  					'if oEventItems("CategoryID") <> 0 then
    			'			sCategory = "(" & getCategoryName(oEventItems("categoryid")) & ") "
  					'end if

 					'Format the event date/time (remove the seconds from the date(s))
 					'This displays the "12:00:00 AM" IF a duration exists
	   			if left(FormatDateTime(oEventItems("eventdate"),vbLongTime),11) = "12:00:00 AM" and oEventItems("eventduration") > 0 then
 				    	lcl_event_date = oEventItems("eventdate") & " " & formatdatetime(oEventItems("eventdate"),vbLongTime)
  					end if

  					if oEventItems("eventduration") > 0 then
		     			if Left(FormatDateTime(Replace(dEnd," - ", ""),vbLongTime),11) = "12:00:00 AM" then
      							lcl_end_time = replace(dEnd," - ", "")
      							lcl_end_time = lcl_end_time & " " & formatdatetime(lcl_end_time,vbLongTime)
       						lcl_end_time = " - " & lcl_end_time
      							dEnd         = lcl_end_time
    						end if
  					end if

  					formatEventDateTime lcl_event_date, dEnd, sDate1, sDate2

  	   	sScroller = sScroller & " pausecontent[" & x & "]= ""<div class='scrollertitle'>"
       sScroller = sScroller & sDate1 & " " & sDate2
       sScroller = sScroller & " "

       if oEventItems("categoryname") <> "" then
          sScroller = sScroller & "("
          sScroller = sScroller & JavaScriptSafe(oEventItems("categoryname"))
          sScroller = sScroller & ") "
       end if

       sScroller = sScroller & JavaScriptSafe(oEventItems("subject"))
       sScroller = sScroller & "</div>"
     		sScroller = sScroller & "<div class='scrollertext'>" & JavaScriptSafe(oEventItems("message"))

     		'if oEventItems("itemlinkurl") <> "" then
		     '  	sScroller = sScroller & "<br /><a href='" & oEventItems("itemlinkurl") & "' target='_top'> More &gt;&gt;</a>"
     		'end if

     		sScroller = sScroller & "</div>"";" & vbcrlf

  	   	oEventItems.movenext
     		x = x + 1
   	loop
 end if

	oEventItems.close
	set oEventItems = nothing 
%>

<html>
<head>

 	<title></title>
	
	 <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
	
 	<link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

<style>
  div#pscroller1 {
     border: 1pt solid #ff0000;
     width:  1000px;
     height: 800px;
  }

  div.scrollertitle {
     font-size: 32px;
  }

  div.scrollertext {
     font-size: 26px;
  }
</style>

 	<script type="text/javascript">
  		var pausecontent = new Array();
  		<%=sScroller%>
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
<body id="scrollerbody">
<%
 'If at least one news item is returned then display the scroller
  if x > 0 then
     response.write "<script type=""text/javascript"">" & vbcrlf
     response.write "  new pausescroller(pausecontent, ""pscroller1"", ""someclass"", 3000)" & vbcrlf
     response.write "</script>" & vbcrlf
  end if

  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
 function JavaScriptSafe( sString )
   lcl_return = ""

	  if not VarType( sString ) = vbString then JavaScriptSafe = sString : exit function
  	lcl_return = sString
   lcl_return = replace( lcl_return, chr(34), "'" )
   lcl_return = replace( lcl_return, chr(10), "<br />" )
   lcl_return = replace( lcl_return, chr(13), "<br />" )

   JavaScriptSafe = lcl_return

 end function

'------------------------------------------------------------------------------
function getEventsDaysLimit(iOrgID)
  dim lcl_return, sOrgID

  lcl_return = 14
  sOrgID     = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if sOrgID > 0 then
     sSQL = "SELECT eventScrollerDateLimit "
     sSQL = sSQL & " FROM organizations "
     sSQL = sSQL & " WHERE orgid = " & sOrgID

   	set oGetOrgDateLimit = Server.CreateObject("ADODB.Recordset")
   	oGetOrgDateLimit.Open sSQL, Application("DSN"), 3, 1

    if not oGetOrgDateLimit.eof then
       if oGetOrgDateLimit("eventScrollerDateLimit") <> "" then
          lcl_return = oGetOrgDateLimit("eventScrollerDateLimit")
       end if
    end if

    oGetOrgDateLimit.close
    set oGetOrgDateLimit = nothing

  end if

  getEventsDaysLimit = lcl_return

end function
%>