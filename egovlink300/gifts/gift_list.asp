<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: gift_list.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  The display of all gifts available
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>

<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>


	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/easyform.js"></script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<p>
<font class=pagetitle>Commemorative Gifts</font> <br />


<%	RegisteredUserDisplay( "../" ) %>

<div class="facilitymain giftsmain">  
<% If orgid = 26 Then %>
	Welcome to the City of Montgomery's commemorative gift program. You now have the 
	ability to give a commemorative gift online to honor of a special someone in your 
	life. Birthdays, anniversaries, memorials, and other special occasions can be 
	chosen to honor a loved one. Please select a gift from below.  For more information, 
	please call (513) 792-8314.
<% End If %>
<% If orgid = 37 Then %>
	Welcome to the Park City's commemorative gift program. You now have the 
	ability to give a commemorative gift online to honor of a special someone in your 
	life. Birthdays, anniversaries, memorials, and other special occasions can be 
	chosen to honor a loved one. Please select a gift from below.  For more information, 
	please call (513) 681-4030.
<% End If %>
<% If orgid = 48 Then %>
	Welcome to the City of Hanahan's commemorative gift program. You now have the 
	ability to give a commemorative gift online to honor of a special someone in your 
	life. Birthdays, anniversaries, memorials, and other special occasions can be 
	chosen to honor a loved one. Please select a gift from below.  For more information, 
	please call (843) 554-4221.
<% End If %>
</div>


<P><% DisplayGifts(iorgid) %></P>

			
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
'  PUBLIC SUB DISPLAYGIFTS(BYVAL ORGID)
'--------------------------------------------------------------------------------------------------
Public Sub DisplayGifts(orgid)

    ' DISPLAY LIST OF FACILITIES
		sSQL = "select * from egov_gift where orgid = " & orgid
		Set oGift = Server.CreateObject("ADODB.Recordset")
		oGift.Open sSQL, Application("DSN"), 3, 1

		If NOT oGift.EOF Then

			' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
			Response.Write("<div class=""facilitylist gifts"">")
			Do While NOT oGift.EOF
					Response.Write vbcrlf & "<p><div class=""gift_header"">" & oGift("giftname") & " - <a href=""gift_form.asp?G=" & oGift("giftid") & """>Click here to purchase this Commemorative Gift</a></div>"
					'Response.Write "<div class=groupSmall4>" & vbCrLf
						Response.Write vbcrlf & "<table border=0 width=700><tr><td>"
							 
							 ' WRITE DESCRIPTION
								Response.Write vbcrlf & "<p class=""facilitydesc giftdesc"">" & replace(oGift("giftdescription"),"http://www.egovlink.com","https://www.egovlink.com") & "</p>"
								' WRITE LINK TO RESERVATION
								Response.Write vbcrlf & "<p class=""facilitydesc giftdesc"" align=""right""><b><a href=""gift_form.asp?G=" & oGift("giftid")& """>Click here to purchase this Commemorative Gift</a></b></p>"
			
						  Response.Write vbcrlf & "</td></tr></table>"
					'Response.Write "</div>" & vbCrLf
					
				
				
				
				' WRITE TITLE
				'Response.Write("<div class=""facilityname"">" & oGift("giftname") & "</div>" & vbCrLf)
				' WRITE LINK TO RESERVATION
				'Response.Write("<div class=""reserve_link"" align=""left""><a href=""gift_form.asp?G=" & oGift("giftid") & """>Purchase Commemorative Gift</a></br></div>" & vbCrLf)
				' WRITE DESCRIPTION
				'Response.Write("<div class=""facilitydesc"">" & oGift("giftdescription") & "</div>" & vbCrLf)
				' WRITE LINK TO RESERVATION
				'Response.Write("<div class=""reserve_link"" align=""right""><a href=""gift_form.asp?G=" & oGift("giftid")& """>Purchase Commemorative Gift</a></br><P></div>" & vbCrLf)
			
			
			oGift.MoveNext
			Loop
			Response.Write "</div>" & vbCrLf 

		End If

        ' CLOSE OBJECTS
        Set oGift = Nothing
End Sub
%>
