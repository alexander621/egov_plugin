<%
	'=======================================================================================
	' Title : PureAspGraph                                      
	' Author : Stig Christensen
	' Email : stig@stigc.dk
	' Website : www.stigc.dk
	' Changes :
	'	2004-06-01	Added new chars.
	'	2003-09-18	Description for each axes
	'	2003-09-16	Bug fix with Explorer 6.x. Wrong DOCTYPE/DTD
	'	2003-09-15	Added new vertical font, verdana 10px
	'	2003-09-14	Bug fixes
	'	2003-09-13	Valid HTML 4.01 Transitional + added CSS styles
	'	2003-09-11	Better support for database data
	'	2003-09-09	Added support for grouping. Added a label box
	'	2003-09-07	Added better support for horisontal graphs
	'=======================================================================================

Class PureAspGraph

	Private VERSION		
	
	Private lngValueSetCount
	Private arrItems()				'Array for the x-axes names
	Private arrValues()				'Array for the y-axes values
	Private arrLabels()				'Array for the text in the label box.
	Private arrColors()				'Array of 10 different colors
	Private dblDelta
	
	Private strTitle
	Private innerGraphHeight		'Height of the grey area. Do not change!
	Private outerGraphWidth			'Width of the graph.
	Private barWidth				'Width of the border.
	Private blnFlipText 			'Flip the text in X Axes?

	Private decimalCount 			'Numbers of decimals
	
	Private graphType				'0: horisontal, 1: vertical
	Dim strPicPath 					'Relative path for the images.
	Private MAX
	Private MIN
	Private MaxMinusMin
	
	Private strFont
	Private groupingCount			
	
	Private strYAxesTitle
	Private strXAxesTitle
	
	Private showValue 				'Show value for each bar?
	
	
	'DEFINES
	Private BAR_BORDER				'Border width on the bars
	Private BAR_SPACING				'Spacing between groups of bars
	Private TOPSPACING		
	Private BORDER					'Debug border.
	Private YCOUNT					'Count of numbers in the y-aks
	

   Private Sub Class_Initialize  
   
		VERSION = "0.38"
		
		groupingCount = 10
		
		Redim arrItems(0)
		Redim arrLabels(0)
		Redim arrValues(groupingCount-1, 0)
		Redim arrColors(groupingCount-1)
		
		arrColors(0)="#aa0011"
		arrColors(1)="#cc0011"
		arrColors(2)="#0000ff"
		arrColors(3)="#ffff00"
		arrColors(4)="#00ffff"
		arrColors(5)="#ff00ff"
		arrColors(6)="#ff8844"
		arrColors(7)="#4488ff"
		arrColors(8)="#888888"
		arrColors(9)="#ffffff"
		
		
		lngValueSetCount = 0
		strPicPath="pureAspGraph/"
		strFont="verdana10"
		decimalCount = 1
		innerGraphHeight=400
		outerGraphWidth=100
		barWidth=22
		blnFlipText=true
		
		YCOUNT=10
		BORDER=0
		TOPSPACING=7			
		graphType=0
		BAR_SPACING = 20
		BAR_BORDER = 1
   
   End Sub
 

	' ============================================================================
	' Public Functions
	' ============================================================================


	'Sets the data from 2 arrays.
	'x: Array of item texts
	'y: Array of item values
	Function addDataFromRecordset(rs, index)
		
		Dim i
		Do While Not rs.Eof
			arrValues(lngValueSetCount,i)=rs(index)
			rs.movenext
			i=i+1
		Loop
		lngValueSetCount = lngValueSetCount + 1

		'No need to flip text..
		if lngValueSetCount>4 Then
			blnFlipText = False
		End if
		
	End Function
	
	'Sets the data from a ADODB recordset.
	'rs: ADODB.recordset
	'f: index of item texts
	'd: index of item values
	Function setDataFromRecordset(rs, f, d)

		Dim i
		Do While Not rs.Eof
			Redim Preserve arrItems(i)
			Redim Preserve arrValues(groupingCount-1, i)
			arrItems(i)=rs(f)
			arrValues(lngValueSetCount , i)=rs(d)
			i=i+1
			rs.movenext
		Loop
		lngValueSetCount = lngValueSetCount + 1
		
	End Function

	'Sets the data from 2 arrays.
	'x: Array of item texts
	'y: Array of item values
	Function addData(x)
		
		Dim i
		For i=0 To ubound(x)
			arrValues(lngValueSetCount, i)= x(i)
		Next 
		lngValueSetCount = lngValueSetCount + 1
		
		'No need to flip text..
		if lngValueSetCount>4 Then
			blnFlipText = False
		End if
		
	End Function

		 		 
	'Sets the data from 2 arrays.
	'x: Array of item texts
	'y: Array of item values
	Function setData(x, y)
	
		Dim i
		Redim Preserve arrItems (ubound(x))
		Redim Preserve arrValues(groupingCount-1, ubound(y))
		For i=0 To ubound(y)
			arrItems(i) = x(i)
			arrValues(lngValueSetCount, i)= y(i)
		Next 
		lngValueSetCount = lngValueSetCount + 1
		
	End Function
	
	Function setTitle (v)
		strTitle=v
	End Function
	
	'Adds a title to the title box
	Function addLabel(v)
		
		Redim Preserve arrLabels(ubound(arrLabels)+1)
		arrLabels(ubound(arrLabels)-1) = v
		
	End Function

	Function setSize(v)
		If v=0 Then 
			innerGraphHeight=400
			YCOUNT=10
		ElseIf v=1 Then 
			innerGraphHeight=200
			YCOUNT=5
		End If
	End Function
	
	
	Function setShowValue(v)
		showValue = v
	End Function
	
	Function setBarColor(index, v)
		arrColors(index) = v
	End Function
	
	Function setYAxesTitle(v)
		strYAxesTitle = v
	End Function
	
	Function setXAxesTitle(v)
		strXAxesTitle = v
	End Function

	Function getVersion()
		getVersion = VERSION
	End Function

	Function setBarBorder(v)
		BAR_BORDER = v
	End Function

	Function setBarSpacing(v)
		BAR_SPACING = v
	End Function			
	 		
	Function setPicPath(v)
		strPicPath = v
	End Function
		
	Function setBarWidth(v)
		barWidth = v
	End Function
	
	Function setFlipText(v)
		blnFlipText=v
	End Function
		
	Function setType(v)
		graphType=v
	End Function

	Function setFont(v)
		strFont=v
	End Function	
				 			 


	
	' Prints the graph
	Function print()
		
		Call initPrint() 
		
		response.write "<!-- PureAspGraph " & VERSION & ". www.stigc.dk -->" & vbNewLine
		response.write "<table border=0 cellpadding=10 style=""border: 1px solid #000000;""><tr><td>" & vbNewLine
		
		If graphType=0 Then
			Call printHorizontal()
		Else
			Call printVertical()
		End If
		
		response.write "</td></tr></table>" & vbNewLine
	
	End Function
	

	' ============================================================================
	' Private Functions
	' ============================================================================


	Private Function validProperty(v)
		v = cLng(v)
		if v<0 Then v=0
		validProperty = v
	End Function
	
	
	

	'Finds the appropiate MAX y value
	Private Function findMaxY (byval maxValueY)
	
		Dim d
		d=1
		Do While maxValueY>99
			d=d*10
			maxValueY=maxValueY/10
		Loop	
		findMaxY = cLng(maxValueY+1)*d
		
		if (findMaxY MOD YCOUNT=0) Then decimalCount=0 'No need for decimals
		
	End Function
	
	
	'Calculates som importen values
	Private Function initPrint() 
		
		dim i, j
		MAX = -9999999
		MIN = 9999999
		
		'Finds the MAX and MIN value
		For i=0 To lngValueSetCount
			For j=0 to ubound(arrItems)
				if arrValues(i,j)>MAX Then MAX=arrValues(i,j)
				if arrValues(i,j)<MIN Then MIN=arrValues(i,j)
			Next 
		Next 

		MAX=findMaxY(MAX)
		dblDelta = innerGraphHeight / MAX
		MaxMinusMin = MAX - 0

	End Function		
	
	
	'Prints the color description box
	Private Function printLabelBox ()
		
		Dim i 
		
		If ubound(arrLabels)>0 Then
			response.write "<table border=0 width=100 cellpadding=10 style=""height:50px; border: 1px solid #000000;"">"
			response.write "<tr>"
			response.write "<td>"
			
			For i=0 To ubound(arrLabels)-1
				response.write "<table cellpadding=0 cellspacing=0><tr><td bgColor=""" & arrColors(i) & """>"
				response.write "<img alt="""" border=1 src=""" & strPicPath  & "images/blank.gif"" width=14 height=14>"
				response.write "</td><td>&nbsp;&nbsp;</td><td nowrap class=""pureAspGraphLabelText"">" & arrLabels(i) & "</td></tr></table>"
				if i<>ubound(arrLabels)-1 Then response.write "<br>"
			Next
			
			response.write "</td>"
			response.write "</tr>"
			response.write "</table>"
		End if
		
	End Function
	

	'Prints a horizontal bar chart
	Private Function printHorizontal()
		
		Dim yMin, i, g 
		yMax = 0
		
		if strTitle<>"" Then
			response.write "<div align=center class=""pureAspGraphTitleText"">" & strTitle & "</div>"
			response.write "<br>"
		End if
			
		response.write "<table border=" & BORDER & " cellpadding=0 cellspacing=0 width=""" & outerGraphWidth & """ style=""height:100%"">"
		
		
		if strYAxesTitle<>"" Then
			response.write "<tr><td rowspan=4>"
			response.write flipText(strYAxesTitle, true)
			response.write "</td></tr>"
			response.write "<tr><td rowspan=4>"
			response.write "<img alt="""" border=0 src=""" & strPicPath  & "images/blank.gif"" width=10>"
			response.write "</td></tr>"
		End if
		
		response.write "<tr valign=top>"
		response.write "<td>"
		
		response.write "<table cellpadding=0 cellspacing=0 border=" & BORDER & " style=""height:100%"">"
			response.write "<tr style=""height:1px;""><td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & BAR_SPACING & "></td></tr>"
			For i=0 to ubound(arrItems)
				response.write "<tr>"
					response.write "<td nowrap class=""pureAspGraphYaxesText"">" & arrItems(i) & "&nbsp;&nbsp;</td>"
				response.write "</tr>" & vbNewLine
				
				If lngValueSetCount>1 And i<ubound(arrItems) Then
					response.write "<tr><td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & BAR_SPACING & "></td></tr>"
				End If
			Next 
			response.write "<tr style=""height:1px;""><td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & BAR_SPACING & "></td></tr>"
		response.write "</table>"
		
		response.write "</td>"
		response.write "<td>"
		
		
		response.write "<table border=0 style=""height:100%; background-image: url(" & strPicPath & "images/gitterHorisontal.gif);"" border=" & BORDER & " cellpadding=0 cellspacing=0 width=""400"">"
		response.write "<tr style=""height:1px;""><td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & BAR_SPACING & "></td></tr>"
			For i=0 to ubound(arrItems)
				response.write "<tr>"
					response.write "<td align=""left"">"
						For g=0 to lngValueSetCount-1
							if showValue Then response.write "<table cellpadding=0 cellspacing=0><tr><td>"
							'No background color if value=0
							thisBarValue = validProperty(dblDelta*arrValues(g, i)-2)
							If thisBarValue>0 Then
								response.write "<table cellpadding=0 cellspacing=0 bgColor=""" & arrColors(g) & """><tr><td>"
								response.write "<img alt="""" title=""" & arrValues(g, i) & """ border=" & BAR_BORDER & " src=""" & strPicPath  & "images/blank.gif"" height=" & barWidth-2 & " width=" & thisBarValue & ">"
							Else
								response.write "<table cellpadding=0 cellspacing=0><tr><td>"
								response.write "<img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & barWidth-2 & " width=0>"
							End if
							response.write "</td></tr></table>"
							if showValue Then response.write "</td><td>&nbsp;" & arrValues(g, i) & "</td></tr></table>"
						Next 
					response.write "</td>"
				response.write "</tr>" & vbNewLine
				if lngValueSetCount>1 AND  i<ubound(arrItems) Then
					response.write "<tr><td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=20></td></tr>"
				End if	
			Next 
		response.write "<tr style=""height:1px;""><td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & BAR_SPACING & "></td></tr>"
		response.write "</table>"	
		
		response.write "</td>"
		response.write "<td valign=""middle"">"
		%>
						<table>
							<tr>
								<td>
									<img alt="" border=0 src="<%=strPicPath%>images/blank.gif" width=10>
								</td>
								<td>
									<%
										Call printLabelBox ()
									%>
								</td>
							</tr>
						</table>
		<%
		response.write "</td>"
		response.write "</tr>"
		
		
		response.write "<tr>"
		response.write "<td></td>"
		response.write "<td colspan=2>"
		
			response.write "<table border=" & BORDER & " cellpadding=0 cellspacing=0 width=430><tr><td>"
			response.write "<img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=1 width=30>"
			response.write "</td><td>"
				response.write "<table cellpadding=0 cellspacing=0 border=0><tr valign=top>"
					For i=1 to YCOUNT
						yMax = yMax + MaxMinusMin/YCOUNT
						response.write "<td class=""pureAspGraphXaxesText"" width=40>"
						response.write "<img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=5 width=1><br>"
						response.write FlipText(FormatNumber(yMax, decimalCount,0,0,0), true)
						response.write "</td>"
					Next 
				response.write "</tr></table>"
			response.write "</td></tr></table>"
			
		response.write "</td>"
		response.write "</tr>"


		response.write "</table>"	

		if strXAxesTitle<>"" Then
			response.write "<br><div align=center class=""pureAspGraphXaxesDescription"">" & strXAxesTitle & "</div>"
		End if
					
	End Function
	
	
	'Prints a vertical bar chart
	Private Function printVertical()

			Dim yMax, i, g
			yMax = MAX
			
			if strTitle<>"" Then
				response.write "<div align=center class=""pureAspGraphTitleText"">" & strTitle & "</div>"
			End if
			%>
			
			<table border="<%=BORDER%>" width="<%=outerGraphWidth%>" cellpadding="0" cellspacing="0">
				<% If strYAxesTitle<>"" Then %>
				<tr>
					<td rowspan=5>
						<%=FlipText(strYAxesTitle, True)%>
					</td>
				</tr>
				<tr>
					<td rowspan=5><img alt="" src="<%=strPicPath%>images/blank.gif" height="4" width="4"></td>
				</tr>
				<% End If %>
				<tr valign="top">
					<td rowspan="2" width="1">
						<table border=<%=BORDER%> cellpadding=0 cellspacing=0>
							<%
								For i=1 to YCOUNT
									response.write "<tr>"
										response.write "<td nowrap valign=""top"" align=""right"" class=""pureAspGraphYaxesText"">"
											response.write FormatNumber(yMax, decimalCount,0,0,0) & " >"
											yMax = yMax - MaxMinusMin/YCOUNT
										response.write "</td>"
										if i=1 Then
											response.write "<td rowspan=" & YCOUNT+1 & "><img alt="""" src=""" & strPicPath  & "images/blank.gif"" height=" & innerGraphHeight & " width=8></td>"
										End if
									response.write "</tr>" & vbNewLine
								Next 
							%>
						</table>
						<img alt="" src="<%=strPicPath%>images/blank.gif" border="0" height=<%=TOPSPACING%> width="8">
					</td>
					<td><img alt="" src="<%=strPicPath%>images/blank.gif" height="<%=TOPSPACING%>" width="8"></td>
					<td></td>
				</tr>
				<tr valign=top>
					<td>
						<%
							'Prints the charts.
							response.write "<table style=""background-image: url(" & strPicPath & "images/gitter.gif);"" width=""100%"" border=0 cellpadding=0 cellspacing=0>"
								response.write "<tr valign=bottom>"
								response.write "<td width=1><img alt="""" border=0 src=""" & strPicPath  & "images/blank.gif"" height=" & innerGraphHeight & " width=" & BAR_SPACING & "></td>" 'Mellemrum
									For i=0 to ubound(arrItems)
										response.write "<td align=""center"">"
										
											response.write "<table cellpadding=0 cellspacing=0><tr valign=bottom>"
												for g=0 to lngValueSetCount-1
													response.write "<td>"
													if showValue Then response.write Fliptext(arrValues(g,i), true) & "<br style=""font-size:5px;"">"
													thisBarValue = validProperty(dblDelta*arrValues(g,i)-(BAR_BORDER*2))
													if thisBarValue>0 Then
														response.write "<table border=0 cellpadding=0 cellspacing=0 bgColor=""" & arrColors(g) & """><tr><td>"
														response.write "<img alt="""" border=" & BAR_BORDER & " title=" & arrValues(g,i) & " src=""" & strPicPath  & "images/blank.gif"" width=" & barWidth-(BAR_BORDER*2) & " height=" & thisBarValue & ">"
													Else
														response.write "<table border=0 cellpadding=0 cellspacing=0 ><tr><td>"
														response.write "<img alt="""" border=0 title=" & arrValues(g,i) & " src=""" & strPicPath  & "images/blank.gif"" width=" & barWidth-(BAR_BORDER*2) & " height=1>"
													End If
													response.write "</td></tr></table>" & vbNewLine
													response.write "</td>"
												'if showValue Then response.write "<table cellpadding=0 cellspacing=0><tr><td>"
												'if showValue Then response.write "</td><td>&nbsp;" & arrValues(g, i) & "</td></tr></table>"

												Next 
											response.write "</tr></table>"
											
										response.write "</td>" & vbNewLine
										if i< ubound(arrItems) And lngValueSetCount>1 And BAR_SPACING>0 Then
											response.write "<td><img alt="""" src=""" & strPicPath  & "images/blank.gif"" width=" & BAR_SPACING & "></td>" & vbNewLine
										End if
									Next 
								response.write "<td width=1><img alt="""" border=0 src=""" & strPicPath  & "images/blank.gif"" width=" & BAR_SPACING & "></td>" 'Mellemrum
								response.write "</tr>"
							response.write "</table>"	
						%>
					</td>
					<td valign="middle">
						<table cellpadding=0 cellspacing=0>
							<tr>
								<td>
									<img alt="" border=0 src="<%=strPicPath%>images/blank.gif" width=10>
								</td>
								<td>
									<%
										Call printLabelBox ()
									%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr valign=top>
					<td></td>
					<td>
						<%
							Dim dblPercent, barSpacing
							barSpacing = cLng(BAR_SPACING/2)
							if lngValueSetCount=1 Then barSpacing=BAR_SPACING
							
							dblPercent = cLng(100/(ubound(arrItems)+1))
							'response.write dblPercent
							'Print the texts in the x-axes
							response.write "<table width=""100%"" border=0 cellpadding=0 cellspacing=0>"
								response.write "<tr valign=top>"
									response.write "<td width=1><img alt="""" border=0 src=""" & strPicPath  & "images/blank.gif"" width=" & barSpacing & "></td>" & vbNewLine 'Mellemrum
										For i=0 to ubound(arrItems)
											response.write "<td width=""" & dblPercent & "%"" align=center class=""pureAspGraphXaxesText"">"
												response.write FlipText(arrItems(i), blnFlipText)
												response.write "<img alt="""" border=0 src=""" & strPicPath  & "images/blank.gif"" width=" & barWidth-(BAR_BORDER*2) & ">"
											response.write "</td>" & vbNewLine
										Next 
									response.write "<td width=1><img alt="""" border=0 src=""" & strPicPath  & "images/blank.gif"" width=" & barSpacing & "></td>" 'Mellemrum
								response.write "</tr>"		
							response.write "</table>"	
						%>
					</td>
					<td></td>
				</tr>
			</table>
<%

			if strXAxesTitle<>"" Then
				response.write "<div align=center class=""pureAspGraphXaxesDescription"">" & strXAxesTitle & "</div>"
			End if


	End Function
	
	'Prints the text for the x-axes.
	Private Function FlipText(s, blnFlip)

			
		If s="" Then Exit Function
		
		dim i, r, c, width
		for i=1 to len(s)
			c = Mid(s,i,1)
			if blnFlip Then
				if c ="." Then c ="dot"
				if c =" " Then c ="space"
				if c ="," Then c ="comma"
				if c ="/" Then c ="slash"
				if c ="-" Then c ="line"
				if c ="(" Then c ="left_p"
				if c =")" Then c ="right_p"
				if c ="*" Then c ="star"
				if c =":" Then c ="colon"
				if lcase(c) ="æ" Then c ="ae"
				if lcase(c) ="ø" Then c ="oe"
				if lcase(c) ="å" Then c ="aa"
				
				
				if strFont="verdana10" Then
					width = "width=15"
				Else
					width = "width=17"
				End if
				
				If c=lcase(c) Then
					r = "<img alt="""" " & width & " src=""" & strPicPath  & "images/" & strFont & "/" & c & ".gif""><br>" & r & vbNewLine
				Else
					r = "<img alt="""" " & width & " src=""" & strPicPath  & "images/" & strFont & "/capitals/" & c & ".gif""><br>" & r & vbNewLine					
				End If
			Else
				r = r & c
			End If
		Next
		FlipText = r
	End Function
	
End Class

%>