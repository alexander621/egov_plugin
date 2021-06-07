<%
Function URLDecode( encodedstring )
	Dim strIn, strOut, intPos, strLeft
	Dim strRight, intLoop

	strIn  = encodedstring
  strOut = ""
  intPos = Instr(strIn, "+")

	Do While intPos
		strLeft = ""
    strRight = ""

		If intPos > 1 Then strLeft = Left(strIn, intPos - 1)
		If intPos < len(strIn) Then strRight = Mid(strIn, intPos + 1)

		strIn = strLeft & " " & strRight
		intPos = InStr(strIn, "+")
		intLoop = intLoop + 1
	Loop

	intPos = InStr(strIn, "%")

	Do while intPos
		If intPos > 1 then _
			strOut = strOut & Left(strIn, intPos - 1)
		  strOut = strOut & Chr(clng("&H" & Mid(strIn, intPos + 1, 2)))
		If intPos > (len(strIn) - 3) then
			strIn = ""
		Else
			strIn = Mid(strIn, intPos + 3)
		End If
		intPos = InStr(strIn, "%")
	Loop

	URLDecode = strOut & strIn
End Function
%>
