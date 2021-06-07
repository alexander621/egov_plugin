<%
'Option Explicit

'----------------------------------------
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'----------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'----------------------------------------
'  Make buffer Quoted 'safe' 
'  Useful in building lines like:
'     "<INPUT TYPE=Hidden VALUE='" & HTMLQuotes(strValue) & "'>"
'----------------------------------------
Function HTMLQuotes( strHTML )
  HTMLQuotes = Replace(    strHTML, Chr(38), "&#38;" )  ' Ampersand
  HTMLQuotes = Replace( HTMLQuotes, Chr(34), "&#34;" )  ' Double Quotes
  HTMLQuotes = Replace( HTMLQuotes, Chr(39), "&#39;" )  ' Single Quotes
End Function

'----------------------------------------
'  Change a CarriageReturn to a "<br>"
'----------------------------------------
Function CRsafe( strCR )
  If Not VarType( strCR ) = vbString Then CRsafe = strCR: Exit Function
  CRsafe = Replace( strCR, Chr(13), "<br>" )
End Function

'----------------------------------------
'  Make buffer HTML 'safe'
'----------------------------------------
Function HTMLsafe( strHTML )
  If VarType( strHTML ) <> vbString Then HTMLsafe = strHTML : Exit Function
  If Len( strHTML ) < 1 Then HTMLsafe = "" : Exit Function
  HTMLsafe = Replace(strHTML, "<br>", "")
End Function
  
'----------------------------------------
'  Make buffer MIME 'safe'
'----------------------------------------
Function MIMESafe( strMIME )
  If VarType( strMIME ) <> vbString Then MIMESafe = strMIME : Exit Function
  MIMESafe = Replace( strMIME, "%", "%25" )
  MIMESafe = Replace( strMIME, "$", "%24" )
  MIMESafe = Replace( strMIME, "&", "%26" )
  MIMESafe = Replace( strMIME, "!", "%21" )
  MIMESafe = Replace( strMIME, "+", "%2B" )
  MIMESafe = Replace( strMIME, " ", "+" )
End Function
  
'----------------------------------------
'  Make buffer URL 'safe'
'----------------------------------------
Function URLSafe( strURL )
  If VarType( strURL ) <> vbString Then URLSafe = strURL : Exit Function
  URLSafe = Replace( strURL, " ", "%20" )
End Function
%>