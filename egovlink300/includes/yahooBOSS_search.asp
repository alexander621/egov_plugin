<%
'/// <page>
'///	<created>05/11/2003</created>
'///	<author>Chris Surfleet</author>
'///	<summary>Contains common functions for working with XML files.</summary>
'/// </page>

'/// <summary>Returns an XSL Style Sheet</summary>
'/// <param name="strStylesheetLocation" type="String">The location of the XSL Sytlesheet</param>
'/// <returns type="MSXML2.DomDocument.4.0">The XSL stylesheet</returns>
Function GetXslStyleSheet(ByVal strStyleSheetLocation)
	Set GetXslStyleSheet = GetXslStyleSheetWithParams(strStyleSheetLocation, "")
End Function

'/// <summary>Returns an XSL Style Sheet, with Parameters passed in.</summary>
'/// <param name="strStylesheetLocation" type="String">The location of the XSL Stylesheet</param>
'/// <param name="strParams" type="String">
'///	Parameters to be passed in, in the format "param=value;param=value;" e.t.c
'/// </param>
'/// <returns type="MSXML2.DomDocument.4.0">The XSL stylesheet</returns>
Function GetXslStyleSheetWithParams(ByVal strStylesheetLocation, ByVal strParams)	
	Dim xslStyleSheet 'As MSXML2.DomDocument.4.0
	Dim strParamArray 'As String
	Dim strParam 'As String
	Dim intCount 'As Integer
	
	strParamArray = Split(strParams, ";")
	
	Set xslStyleSheet = GetXmlDocument(strStylesheetLocation)
	xslStyleSheet.setProperty "SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'"
	
	For intCount = 0 To UBound(strParamArray)
		If InStr(strParamArray(intCount), "=") > 0 Then
			strParam = Split(strParamArray(intCount), "=")			
			xslStyleSheet.selectSingleNode("//xsl:variable[@name='" & strParam(0) & "']/@select").text = "'" & strParam(1) & "'"
		End If
	Next 'intCount
	
	Set GetXslStyleSheetWithParams = xslStyleSheet
	
	Set xslStyleSheet = Nothing
End Function

'/// <summary>Gets an XML Document</summary>
'/// <param name="strDocumentLocation" type="String">The location of the Document</param>
'/// <returns type="MSXML2.DomDocument.4.0">The XML Document</returns>
Function GetXmlDocument(ByVal strDocumentLocation)
	Dim xmlDocument 'As MSXML2.DomDocument.4.0
	
	Set xmlDocument = Server.CreateObject("MSXML2.DomDocument.4.0")
	xmlDocument.async = False
	xmlDocument.load Server.MapPath(strDocumentLocation)
	
	Set GetXmlDocument = xmlDocument
	
	Set xmlDocument = Nothing
End Function

'/// <summary>Returns the supplied XML Document, parsed by the supplied XSL stylesheet</summary>
'/// <param name="strDocumentLocation" type="String">The location of the XML Document</param>
'/// <param name="strStylesheetLocation" type="String">The location of the XSL Sytlesheet</param>
'/// <returns type="MSXML2.DomDocument.4.0">The Resulting Document</returns>
Function GetXmlDocumentByStyleSheet(ByVal strDocumentLocation, ByVal strStyleSheetLocation) 'As MSXML2.DomDocument.4.0
	Set GetXmlDocumentByStyleSheet = GetXmlDocumentByStyleSheetWithParams(strDocumentLocation, strStyleSheetLocation, "")
End Function

'/// <summary>Returns the supplied XML Document, parsed by the supplied XSL stylesheet</summary>
'/// <param name="strDocumentLocation" type="String">The location of the XML Document</param>
'/// <param name="strStylesheetLocation" type="String">The location of the XSL Sytlesheet</param>
'/// <param name="strParams" type="String">
'///	Parameters to be passed in, in the format "param=value;param=value;" e.t.c
'/// </param>
'/// <returns type="MSXML2.DomDocument.4.0">The Resulting Document</returns>
Function GetXmlDocumentByStyleSheetWithParams(ByVal strDocumentLocation, ByVal strStyleSheetLocation, ByVal strParams) 'As MSXML2.DomDocument.4.0
	Dim xmlDocument 'As MSXML2.DomDocument
	
	Set xmlDocument = Server.CreateObject("MSXML2.DomDocument.4.0")
	
	GetXmlDocument(strDocumentLocation).transformNodeToObject GetXslStyleSheetWithParams(strStyleSheetLocation, strParams), xmlDocument
		
	Set GetXmlDocumentByStyleSheetWithParams = xmlDocument
	
	Set xmlDocument = Nothing
End Function
%>

