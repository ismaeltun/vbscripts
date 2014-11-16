Option Explicit

If WScript.Arguments.Count = 0 Then
  WScript.Echo "A script to prettify xml files. " & vbCrLf & _
        "usage: cscript " & WScript.ScriptName & " [XML File To Format]" 
  WScript.Quit
End If 

Dim strInputFile: strInputFile = WScript.Arguments(0)

Dim fso : Set fso = WScript.CreateObject("Scripting.FileSystemObject")

Dim ip: Set ip = fso.OpenTextFile(strInputFile, 1, False, -2)
Dim strXML: strXML = ip.ReadAll
strXML = Replace(strXML,"><",">" & vbCrLf & "<")
ip.Close

Dim op: Set op = fso.CreateTextFile(strInputFile, True, False)
op.Write strXML
op.Close

Dim strXSL: strXSL = _
    "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
    "<xsl:output method=""xml"" indent=""yes""/>" & _
    "<xsl:template match=""/"">" & _
    "<xsl:copy-of select="".""/>" & _
    "</xsl:template>" & _
    "</xsl:stylesheet>"

Dim xslDom : Set xslDom = WScript.CreateObject("Msxml2.DOMDocument")
xslDom.async = False
xslDom.loadXML strXSL

Dim xmlDom : Set xmlDom = WScript.CreateObject("Msxml2.DOMDocument")
xmlDom.async = False
xmlDom.load strInputFile
xmlDom.transformNode xslDom
xmlDom.save strInputFile
