VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private strXML As String

Public Function MethodToExec(objName As String, Method As String, Params() As Variant)

Dim strObjName   As String
Dim strObjMethod As String
Dim vParamList   As Variant
Dim strResponse  As String

'Parse Params
For x = 0 To UBound(Params) - 1
    If x = 0 Then
        vParamList = Params(x)
    Else
        vParamList = vParamList & ", " & Params(x)
    End If
Next

Set objToCall = CreateObject(objName)

If IsObject(objToCall) Then
    strResponse = CallByName(objToCall, Method, VbMethod, vParamList)
    MethodToExec = "<xml><Response>" & strResponse & "</Response></xml>"
End If

End Function

Public Property Let XMLToParse(ByVal XMLString As Variant)

strXML = CStr(XMLString)

End Property

Public Function ResponseXML() As String

Dim objXMLDOM       As New MSXML.DOMDocument
Dim objRoot         As MSXML.IXMLDOMElement
Dim objNameNode     As MSXML.IXMLDOMNode
Dim objMethodNode   As MSXML.IXMLDOMNode
Dim objParamList    As MSXML.IXMLDOMNodeList
Dim Loaded          As Boolean
Dim strObjectName   As String
Dim strObjectMethod As String
Dim Params()        As Variant

ReDim Params(0)

Loaded = objXMLDOM.loadXML(Trim(strXML))

If Loaded = True Then
    'Load XML Nodes
    Set objRoot = objXMLDOM.documentElement
    Set objNameNode = objRoot.selectSingleNode(".//Name")
    Set objMethodNode = objRoot.selectSingleNode(".//Proc")
    Set objParamList = objRoot.selectNodes(".//Params")
    
    'Assign Node Values to Variables
    strObjectName = objNameNode.Text
    strObjectMethod = objMethodNode.Text
    For Each Node In objParamList
        ReDim Preserve Params(UBound(Params) + 1)
        Params(UBound(Params) - 1) = Node.Text
    Next
    'Call Object
    ResponseXML = MethodToExec(strObjectName, strObjectMethod, Params())
Else
    ResponseXML = "<xml><Error>Could not load XML Request</Error></xml>"
End If
    
Set objXMLDOM = Nothing
Set objRoot = Nothing
Set objNameNode = Nothing
Set objMethodNode = Nothing
Set objParamList = Nothing

End Function

