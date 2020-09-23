VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "XML Test"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3270
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTextReturned 
      Height          =   285
      Left            =   330
      TabIndex        =   2
      Top             =   1350
      Width           =   2625
   End
   Begin VB.TextBox txtToSend 
      Height          =   315
      Left            =   330
      TabIndex        =   1
      Top             =   600
      Width           =   2625
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   1740
      TabIndex        =   0
      Top             =   1890
      Width           =   1245
   End
   Begin VB.Label lblTextReturned 
      Caption         =   "Text Returned"
      Height          =   255
      Left            =   330
      TabIndex        =   4
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label lblTextToSend 
      Caption         =   "Text to Send"
      Height          =   255
      Left            =   330
      TabIndex        =   3
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSend_Click()

'Dim objXMLWrapper As New RPC.MethodWrapper
Dim objXMLRequest As New MSXML.XMLHTTPRequest
Dim objXMLDOM     As New MSXML.DOMDocument
Dim objRoot       As MSXML.IXMLDOMElement
Dim objRetValNode As MSXML.IXMLDOMNode
Dim xmlString     As String
'Redim Params Array
'ReDim Params(0)

Me.MousePointer = vbHourglass

'Create XML String
xmlString = "<Obj_Call>" & vbCrLf
xmlString = xmlString & "<Object>" & vbCrLf
xmlString = xmlString & "<Name>SoapDest.cStringTest</Name>" & vbCrLf
xmlString = xmlString & "<Proc>StringReverse</Proc>" & vbCrLf
xmlString = xmlString & "<Params>" & vbCrLf
xmlString = xmlString & "<Param>" & Trim(txtToSend.Text) & "</Param>"
xmlString = xmlString & "</Params>"
xmlString = xmlString & "</Object>" & vbCrLf
xmlString = xmlString & "</Obj_Call>"
    
'Send XML to Server
objXMLRequest.open "GET", "http://216.32.32.175/rpcxml.asp?Request=" & xmlString, False
objXMLRequest.send

'Check Status from Server
If objXMLRequest.Status = 200 Then 'Successful Response
    Loaded = objXMLDOM.loadXML(objXMLRequest.responseText)
    If Loaded Then
        Set objRoot = objXMLDOM.documentElement
        Set objRetValNode = objRoot.selectSingleNode(".//Response")
        txtTextReturned.Text = objRetValNode.Text
    End If
End If
   
Set objXMLRequest = Nothing
Set objXMLDOM = Nothing
Set objRoot = Nothing
Set objRetValNode = Nothing
 
Me.MousePointer = vbDefault

End Sub
