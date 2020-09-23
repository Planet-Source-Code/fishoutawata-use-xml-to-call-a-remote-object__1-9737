<%@Language="VBScript"%>

<%
'Request String
Dim Request
'XML Wrapper Object
Dim objRPCWrapper
'XML Objects
Dim objDOM
Dim objRoot
Dim objNameNode
Dim objMethodNode
Dim objParamList
'Node Value Vars
Dim ObjectName
Dim ObjectMethod
Dim Params()
'Object to be Called
Dim objToCall
'Redim Array
Redim Params(0)

'Initialize XML Wrapper Object
Set objRPCWrapper = Server.CreateObject("RPC.MethodWrapper")
'Initialize XML DOM Object
Set objDOM = Server.CreateObject("Microsoft.XMLDOM")

'Get Request From Client
Request = Request("Request")

'Load XML Document
Loaded = objDOM.LoadXML(Trim(Request))

'Parse XML
If Loaded = True Then
	'Load XML Nodes
	Set objRoot = objDOM.documentElement
	Set objNameNode = objRoot.selectSingleNode(".//Name")
	Set objMethodNode = objRoot.selectSingleNode(".//Proc")
	Set objParamList = objRoot.selectNodes(".//Params)
	
	'Assign Node Values to Variables
	ObjectName = objNameNode.Text
	ObjectMethod = objMethodNode.Text
	For Each Node in objParamList
		ReDim Preserve Params(Ubound(Params) + 1)
		Params(Ubound(Params) - 1) = Node.Text
	Next

	'Execute Object
	retval = objRPCWrapper.MethodToExec(cstr(ObjectName), cstr(ObjectMethod), Params())	
	'Return Value to Client
	Response.Write retval
End If

set objDOM = Nothing
Set objRPCWrapper = Nothing
Set objRoot = Nothing
Set objNameNode = Nothing
Set objMethodNode = Nothing
Set objParamList = Nothing
					
%>

