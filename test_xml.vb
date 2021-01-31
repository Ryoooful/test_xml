Sub ko()

    
    Set xmlDoc = New MSXML2.DOMDocument60
    
    Dim myDic As Dictionary
    Set myDic = CreateObject("Scripting.Dictionary")
    myDic.Add "aaa", "bbb1"
    myDic.Add "ccc", "ddd1"
    
    Set myDic2 = CreateObject("Scripting.Dictionary")
    myDic2.Add "bbb2", "eee"
    myDic2.Add "ddd2", "fff"
    
    Dim xmlPI As IXMLDOMProcessingInstruction
    Set xmlPI = xmlDoc.appendChild(xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""Shift_JIS"""))
    
    Dim rootNode As IXMLDOMNode
    Set rootNode = xmlDoc.appendChild(xmlDoc.createNode(NODE_ELEMENT, "Root", ""))
    
    Call set_xmlDoc(rootNode, myDic, "AnalysisDoc")
    Call set_xmlDoc(rootNode, myDic2, "LayoutDoc")
    
    xmlDoc.Save ("C:\Users\Ryoooful\Desktop\aaa.xml")
End Sub

Sub set_xmlDoc(ByVal rootNode As IXMLDOMNode, ByVal myDic As Dictionary, ByVal tag_Name As String)
    Dim custNode As IXMLDOMNode
    Set custNode = rootNode.appendChild(rootNode.OwnerDocument.createNode(NODE_ELEMENT, tag_Name & "s", ""))
    
    Dim Var As stirng
    For Each k In myDic.Keys
        Set listNode = custNode.appendChild(rootNode.OwnerDocument.createNode(NODE_ELEMENT, tag_Name, ""))
        With listNode.appendChild(rootNode.OwnerDocument.createNode(NODE_ELEMENT, "Parent", ""))
            .Text = Var
        End With
        With listNode.appendChild(rootNode.OwnerDocument.createNode(NODE_ELEMENT, "Child", ""))
            .Text = myDic.Item(Var)
        End With
    Next
End Sub


Sub kokokoko()
    
    Set xmlDoc = New MSXML2.DOMDocument60
    xmlDoc.Load ("C:\Users\Ryoooful\Desktop\aaa.xml")
    
    Dim listNode As IXMLDOMNode
    Dim docNode As IXMLDOMNode
    
    Dim myDic As New Dictionary
    Dim objDict As Dictionary
    
    For Each listNode In xmlDoc.SelectSingleNode("//Root").ChildNodes
        For Each docNode In listNode.ChildNodes
            If Not myDic.Exists(docNode.BaseName) Then
                Set objDict = New Dictionary
                myDic.Add docNode.BaseName, objDict
            End If
            myDic.Item(docNode.BaseName).Add docNode.ChildNodes.Item(0).Text, docNode.ChildNodes.Item(1).Text
        Next
    Next
    
    
    For Each Var In myDic
        Set bbb = myDic.Item(Var)
        
        For Each ggg In bbb
            Debug.Print ggg & " : " & bbb.Item(ggg)
        Next
        
    Next Var
    
End Sub

