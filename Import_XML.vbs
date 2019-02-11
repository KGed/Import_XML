Sub Import_XML()
    Dim ows As Worksheet 'Output Worksheet
    Dim iws As Worksheet 'Input Worksheet
    Dim tab_name As String 'Where the tab name, from the import worksheet, is stored
    Dim dict As Scripting.Dictionary 'Dictionary of Tag Names'
    Set dict = New Scripting.Dictionary
    
    'Values used when creating the Dictionary
    Dim ows_range As Range, cell As Range
    
    Dim xCount As Long, yCount As Long
    Dim Node As IXMLDOMNode
    Dim xmlDom As MSXML2.DOMDocument60
    Set xmlDom = New MSXML2.DOMDocument60
    Set iws = ThisWorkbook.Sheets("Import XML Records")
    
    'File System Objects
    Dim MyFolder As folder 'Folder where the XML files are stored.
    Dim MyFile As File 'Each File that appears in the folder
    Dim MyFSO As FileSystemObject 'File System Object to open folder.
    Set MyFSO = New FileSystemObject
    
    'Collection of missing values
    Dim missing_dict As New Scripting.Dictionary
    Dim missing_fields As New Collection
    
    'Get the tab name from the import worksheet.
    tab_name = iws.Cells(2, 2).Value
    
    'Set the output worksheet based on the value for tab_name.
    Set ows = ThisWorkbook.Sheets(tab_name)
    
    'Create Dictionary of Values from Row 1 of the Output Worksheet (ows) (Cannot have any blank cells in the names, or it will throw an error)
    Set ows_range = ows.Range("A1", ows.Range("A1").End(xlToRight))
    For Each cell In ows_range
        dict.Add Key:=cell.Value, Item:=True
    Next cell
    
    'Get the folder path from the import worksheet (iws)
    strFolderPath = iws.Cells(1, 2).Value
    Set MyFolder = MyFSO.GetFolder(strFolderPath)
    
    'Prevent visuals from Excel
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Algorithm for copying nodes into Excel from XML files
    yCount = 2 'Start at Cell B2 (This is so we can include the filename in the first column)
    xCount = 2
    For Each MyFile In MyFolder.Files                           'For every file in the folder
      If InStr(MyFile.Name, ".xml") Then                        'If it's an XML file.
        xmlDom.Load (MyFile)                                    'Load the file
        Set Nodes = xmlDom.SelectSingleNode("ASSESSMENT")       'Select the main, parent tag.
        'Remove .xml from the filenames
        ows.Cells(yCount, 1).Value = Left(MyFile.Name, Len(MyFile.Name) - 4) 'Remove the file extension from the filename
        For Each Node In Nodes.ChildNodes                       'For every child node
            If dict.Exists(Node.BaseName) Then   'If the tag name exists in the dictionary
                If Node.BaseName = ows.Cells(1, xCount).Value Then
                    ows.Cells(yCount, xCount).Value = Node.Text
                Else
                    While (Node.BaseName <> ows.Cells(1, xCount).Value And ows.Cells(1, xCount) <> "") 'Skips
                        xCount = xCount + 1
                    Wend
                    ows.Cells(yCount, xCount).Value = Node.Text
                    xCount = xCount + 1
                End If
            Else
                If Not (missing_dict.Exists(Node.BaseName)) Then
                    missing_dict.Add Key:=Node.BaseName, Item:=True
                    missing_fields.Add (Node.BaseName)
                End If
            End If
        Next Node
        xCount = 2
        yCount = yCount + 1
      End If        'End If testing if file is an XML
    Next MyFile
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets(tab_name).Activate
    MsgBox "Completed Importing " & yCount - 2 & " Files"
    If missing_fields.Count <> 0 Then
        Dim output As String
        output = "Missing fields from Stress Tool:" & vbNewLine
        Dim i As Long
        For i = 1 To missing_fields.Count
            output = output & missing_fields.Item(i) & vbNewLine
            Next i
        MsgBox output
    Else
        MsgBox "No Missing Fields from Stress Tool"
    End If
End Sub



