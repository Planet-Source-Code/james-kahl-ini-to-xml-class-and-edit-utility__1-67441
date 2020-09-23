VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************************
'Module:        mCONFIG - Class Module
'Filename:      mCONFIG.cls
'Author:        Jim Kahl
'Based On:      code in this class is based on two applications
'               cINIFile by Steve McMahon www.vbaccelerator.com
'
'               [[ a ini file to XML converter by Pamela RAI
'               www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=48510&lngWId=1
'               this code uses a modified version of that xml format
'
'Purpose:       to be able to read/write settings to an XML file the same way that you
'               can read/write to an INI file
'Instancing:    5 - Multiuse
'Depends:       Microsoft XML, v2.6 or later - this references v2.6 but has also been
'               tested with v6 and no problems
'
'Assumes:       if you are attempting to open an existing XML file it must be in the
'               correct format - see sample.ini.xml in Related Documents for the utility
'               project for the format used
'
'NOTE:          this was written so that it could almost be a direct drop in replacement
'               of Steve McMahon's cINIFile.cls with some omissions/additions on my part
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function GetPrivateProfileSection Lib "kernel32.dll" _
        Alias "GetPrivateProfileSectionA" ( _
                ByVal lpAppName As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) _
                As Long
                
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" _
        Alias "GetPrivateProfileSectionNamesA" ( _
                ByVal lpszReturnBuffer As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) _
                As Long

'****************************************************************************************
'CONSTANTS - PRIVATE
'****************************************************************************************
Private Const INITIAL_BUFFER_SIZE As Long = 1024
Private Const mCONFIG As String = "Configuration"
Private Const mSECTION As String = "Section"
Private Const mBLOCK As String = mCONFIG & "/" & mSECTION
Private Const mKEY As String = "Setting"
Private Const mAttrName As String = "Name"
Private Const mAttrValue As String = "Value"
Private Const mAttrDesc As String = "Description"
Private Const mAttrNameID As Long = 0
Private Const mAttrValueID As Long = 1
Private Const mAttrDescID As Long = 2

'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private msFilename As String
Private msSection As String
Private msKey As String
Private moDoc As New MSXML2.DOMDocument

'****************************************************************************************
'PROPERTIES - PUBLIC READ/WRITE
'****************************************************************************************
'Path - sets/returns the fully qualified path and filename of the xml file
Public Property Get Path() As String
    Path = msFilename
End Property

Public Property Let Path(ByRef sFilename As String)
    msFilename = sFilename
    moDoc.Load sFilename
End Property

'*******************************************
'Section - sets/returns the section of the XML file
Public Property Get Section() As String
    Section = msSection
End Property

Public Property Let Section(ByRef sSection As String)
    msSection = sSection
End Property

'*******************************************
'Key - sets/returns the key within a section of the XML file
Public Property Get Key() As String
    Key = msKey
End Property

Public Property Let Key(ByRef sKey As String)
    msKey = sKey
End Property

'*******************************************
'Value - sets/returns the value of a specific key within a section
Public Property Get Value() As String
    On Error GoTo ErrHandler
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)

    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            Set oNodes = oBlock.childNodes

            If oNodes.length <> 0 Then
                For Each oNode In oNodes
                    If oNode.Attributes(mAttrNameID).Text = msKey Then
                        Value = oNode.Attributes(mAttrValueID).Text
                        Exit Property
                    End If
                Next oNode
            End If
            Exit For
        End If
    Next oBlock

ExitProc:
    Exit Property
ErrHandler:
    Resume ExitProc
End Property

Public Property Let Value(ByRef sValue As String)
    'Note:   if the section and/or key do not exist this routine will call routines
    '        to create them
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNodeChild As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMNode
    
    'make sure the document has the Configure node and at least one Section
    'before attempting to write
    Set oBlocks = moDoc.selectNodes(mCONFIG)

    If oBlocks.length = 0 Then
        CreateConfigure
    End If
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)
    
    If oBlocks.length = 0 Then
        CreateSection
        Value = sValue
    End If
    
    'iterate through the blocks until we find the correct section and key
    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            Set oNodes = oBlock.childNodes

            If oNodes.length <> 0 Then
                For Each oNode In oNodes
                    If oNode.Attributes(mAttrNameID).Text = msKey Then
                        'set the key value and save the file
                        oNode.Attributes(mAttrValueID).Text = sValue
                        moDoc.save msFilename
                        Exit Property
                    End If
                Next oNode
                
                'if we reach this point then the section exists but the key does not
                'so we create the new key and value pair
                CreateKeyValue oBlock, sValue
                moDoc.save msFilename
                Exit Property
            End If
            Exit For
        End If
    Next oBlock
    
ExitProc:
    Exit Property
ErrHandler:
    Debug.Print Err.Description
    Resume ExitProc
End Property

'*******************************************
'KeyDescription - sets/returns the descriptive comment for a key/value pair
'NOTE:  this is mainly used for documentation purposes
Public Property Get KeyDescription() As String
    'Returns:       if the key does not exist will return ""
    '               Otherwise will return the contents of the attribute
    On Error GoTo ErrHandler
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)

    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            Set oNodes = oBlock.childNodes

            If oNodes.length <> 0 Then
                For Each oNode In oNodes
                    If oNode.Attributes(mAttrNameID).Text = msKey Then
                        KeyDescription = oNode.Attributes(mAttrDescID).Text
                        Exit Property
                    End If
                Next oNode
            End If
            Exit For
        End If
    Next oBlock

ExitProc:
    Exit Property
ErrHandler:
    Resume ExitProc
End Property

Public Property Let KeyDescription(sValue As String)
    'Note:   if the section and/or key do not exist this routine will call routines
    '        to create them
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNodeChild As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMNode
    
    'setup the document for access
    Set oBlocks = moDoc.selectNodes(mCONFIG)

    If oBlocks.length = 0 Then
        CreateConfigure
    End If
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)
    
    If oBlocks.length = 0 Then
        CreateSection
        KeyDescription = sValue
    End If
    
    'iterate through the blocks until we find the correct section and key
    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            Set oNodes = oBlock.childNodes

            If oNodes.length <> 0 Then
                For Each oNode In oNodes
                    If oNode.Attributes(mAttrNameID).Text = msKey Then
                        'set the description value and save the file
                        oNode.Attributes(mAttrDescID).Text = sValue
                        moDoc.save msFilename
                        Exit Property
                    End If
                Next oNode
            Else
                'if we reach this point then the section exists but the key does not
                'so we create the new key and value pair
                CreateKeyValue oBlock, , sValue
                moDoc.save msFilename
                Exit Property
            End If
            Exit For
        End If
    Next oBlock
    
ExitProc:
    Exit Property
ErrHandler:
    Debug.Print Err.Description
    Resume ExitProc
End Property

'*******************************************
'SectionDescription - sets/returns the descriptive comment for a section
'NOTE:  this is mainly used for documentation purposes
Public Property Get SectionDescription() As String
    'Returns:       if the section does not exist will return ""
    '               Otherwise will return the contents of the attribute
    On Error GoTo ErrHandler
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)

    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            SectionDescription = oBlock.Attributes(mAttrValueID).Text
            Exit Property
        End If
    Next oBlock

ExitProc:
    Exit Property
ErrHandler:
    Resume ExitProc
End Property

Public Property Let SectionDescription(sValue As String)
    'Note:   if the section does not exist this routine will call the routine to create it
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNodeChild As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMNode
    
    'setup the document for access
    Set oBlocks = moDoc.selectNodes(mCONFIG)

    If oBlocks.length = 0 Then
        CreateConfigure
    End If
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)
    
    If oBlocks.length = 0 Then
        CreateSection
        SectionDescription = sValue
    End If
    
    'iterate through the blocks until we find the correct section and key
    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            oBlock.Attributes(mAttrValueID).Text = sValue
            moDoc.save msFilename
            Exit Property
        End If
    Next oBlock
    
ExitProc:
    Exit Property
ErrHandler:
    Debug.Print Err.Description
    Resume ExitProc
End Property

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Sub DeleteKey()
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList

    'setup the document for access
    Set oBlocks = moDoc.selectNodes(mBLOCK)

    'iterate through the blocks until we find the correct section
    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            Set oNodes = oBlock.childNodes

            If oNodes.length <> 0 Then
                For Each oNode In oNodes
                    If oNode.Attributes(mAttrNameID).Text = msKey Then
                        'delete the key and save the file
                        oNode.parentNode.removeChild oNode
                        moDoc.save msFilename
                        Exit For
                    End If
                Next oNode
            End If
            Exit For
        End If
    Next oBlock

ExitProc:
    Exit Sub
ErrHandler:
    Debug.Print Err.Description
    Resume ExitProc
End Sub

Public Sub DeleteSection()
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)

    'iterate through the blocks until we find the correct section
    For Each oBlock In oBlocks
        If oBlock.Attributes(mAttrNameID).Text = msSection Then
            'delete the section and save the file
            oBlock.parentNode.removeChild oBlock
            moDoc.save msFilename
        End If
    Next oBlock

ExitProc:
    Exit Sub
ErrHandler:
    Debug.Print Err.Description
    Resume ExitProc
End Sub

Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef lCount As Long)
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oNodes As MSXML2.IXMLDOMNodeList
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)
    
    'iterate through the blocks until we find the correct section
    For Each oBlock In oBlocks
        If oBlock.baseName = mSECTION Then
            If oBlock.Attributes(mAttrNameID).Text = msSection Then
                Erase sKey
                lCount = 0
                Set oNodes = oBlock.childNodes

                If oNodes.length <> 0 Then
                    For Each oNode In oNodes
                        'resize the array and set the key name
                        If oNode.baseName = mKEY Then
                            lCount = lCount + 1
                            ReDim Preserve sKey(1 To lCount) As String
                            sKey(UBound(sKey)) = oNode.Attributes(mAttrNameID).Text
                        End If
                    Next oNode
                End If
                Exit For
            End If
        End If
    Next oBlock

ExitProc:
    On Error Resume Next
    Exit Sub
ErrHandler:
    Debug.Print Err.Number & ": " & Err.Description
    Resume ExitProc
End Sub

Public Sub EnumerateAllSections(ByRef sSection() As String, ByRef lCount As Long)
    On Error GoTo ErrHandler
    
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oBlocks As MSXML2.IXMLDOMNodeList
    Dim lSection As Long
    
    Set oBlocks = moDoc.selectNodes(mBLOCK)

    If oBlocks.length = 0 Then
        CreateConfigure
    End If
    
    Erase sSection
    lCount = 0
    
    'iterate through the blocks and set the section name
    For Each oBlock In oBlocks
        If oBlock.baseName = mSECTION Then
            lCount = lCount + 1
            ReDim Preserve sSection(1 To lCount) As String
            sSection(UBound(sSection)) = oBlock.Attributes(mAttrNameID).Text
        End If
    Next oBlock

ExitProc:
    On Error Resume Next
    Exit Sub
ErrHandler:
    Debug.Print Err.Number & ": " & Err.Description
    Resume ExitProc
End Sub

Public Function IniToXml(ByRef INIFile As String, Optional ByRef XMLFile As String = vbNullString) As Long
    'Code in this routine is based on code from Pamela RAI and the resulting XML
    'file is based on Pamelas format with the addition of the Description attribute
    'and the code for the processing instructions
    'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=48510&lngWId=1
    'The exception is that I am using the DOMDocument object to create the XML file
    'instead of opening a file and writing strings to it, but the methodology is pretty
    'much the same
    Dim oKey As MSXML2.IXMLDOMNode
    
    Dim lpSections As String
    Dim nSize As Long
    Dim nMaxSize As Long
    Dim sSections() As String
    Dim lSection As Long
    Dim lpKeys As String
    Dim sKeys() As String
    Dim lKey As Long
    Dim sValues() As String
    
    If XMLFile = vbNullString Then
        If msFilename = vbNullString Or IsMissing(XMLFile) Then
            ' set the XML's file name based on the INI file name
            'in this case we simply add .xml to the end of the INI filename
            'ie. Sample.ini will become sample.ini.xml
            XMLFile = INIFile & ".xml"
        Else
            'if the person has already set a path for the file then use that
            XMLFile = msFilename
        End If
    End If

    ' Get all sections names
    ' Making sure allocate enough space for data returned
    nMaxSize = INITIAL_BUFFER_SIZE / 2
    Do
        nMaxSize = nMaxSize * 2
        lpSections = Space$(nMaxSize)
        nSize = GetPrivateProfileSectionNames(lpSections, nMaxSize, INIFile)
    Loop Until nSize = 0 Or nSize < nMaxSize - 2
    
    sSections() = Split(lpSections, Chr(0))
    
    'create the main "Configuration" element
    CreateConfigure
    
    'get key names within the sections
    For lSection = LBound(sSections) To UBound(sSections)
        If (TrimNull(sSections(lSection)) <> vbNullString) Then
            'create the "Section" node and its attributes
            msSection = sSections(lSection)
            Set oKey = CreateSection
            
            'now get the key names and values in the section
            nMaxSize = INITIAL_BUFFER_SIZE / 2
            Do
                nMaxSize = nMaxSize * 2
                lpKeys = Space$(nMaxSize)
                nSize = GetPrivateProfileSection(sSections(lSection), lpKeys, nMaxSize, INIFile)
            Loop Until nSize = 0 Or nSize < nMaxSize - 2
            
            sKeys() = Split(lpKeys, Chr(0))
            For lKey = LBound(sKeys) To UBound(sKeys)
                If (TrimNull(sKeys(lKey)) <> vbNullString) Then
                    sValues() = Split(sKeys(lKey), "=")
                    'now set the attributes
                    msKey = sValues(mAttrNameID)
                    CreateKeyValue oKey, sValues(mAttrValueID)
                End If
            Next lKey
        End If
    Next lSection
    moDoc.save XMLFile
    
    IniToXml = True

End Function

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Sub CreateKeyValue(ByRef oParent As IXMLDOMNode, Optional ByRef sValue As String, Optional ByRef sDesc As String)
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMNode
        
    Set oNode = moDoc.createNode(NODE_ELEMENT, mKEY, vbNullString)
    oParent.appendChild oNode
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, mAttrName, vbNullString)
    oNode.Attributes.setNamedItem oAttr
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, mAttrValue, vbNullString)
    oNode.Attributes.setNamedItem oAttr
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, mAttrDesc, vbNullString)
    oNode.Attributes.setNamedItem oAttr
    oNode.Attributes(mAttrNameID).Text = msKey
    If sValue <> vbNullString Then
        oNode.Attributes(mAttrValueID).Text = sValue
    End If
    If sDesc <> vbNullString Then
        oNode.Attributes(mAttrDescID).Text = sDesc
    End If
End Sub

Private Function CreateSection() As IXMLDOMNode
    Dim oBlock As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMNode

    Set oBlock = moDoc.selectSingleNode(mCONFIG)
    Set oNode = moDoc.createNode(NODE_ELEMENT, mSECTION, vbNullString)
    oBlock.appendChild oNode
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, mAttrName, vbNullString)
    oNode.Attributes.setNamedItem oAttr
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, mAttrDesc, vbNullString)
    oNode.Attributes.setNamedItem oAttr
    oNode.Attributes(mAttrNameID).Text = msSection
    Set CreateSection = oNode
End Function

Private Sub CreateConfigure()
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMAttribute
    
    On Error Resume Next
    'add processing information - this is in case we decide later to go with
    'different versions of the xml format and need to support the different
    'formats
    Set oNode = moDoc.createNode(NODE_PROCESSING_INSTRUCTION, "xml", vbNullString)
    moDoc.appendChild oNode
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, "version", vbNullString)
    oNode.Attributes.setNamedItem oAttr
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, "encoding", vbNullString)
    oNode.Attributes.setNamedItem oAttr
    Set oAttr = moDoc.createNode(NODE_ATTRIBUTE, "standalone", vbNullString)
    oNode.Attributes.setNamedItem oAttr
    oNode.Attributes(0).Text = "1.0"
    oNode.Attributes(1).Text = "UTF-8"
    oNode.Attributes(2).Text = "yes"
    'finally add the configuration node
    Set oNode = moDoc.createNode(NODE_ELEMENT, mCONFIG, vbNullString)
    moDoc.appendChild oNode
    moDoc.save msFilename
    
End Sub

Private Function TrimNull(ByVal StrIn As String) As String
    'this routine was obtained from Karl Peterson's web site - it is used in several
    'of his samples
    'http://vb.mvps.org/
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left$(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim$(StrIn)
   End Select
End Function


