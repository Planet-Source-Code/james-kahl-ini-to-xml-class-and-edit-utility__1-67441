VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "INI XML Editor"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picValues 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   3270
      ScaleHeight     =   5175
      ScaleWidth      =   2655
      TabIndex        =   9
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtKey 
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtValue 
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtKeyDesc 
         Height          =   1095
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox txtSectDesc 
         Height          =   975
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtSection 
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Key Value"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Key"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Key Description"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Section Description"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Section"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox picSettings 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   3135
      TabIndex        =   6
      Top             =   0
      Width           =   3135
      Begin MSComctlLib.TreeView tvwSettings 
         Height          =   4695
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   8281
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Sections && Keys"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   6720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   5925
      ScaleHeight     =   5175
      ScaleWidth      =   1455
      TabIndex        =   5
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdView 
         Caption         =   "View XML"
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdWrite 
         Caption         =   "Write Value"
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelSection 
         Caption         =   "Delete Section"
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelKey 
         Caption         =   "Delete Key"
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save As..."
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&View XML"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&Help"
      Index           =   1
      Begin VB.Menu mnuHelp 
         Caption         =   "About..."
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Read Me"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'Module:        frmMain - Form
'Filename:      frmMain.frm
'Author:        Jim Kahl
'Purpose:       a simple utility to allow creation/editing of INI style XML files
'Assumes:       must be used in conjunction with or in some way reference the XMLConfig
'               class object
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API CONSTANTS
'****************************************************************************************
Private Const SW_SHOW As Long = 5

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
                ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) _
                As Long

'****************************************************************************************
'CONSTANTS - PRIVATE
'****************************************************************************************
Private Const cdOpenFilter As String = "INI Files (*.ini)|*.ini|XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
Private Const cdSaveFilter As String = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"

'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private mcXML As New XMLConfig
Private msFilename As String

'****************************************************************************************
'EVENTS - PRIVATE
'****************************************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelKey_Click()
    'provide a way to delete a key
    Dim eRet As VbMsgBoxResult
    eRet = MsgBox("Are you sure you want to delete this key?", vbYesNo)
    If eRet = vbYes Then
        mcXML.Section = txtSection.Text
        mcXML.Key = txtKey.Text
        mcXML.DeleteKey
        RefreshList
    End If
End Sub

Private Sub cmdView_Click()
    mnuFile_Click 3
End Sub

Private Sub cmdWrite_Click()
    
    'there must be a key name
    If txtKey.Text = vbNullString Then
        MsgBox "You must enter a Key name before attempting to write"
        txtKey.SetFocus
        Exit Sub
    End If
    
    'there must be a section name
    If txtSection.Text = vbNullString Then
        MsgBox "You must enter a Section name before attempting to write"
        txtSection.SetFocus
        Exit Sub
    End If
    
    'make sure we have a valid filename before attempting to write
    If msFilename = vbNullString Then
        mnuFile_Click 1
        If msFilename = vbNullString Then
            Exit Sub
        End If
        mcXML.Path = msFilename
    End If
    
    'now set the properties and refresh the list
    With mcXML
        .Section = txtSection.Text
        .SectionDescription = txtSectDesc.Text
        .Key = txtKey.Text
        .KeyDescription = txtKeyDesc.Text
        .Value = txtValue.Text
    End With
    
End Sub

Private Sub cmdDelSection_Click()
    'provide a way to delete a section
    Dim eRet As VbMsgBoxResult
    eRet = MsgBox("Are you sure you want to delete this section?", vbYesNo)
    If eRet = vbYes Then
        mcXML.Section = txtSection.Text
        mcXML.DeleteSection
        RefreshList
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mcXML = Nothing
End Sub

Private Sub mnuFile_Click(Index As Integer)
    On Error GoTo ErrHandler
    Dim sFile As String
    Dim sXMLFile As String
    
    'set properties for common dialog
    With cdlFile
        .CancelError = True
        .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
    End With
    
    Select Case Index
        Case 0  'Open
            With cdlFile
                .Filter = cdOpenFilter
                .FilterIndex = 2
                .ShowOpen
                sFile = .FileName
            End With
            'if the file is an INI file convert to XML
            If LCase$(Right$(sFile, 3)) = "ini" Then
                'this will be same filename with ".xml" extension
                'ie. sample.ini becomes sample.ini.xml
                mcXML.IniToXml sFile, sXMLFile
                msFilename = sXMLFile
            Else
                msFilename = sFile
            End If
            RefreshList
        Case 1  'Save As
            With cdlFile
                .Filter = cdSaveFilter
                .FilterIndex = 1
                .ShowSave
                sFile = .FileName
            End With
            If msFilename <> vbNullString Then
                'have a current filename but save it as a new filename
                FileCopy msFilename, sFile
            Else
                'we do not have a current file so first we need to kill the old
                Kill sFile
            End If
            msFilename = sFile
        Case 3  'View
            ShellExecute 0, "open", msFilename, vbNullString, vbNullString, SW_SHOW
        Case 5  'Exit
            Unload Me
    End Select
    If msFilename <> vbNullString Then
        mnuFile(1).Enabled = True
        mnuFile(3).Enabled = True
        cmdDelSection.Enabled = True
        cmdDelKey.Enabled = True
        cmdView.Enabled = True
    End If
    Exit Sub
ErrHandler:
    If Err.Number <> 32755 Then
        If Err.Number = 53 Then
            Resume Next
        End If
        'user did not cancel
        Debug.Print Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Select Case Index
        Case 0
            MsgBox "INI XML Editor was written by Jim Kahl"
        Case 1
            ShellExecute 0, "open", App.Path & "\readme.doc", vbNullString, vbNullString, SW_SHOW
    End Select
End Sub

Private Sub tvwSettings_Click()
    On Error Resume Next
    With mcXML
        If tvwSettings.Nodes(tvwSettings.SelectedItem.Text).Children = 0 Then
            'this is for when someone clicks the key name
            .Section = tvwSettings.SelectedItem.Parent.Text
            .Key = tvwSettings.SelectedItem.Text
            txtSection.Text = .Section
            txtSectDesc.Text = .SectionDescription
            txtKey.Text = .Key
            txtKeyDesc.Text = .KeyDescription
            txtValue.Text = .Value
        Else
            'this is for when someone clicks just a section name
            .Section = tvwSettings.SelectedItem.Text
            txtSection.Text = .Section
            txtSectDesc.Text = .SectionDescription
            txtKey.Text = vbNullString
            txtKeyDesc.Text = vbNullString
            txtValue.Text = vbNullString
        End If
    End With
End Sub

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Sub RefreshList()
    Dim sSection() As String
    Dim sKey() As String
    Dim lCount As Long
    Dim lSection As Long
    Dim lKey As Long
    
    On Error GoTo ErrHandler
    
    tvwSettings.Nodes.Clear
    
    'fill the tree view with the Section and Key nodes from the XML file
    With mcXML
        .Path = msFilename
        'EnumerateAllSections will error out if this is a new file or if
        'the file does not contain any Sections - that is ok for now, it
        'just means there is nothing at this time to populate the treeview
        .EnumerateAllSections sSection(), lCount
        For lSection = LBound(sSection) To UBound(sSection)
            tvwSettings.Nodes.Add , , sSection(lSection), sSection(lSection)
            .Section = sSection(lSection)
            'EnumerateCurrentSection will error out if the Section does not
            'have any Keys associated with it, but that's ok since we want to
            'be able to add and delete keys at will
            .EnumerateCurrentSection sKey(), lCount
            For lKey = LBound(sKey) To UBound(sKey)
                .Key = sKey(lKey)
                tvwSettings.Nodes.Add sSection(lSection), tvwChild, , sKey(lKey)
            Next lKey
            tvwSettings.Nodes(sSection(lSection)).Expanded = True
        Next lSection
    End With
    tvwSettings.Nodes(1).Selected = True
    tvwSettings_Click
ErrHandler:
'    Debug.Print Err.Number & ": " & Err.Description
End Sub

