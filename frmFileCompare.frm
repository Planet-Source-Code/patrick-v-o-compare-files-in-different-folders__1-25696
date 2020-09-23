VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileCompare 
   Caption         =   "Compare Files In Two Different Folders"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   Icon            =   "frmFileCompare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRight 
      Height          =   195
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "Show all found files in Right listview - not dynamic"
      Top             =   440
      Width           =   255
   End
   Begin VB.CheckBox chkLeft 
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Show all found files in Left listview - not dynamic"
      Top             =   440
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   285
      Left            =   9120
      TabIndex        =   7
      Top             =   380
      Width           =   975
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "&Compare"
      Height          =   285
      Left            =   9120
      TabIndex        =   6
      Top             =   40
      Width           =   975
   End
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Enter complete path / folder for Left listview"
      Top             =   120
      Width           =   3495
   End
   Begin MSComctlLib.ListView lvcLeft 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   3157
      EndProperty
   End
   Begin VB.TextBox txtRight 
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Enter complete path / folder for Right listview"
      Top             =   120
      Width           =   3495
   End
   Begin MSComctlLib.ListView lvcRight 
      Height          =   3135
      Left            =   5160
      TabIndex        =   9
      Top             =   720
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   3157
      EndProperty
   End
   Begin VB.TextBox txtLeftExt 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Enter file extensions to apply to Left folder"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtRightExt 
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      ToolTipText     =   "Enter file extensions to apply to Left folder"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblRightCount 
      Caption         =   "File Count"
      Height          =   255
      Left            =   6435
      TabIndex        =   11
      Top             =   435
      Width           =   2415
   End
   Begin VB.Label lblLeftCount 
      Caption         =   "File Count"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   435
      Width           =   2295
   End
End
Attribute VB_Name = "frmFileCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ever needed to compare files in two different folders ?
' How did you do it, I bet you used paper and pencil.
' Now you can use the FileCompare tool to indicate difference
'  in two folders, all in seconds flat.
'
' Special thanks goes to all the developers who submit code samples
'  to the Plant Source Code website.
'
' Feel free to use this code, and if you truly like the sample,
'  then don't forget to vote for me.
' Thanks.
'
' Patrick van Oppen
' pvanoppen@mindspring.com

Option Explicit

Private Sub cmdCompare_Click()
    Dim itmRight As ListItem                ' Value returned from the find in the Right listview
    Dim intLeft As Integer                  ' Counter to loop through the Left listview
    Dim intRight As Integer                 ' Counter to loop through the Right listview
    Dim intMatch As Integer                 ' Counter to indicate how many files were matching
    
    Me.MousePointer = vbHourglass
        
    lvcLeft.ListItems.Clear                 ' Clear the Left listview
    lvcRight.ListItems.Clear                ' Clear the Right listview
    lblLeftCount.Caption = ""               ' Clear the Left files count label
    lblRightCount.Caption = ""              ' Clear the Right files count label
    Me.Refresh                              ' Refresh the screen to show empty controls
    Call GetFilesLeft                       ' Collect files and their info, then load in the Left listview
    Call GetFilesRight                      ' Collect files and their info, then load in the Left listview
    Me.Refresh                              ' Refresh the screen to show pre-compare results

    intMatch = 0                            ' Set starting point for matching file count

    ' Loop through the Left listview and find match in Right listview
    For intLeft = 1 To lvcLeft.ListItems.Count
        With lvcLeft.ListItems(intLeft)
            Set itmRight = lvcRight.FindItem(.Text, lvwText)
            If itmRight Is Nothing Then
                ' No action, we start will all files flagged
            Else
                If .SubItems(1) = itmRight.SubItems(1) And _
                    .SubItems(2) = itmRight.SubItems(2) Then
                    .Checked = False
                    itmRight.Checked = False
                    intMatch = intMatch + 1
                Else
                    .Checked = True
                    itmRight.Checked = True
                End If
            End If
        End With
    Next intLeft
    ' Show result in the Left file count label
    With lblLeftCount
        .Caption = .Caption & " - " & lvcLeft.ListItems.Count - intMatch & " different"
    End With
    ' Show result in the Right file count label
    With lblRightCount
        .Caption = .Caption & " - " & lvcRight.ListItems.Count - intMatch & " different"
    End With
    ' If the user choose to, remove all equal files form the Left listview
    If chkLeft.Value = vbUnchecked Then
        intLeft = lvcLeft.ListItems.Count
        For intLeft = lvcLeft.ListItems.Count To 1 Step -1
            With lvcLeft
                If .ListItems(intLeft).Checked = False Then
                    .ListItems.Remove (intLeft)
                End If
            End With
        Next intLeft
    End If
    ' If the user choose to, remove all equal files form the Right listview
    If chkRight.Value = vbUnchecked Then
        intRight = lvcRight.ListItems.Count
        For intRight = lvcRight.ListItems.Count To 1 Step -1
            With lvcRight
                If .ListItems(intRight).Checked = False Then
                    .ListItems.Remove (intRight)
                End If
            End With
        Next intRight
    End If
    
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdExit_Click()
    Unload Me                               ' Remove graphics from screen
    Set frmFileCompare = Nothing            ' Remove code and variables from PC memory
End Sub

Private Sub Form_Load()
    txtLeft.Text = "C:\TEMP"                ' Set the initial search folder for the Left listview
    txtLeftExt.Text = "*.*"                 ' Set the initial search file extension for the Left listview
    lblLeftCount.Caption = ""               ' Clear the Left listview file count label
    chkLeft.Value = vbChecked               ' Set flag to indicate we want to see all found files in Left search folder
    txtRight.Text = "C:\TEMP"               ' Set the initial search folder for the Right listview
    txtRightExt.Text = "*.*"                ' Set the initial search file extension for the Right listview
    lblRightCount.Caption = ""              ' Clear the Right listview file count label
    chkRight.Value = vbChecked              ' Set flag to indicate we want to see all found files in Right search folder
    lvcLeft.ListItems.Clear                 ' Clear Left listview
    lvcRight.ListItems.Clear                ' Clear Right listview
End Sub

Private Sub lvcLeft_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim strColName As String                ' Column name to sort by
    Dim strSort As String                   ' Holds column name and sort direction
    Static booSortAscL As Boolean           ' Flag to define the Left listview sort status
    Static strPrevColL As String            ' Holds the last Left listview sorted column name
    
    strColName = ColumnHeader
    ' Check if clicked column is same column as last sorted one
    If strColName = strPrevColL Then
        ' If same column clicked, reverse sort order
        If booSortAscL Then
            With lvcLeft
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwDescending
                .Sorted = True
            End With
            booSortAscL = False
        Else
            With lvcLeft
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwAscending
                .Sorted = True
            End With
            booSortAscL = True
        End If
    Else
        ' On first time clicking column, sort Ascending and set flag
        With lvcLeft
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End With
        booSortAscL = True
    End If
    strPrevColL = ColumnHeader               ' Set variable to remember which column was last sorted
End Sub

Private Sub lvcRight_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim strColName As String                ' Column name to sort by
    Dim strSort As String                   ' Holds column name and sort direction
    Static booSortAscR As Boolean           ' Flag to define Right listview sort status
    Static strPrevColR As String            ' Holds the last Right listview sorted column name
    
    strColName = ColumnHeader
    ' Check if clicked column is same column as last sorted one
    If strColName = strPrevColR Then
        ' If same column clicked, reverse sort order
        If booSortAscR Then
            With lvcRight
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwDescending
                .Sorted = True
            End With
            booSortAscR = False
        Else
            With lvcRight
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwAscending
                .Sorted = True
            End With
            booSortAscR = True
        End If
    Else
        ' On first time clicking column, sort Ascending and set flag
        With lvcRight
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        End With
        booSortAscR = True
    End If
    strPrevColR = ColumnHeader               ' Set variable to remember which column was last sorted
End Sub

Private Sub txtLeft_GotFocus()
' When field gets focus, highlight all existing text
    With txtLeft
        .SelStart = 0
        .SelLength = Len(txtLeft.Text)
    End With
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    ' Convert every entered alpha character to upper case
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtLeftExt_GotFocus()
' When field gets focus, highlight all existing text
    With txtLeftExt
        .SelStart = 0
        .SelLength = Len(txtLeftExt.Text)
    End With
End Sub

Private Sub txtLeftExt_KeyPress(KeyAscii As Integer)
    ' Convert every entered alpha character to upper case
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtRight_GotFocus()
' When field gets focus, highlight all existing text
    With txtRight
        .SelStart = 0
        .SelLength = Len(txtRight.Text)
    End With
End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
    ' Convert every entered alpha character to upper case
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtRightExt_GotFocus()
' When field gets focus, highlight all existing text
    With txtRightExt
        .SelStart = 0
        .SelLength = Len(txtRightExt.Text)
    End With
End Sub

Private Sub txtRightExt_KeyPress(KeyAscii As Integer)
    ' Convert every entered alpha character to upper case
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub GetFilesLeft()
    Dim strFile As String

    On Error GoTo ErrGetFiles

    ' Take snapshot of files following the Left listview user criteria
    strFile = Dir(txtLeft.Text & "\" & txtLeftExt.Text)
    Do While strFile <> ""
        ' Load all files and their info data into the Left listview
        With lvcLeft.ListItems.Add
            .Text = UCase(strFile)
            .SubItems(1) = GetFileSize(txtLeft.Text & "\" & strFile)
            .SubItems(2) = GetFileDate(txtLeft.Text & "\" & strFile)
            .Checked = True
        End With
        ' Set the next file in the DIR collection
        strFile = Dir
    Loop
    ' Sort the Left listview by the File Name column
    With lvcLeft
        .SortKey = 0
        .SortOrder = lvwAscending
        .Sorted = True
    End With
    ' Show the user how many files are in the Left listview
    lblLeftCount.Caption = lvcLeft.ListItems.Count & " Files"
    
    Exit Sub

ErrGetFiles:
    MsgBox "Error ..." & vbCr & vbCr & Err.Number & vbCr & Err.Description, vbOKOnly + vbCritical, "Error Get Left Files"
    Resume Next
End Sub

Private Sub GetFilesRight()
    Dim strFile As String

    On Error GoTo ErrGetFiles

    ' Take snapshot of files following the Right listview user criteria
    strFile = Dir(txtRight.Text & "\" & txtRightExt.Text)
    Do While strFile <> ""
        ' Load all files and their info data into the Right listview
        With lvcRight.ListItems.Add
            .Text = UCase(strFile)
            .SubItems(1) = GetFileSize(txtRight.Text & "\" & strFile)
            .SubItems(2) = GetFileDate(txtRight.Text & "\" & strFile)
            .Checked = True
        End With
        ' Set the next file in the DIR collection
        strFile = Dir
    Loop
    ' Sort the Right listview by the File Name column
    With lvcRight
        .SortKey = 0
        .SortOrder = lvwAscending
        .Sorted = True
    End With
    ' Show the user how many files are in the Right listview
    lblRightCount.Caption = lvcRight.ListItems.Count & " Files"
    
    Exit Sub

ErrGetFiles:
    MsgBox "Error ..." & vbCr & vbCr & Err.Number & vbCr & Err.Description, vbOKOnly + vbCritical, "Error Get Right Files"
    Resume Next
End Sub

Public Function GetFileDate(strFullFileName As String) As String
    On Error Resume Next
    ' Retrieve the file datestamp
    GetFileDate = FileDateTime(strFullFileName)
End Function

Public Function GetFileSize(strFullFileName As String) As String
    Dim lngSize As Long
    
    On Error Resume Next
    
    ' Retrieve the file size in bytes
    lngSize = FileLen(strFullFileName)

    ' Verify if file is smaller the 1KB, else force a format
    If lngSize <= 1024 Then
        GetFileSize = "1KB"
    Else
        GetFileSize = Format(lngSize, "#,###,####KB")
    End If
End Function
