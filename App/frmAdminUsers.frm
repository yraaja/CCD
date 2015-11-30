VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmAdminUsers 
   Caption         =   "User Administration"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   Icon            =   "frmAdminUsers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   11340
   Begin VB.Frame Frame1 
      Caption         =   "User"
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   210
         Width           =   1815
      End
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   260
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   260
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   8880
      TabIndex        =   5
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10080
      TabIndex        =   4
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   240
      Width           =   1035
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   4260
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   7514
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(23)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Named:id=29:Normal"
      _StyleDefs(45)  =   ":id=29,.parent=0"
      _StyleDefs(46)  =   "Named:id=30:Heading"
      _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=30,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=31:Footing"
      _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=32:Selected"
      _StyleDefs(52)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=33:Caption"
      _StyleDefs(54)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(55)  =   "Named:id=34:HighlightRow"
      _StyleDefs(56)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(57)  =   "Named:id=35:EvenRow"
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&,.ellipsis=0"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29,.ellipsis=0"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11100
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   2
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmAdminUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' SYSTEM ADMINISTRATION FORM FOR CCD USERS
' 9/12/2005 RTD


'
'   Class to handle grid
Dim m_objGridMap As New CAdminUsers
'
Dim m_blnFirstSearch As Boolean     ' Is this the first search we have made on this screen.
Dim m_blnDoubleClick As Boolean     ' Did a double click just occurr
Dim m_sngYCoord As Single
'
'   Recordset to hold query results
Dim m_rec As New ADODB.RecordSet
'
'   True if the Update had errors, used in QueryUnload
Dim m_blnWereErrors As Boolean
'
'   Keeps up with the field that last had focus when form
'   is deactivate, so when activated can set focus.
Dim m_strCurrentFormControl As String
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID As String
'

'
'   Called from frmMain when the user clicks on the
'   toolbar buttons for sorting.
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub EventNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
    Dim varBookmark

    varBookmark = TDBGrid.Bookmark
    If eNotifyType = esnUserRecordupdated Then
        cmdSearch_Click
        TDBGrid.MoveFirst
        Do While Not TDBGrid.EOF
            If TDBGrid.Columns("User ID").Text = sAffectedRecordIdentifier Then
                varBookmark = TDBGrid.Bookmark
                Exit Do
            End If
            TDBGrid.MoveNext
        Loop
        TDBGrid.Bookmark = varBookmark
    End If

End Sub

Private Sub cmdDelete_Click()
    
    m_objGridMap.Delete
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cmdNew_Click()
    Dim fInfo As New dlgUserInfo
    
    fInfo.ShowExtendedInfo = False
    fInfo.NewUser
    fInfo.Show vbModal
    Set fInfo = Nothing
    
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim strSELECT As String
    Dim dtmStart As Date
    Dim sUserId As String
    Dim sUserName As String
    
    TDBGrid.Update

    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        Else
            If Button = vbNo Then
            TDBGrid.DataChanged = False
        ElseIf Button = vbCancel Then
            ' Cancel the search
            Exit Sub
        End If
    End If
    End If
    Screen.MousePointer = vbHourglass
  
    sUserId = SQLFixString(SQLChangeWildcard(txtUserID.Text))
    sUserName = SQLFixString(SQLChangeWildcard(txtUserName.Text))
    
    strSELECT = "EXEC usp_select_user_names"
    strSELECT = strSELECT & " @user_id = '" & sUserId & "'"
    strSELECT = strSELECT & ",@user_name = '" & sUserName & "'"

    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    ' Use g_objDAL to perform select
    blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
    If blnRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription
        lblRowCount.Caption = "0 rows returned."
        Exit Sub
    End If
    
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
        
    If m_rec.RecordCount > 0 Then
        lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned in " + str(DateDiff("s", dtmStart, Now)) + " seconds"
    Else
        lblRowCount.Caption = "0 rows returned."
    End If
    
    ' If the upper bound was hit, inform user
    If m_rec.RecordCount = MAX_RECORDS And m_rec.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If
    
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdUpdate_Click()
' UPDATE RECORDS
    Dim blnRet As Boolean
    Dim vntBookmark As Variant
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    m_blnWereErrors = False
    vntBookmark = TDBGrid.Bookmark
    
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True
        Screen.MousePointer = vbNormal
    End If
    TDBGrid.Bookmark = vntBookmark
    Screen.MousePointer = vbNormal
    Exit Sub
    
Err_Handler:
    MsgBox Err.Description
    
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    
    If Me.WindowState <> vbMinimized Then
        If Len(m_strCurrentFormControl) > 0 Then
            For Each ctl In Me.Controls
                If ctl.Name = m_strCurrentFormControl Then
                    ctl.SetFocus
                    Exit For
                End If
            Next ctl
        End If
        OutputView True
        ShowGridSort
        m_objGridMap.SetMenuBar
        ShowToolbarIcons True
    End If
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()

    Status ("Loading User Administration...")
    Screen.MousePointer = vbHourglass
    sEventSubscriberID = EventSubscriberAdd(Me)
    m_blnFirstSearch = True
    ' Initialize grid only once
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False

End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    cmdSearch_Click
    Status ("")
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Dim Button As VbMsgBoxResult
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbQuestion)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the close
            If m_blnWereErrors Then
                Cancel = True
            End If
        ElseIf Button = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 10500 Then
            TDBGrid.Width = Me.Width - (TDBGrid.Left * 3)
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 10500
        End If
        
        If Me.Height >= 6135 Then
            cmdUpdate.Top = Me.Height - 1035
            cmdNew.Top = cmdUpdate.Top
            cmdDelete.Top = cmdUpdate.Top
            cmdDelete.Left = Me.Width - cmdDelete.Width - 240
            cmdNew.Left = cmdDelete.Left - cmdNew.Width - 240
            cmdUpdate.Left = cmdNew.Left - cmdUpdate.Width - 240
            TDBGrid.Height = cmdUpdate.Top - TDBGrid.Top - 240
        Else
            Me.Height = 6135
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    '   Disables & hides the sort buttons on the main form.
    If m_rec.State <> adStateClosed Then m_rec.Close
    Set m_rec = Nothing
    HideGridSort
    ShowToolbarIcons False
    EventSubscriberRemove sEventSubscriberID
    
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    cmdUpdate.Enabled = True
End Sub

Private Sub TDBGrid_ButtonClick(ByVal ColIndex As Integer)
    
    If ColIndex = TDBGrid.Columns("Role").ColIndex Then
        ' Get new User Role value
        Dim frm As New dlgUserRole
        If TDBGrid.Text = "" Then
            frm.UserRole = 0
        Else
            frm.UserRole = TDBGrid.Text
        End If
        frm.Show vbModal
        If frm.ReturnResult = vbOK Then
            TDBGrid.Text = frm.UserRole
        End If
        Set frm = Nothing
    End If
    
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Unit Cost button, open single record view
            'cmdUnitCost_Click
        End If
    Else
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
                MsgBox strErrorMsg
            End If
        Else
            If TDBGrid.RowContaining(Y) <> TDBGrid.Row Then
                m_sngYCoord = Y
            End If
        End If
    End If
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub

Public Function SetupGridPrint()
' SETUP PROPERTIES FOR TRUE DBGRID PRINTING

    TDBGrid.PrintInfo.PreviewCaption = Me.Caption & " Preview"
    TDBGrid.PrintInfo.PageHeader = "\t" & Me.Caption
    
    TDBGrid.PrintInfo.PreviewInitHeight = START_HEIGHT / Screen.TwipsPerPixelX
    TDBGrid.PrintInfo.PreviewInitWidth = START_WIDTH / Screen.TwipsPerPixelY
    TDBGrid.PrintInfo.PreviewInitPosX = 5 + (fMainForm.Left / Screen.TwipsPerPixelX)
    TDBGrid.PrintInfo.PreviewInitPosY = 4 + ((fMainForm.Top + fMainForm.sbStatusBar.Height + fMainForm.tbToolBar.Height * 2) / Screen.TwipsPerPixelY)
    TDBGrid.PrintInfo.PageHeaderFont.Bold = True
    TDBGrid.PrintInfo.PageHeaderFont.Size = 12
    TDBGrid.PrintInfo.PageFooter = CStr(Now) & "\t\tPage \p"
    ' ORIENTATION 1=PORTRAIT | 2=LANDSCAPE
    TDBGrid.PrintInfo.SettingsOrientation = 2
    TDBGrid.PrintInfo.SettingsMarginBottom = 720
    TDBGrid.PrintInfo.SettingsMarginTop = 720
    TDBGrid.PrintInfo.SettingsMarginLeft = 720
    TDBGrid.PrintInfo.SettingsMarginRight = 720
    
End Function

Public Function PreviewReport()
'PREVIEW THE GRID TO THE SCREEN

    SetupGridPrint
    TDBGrid.PrintInfo.PrintPreview

End Function

Public Function PrintReport()
'SEND THE GRID TO THE PRINTER

    SetupGridPrint
    TDBGrid.PrintInfo.PrintData

End Function

