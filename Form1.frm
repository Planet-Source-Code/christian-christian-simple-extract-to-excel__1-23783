VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple Extract"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract to E&xcel"
      Height          =   375
      Left            =   4725
      TabIndex        =   1
      Top             =   6030
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   5010
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   8837
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Columns  "
      Height          =   645
      Left            =   225
      TabIndex        =   2
      Top             =   180
      Width           =   2625
      Begin VB.OptionButton Option1 
         Caption         =   "All"
         Height          =   285
         Index           =   2
         Left            =   1845
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Even"
         Height          =   285
         Index           =   1
         Left            =   1035
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Odd"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   270
         Width           =   690
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExtract_Click()
    Screen.MousePointer = vbHourglass
    Call Populate_to_Excel(Form1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Initialize()
    With fg
        ' structure
        .Cols = 20
        .Rows = 20
        .FixedCols = 0
        .FixedRows = 1
        
        ' appearance
        .BorderStyle = flexBorderSingle
        .GridLines = flexGridFlat
        .BackColorBkg = .BackColor
        .FocusRect = flexFocusNone
        .AllowUserResizing = flexResizeColumns
        
        ' behavior
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        
        ' content
        Dim i As Integer
        Dim h As Integer
        
        For i = 0 To 19
            For h = 0 To 19
                .TextMatrix(i, h) = CStr(i) & CStr(h)
            Next h
        Next i
    End With
End Sub



Private Sub Option1_Click(Index As Integer)
Dim i As Integer
Dim h As Integer
    If Index = 0 Then
        h = 0
    ElseIf Index = 1 Then
        h = 1
    Else
        h = 2
    End If
        
    For i = 0 To fg.Cols - 1
        If i Mod 2 = h Then
            fg.ColWidth(i) = 0
        Else
            fg.ColWidth(i) = 1000
        End If
    Next i
End Sub
