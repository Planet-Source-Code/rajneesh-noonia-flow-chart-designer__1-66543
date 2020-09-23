VERSION 5.00
Object = "{5BFE14E7-377F-412C-B492-4A7816FBFF97}#1.0#0"; "Graph.ocx"
Begin VB.Form frmTest 
   BackColor       =   &H8000000A&
   Caption         =   "Flow Chart Designer : By Rajneesh"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   750
      Max             =   40
      Min             =   12
      TabIndex        =   5
      Top             =   5850
      Value           =   20
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "100 %"
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   4
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Index           =   4
      Left            =   7290
      TabIndex        =   3
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Index           =   3
      Left            =   6300
      TabIndex        =   2
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Index           =   2
      Left            =   5310
      TabIndex        =   1
      Top             =   5820
      Width           =   945
   End
   Begin oGraph.oConvas Convas1 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9551
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "frmTest.frx":0000
      Top             =   1050
      Width           =   480
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Sub HScroll1_Change()
    Convas1.Zoom = HScroll1.Value * 5
    Convas1.Paint
End Sub


Private Sub Command1_Click(Index As Integer)
    Dim Dummy As Variant, iFile As Integer, data() As Byte
    Dim UNCFile As String

    Select Case Index
        
        Case 2: 'Save
            Dummy = Convas1.BinaryData
            UNCFile = App.Path & "\Convas.blg"
            iFile = FreeFile
            Open UNCFile For Binary As iFile
            Put iFile, , Dummy
            Close iFile '
        Case 3: 'Clear
            Convas1.ClearWorkSheet
        Case 4: 'Load
            iFile = FreeFile
            UNCFile = App.Path & "\Convas.blg"
            
            Open UNCFile For Binary As iFile
            Get iFile, , Dummy
            Close iFile
            'Assign the Variant to a bytearray to the bag.contents
            
            If Len(Dummy) > 0 Then
                'Convas1.BinaryData = ""
                Convas1.ClearWorkSheet
                Convas1.BinaryData = Dummy
            End If
            'Convas1.Paint
        Case 5: '100 %
            Convas1.Zoom = 100
            Convas1.Paint
        Case 6:
        Case 7:
    End Select
End Sub

    'Dim p1 As oGraph.oPicture, p2 As oGraph.oPicture, p3 As oGraph.oPicture, p4 As oGraph.oPicture
    'Dim t1 As oGraph.oText
    'Dim s1 As oGraph.oLine, s2 As oGraph.oLine, s3 As oGraph.oLine
    




Private Sub Form_Load()
  
    Me.Show
    
    With Convas1
        Set p1 = .AddNode("p1")
        Set p2 = .AddNode("p2")
        Set p3 = .AddNode("p3")
        Set p4 = .AddNode("p4")
        Set t1 = .AddText("t1")
        Set s1 = .AddStep(OnTCompletion, "p1", "p2", "s1")
        Set s2 = .AddStep(OnTFail, "p2", "p3", "s2")
        Set s3 = .AddStep(OnTSuccess, "p2", "p4", "s3")
    End With
    
    With p1
        .Caption = "Initilize"
        .CentreX = 1200
        .CentreY = 1200
        .ToolTipText = "I am " & .Caption
        Set .Image = Image1.Picture
    End With
    
    With p2
        .Caption = "Process Data"
        .CentreX = 4200
        .CentreY = 1200
        .ToolTipText = "I am " & .Caption
        Set .Image = Image1.Picture
    End With
    
    With p3
        .Caption = "Do Some Alert"
        .CentreX = 1200
        .CentreY = 4200
        .ToolTipText = "I am " & .Caption
        Set .Image = Image1.Picture
    End With
    
    With p4
        .Caption = "All Gone Well - Winner"
        .CentreX = 6200
        .CentreY = 4200
        .ToolTipText = "I am " & .Caption
        Set .Image = Image1.Picture
    End With
    
    With t1
        .Caption = "Sample Flow Chart (Double Click Me)"
        .CentreX = 4500
        .CentreY = 300
        .Font.Size = 20
        .ToolTipText = "I am " & .Caption
        .Font.Bold = True
    End With
    
     s1.LayereLineType = OnTCompletion
     s2.LayereLineType = OnTFail
     s3.LayereLineType = OnTSuccess
     
     s1.ToolTipText = "I am line s1"
     s2.ToolTipText = "I am line s2"
     s3.ToolTipText = "I am line s3"
        
     p1.Visible = True
     p2.Visible = True
     p3.Visible = True
     p4.Visible = True
     t1.Visible = True
     s1.Visible = True
     s2.Visible = True
     s3.Visible = True
     p1.Activate
     p2.Activate
     p3.Activate
     p4.Activate
     t1.Activate
End Sub



Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Convas1.Width = Me.Width - 220
        Convas1.Height = Me.Height - 1000
        HScroll1.Top = Convas1.Height + Convas1.Top + 50
        Command1(2).Top = HScroll1.Top
        Command1(3).Top = HScroll1.Top
        Command1(4).Top = HScroll1.Top
        Command1(5).Top = HScroll1.Top
    End If
End Sub



