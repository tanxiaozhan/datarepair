VERSION 5.00
Begin VB.Form autoclose 
   Caption         =   "数据修复"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   2880
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "继续"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "autoclose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Unload Me
End Sub
