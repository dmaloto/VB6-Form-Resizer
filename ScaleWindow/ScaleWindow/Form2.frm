VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ScaleWindow.ScalableForm ScalableForm1 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
