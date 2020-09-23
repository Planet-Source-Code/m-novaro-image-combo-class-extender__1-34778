VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Image Combo EX !"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Drop with down key"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Auto size dropdown"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide dropdown"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Autoselect on entry"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show dropdown"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Allow only present items"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "ImageCombo1"
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Found:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Item:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myComboEx As cImgComboEx
Attribute myComboEx.VB_VarHelpID = -1

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        myComboEx.AllowOnlyPresentItems = True
    Else
        myComboEx.AllowOnlyPresentItems = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        myComboEx.SelectOnEntry = True
    Else
        myComboEx.SelectOnEntry = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        myComboEx.DropWithDownKey = True
    Else
        myComboEx.DropWithDownKey = False
    End If
End Sub

Private Sub Command1_Click()
    myComboEx.ShowDropDown
End Sub

Private Sub Command2_Click()
    myComboEx.HideDropDown
End Sub

Private Sub Command3_Click()
    myComboEx.AutosizeDropDown
End Sub

Private Sub Form_Load()
    Set myComboEx = New cImgComboEx
    With myComboEx
        .SetRefToCombo ImageCombo1
        .AllowOnlyPresentItems = False
    End With
    
    With ImageCombo1
        .Text = ""
        .ComboItems.Add , , "Test"
        .ComboItems.Add , , "AutoComplete"
        .ComboItems.Add , , "ImageCombo"
        .ComboItems.Add , , "This is the longest text in the combo I have entered"
    End With

    Label1.Caption = ""
    Label4.Caption = ""
End Sub

Private Sub myComboEx_ItemAccepted(theItem As String, bItemFound As Boolean)
    Label1.Caption = theItem
    Label4.Caption = bItemFound
End Sub
