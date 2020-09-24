VERSION 5.00
Begin VB.Form frmSerial 
   Caption         =   "Get Serial No....         by Vikrant Thakker (AnaSys Softwares)"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   3120
      TabIndex        =   4
      Top             =   2460
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   435
      Left            =   1980
      TabIndex        =   3
      Top             =   2460
      Width           =   1035
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   2595
   End
   Begin VB.TextBox txtDisk 
      Height          =   435
      Left            =   1980
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetNo 
      Caption         =   "&Get It"
      Height          =   435
      Left            =   3360
      TabIndex        =   1
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Hey Guys !Plz Vote if you liked the Program ! It feels good ;-)"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   3900
      Width           =   6435
   End
   Begin VB.Label Label2 
      Caption         =   "Serial No of this Drive is"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   1800
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Drive whose serial no. is needed e.g. C:\ or D:\ or etc."
      Height          =   675
      Left            =   60
      TabIndex        =   5
      Top             =   660
      Width           =   1875
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Disk As String

Private Sub cmdClear_Click()
txtDisk.Text = ""
txtSerialNo.Text = ""
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGetNo_Click()
' It will give the serial no of the Drive written in textbox
Disk = txtDisk.Text
txtSerialNo.Text = VolumeSerialNumber(Disk)
End Sub
