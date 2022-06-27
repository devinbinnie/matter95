VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matter95 Login"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "LoginForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LoginButton 
      Caption         =   "Login"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox ServerURLTextBox 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Text            =   "http://192.168.1.118:8065/mattermost"
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox PasswordTextBox 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "SampleUs@r-1"
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox UsernameTextBox 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "user-1"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "server url:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "password:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "login:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoginButton_Click()
    MattermostClient.sServerURL = ServerURLTextBox.Text
    MattermostClient.Login UsernameTextBox.Text, PasswordTextBox.Text
    
    If MattermostClient.sLoginToken = "" Then
        MsgBox ("Failed to login")
    Else
        LoginForm.Hide
        MainChatForm.Show
    End If
End Sub
