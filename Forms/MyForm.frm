VERSION 5.00
Begin VB.Form MyForm 
   Caption         =   "MyForm"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Btn1 
      Caption         =   "Write greek alphabet"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2700
   End
   Begin VB.Label Label2 
      Caption         =   "hDC: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "hWnd: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_hWnd As LongPtr
Private m_hDC  As LongPtr

#If VBA6 Or VBA7 Then
Private Sub UserForm_Activate()
#Else
Private Sub Form_Activate()
#End If
    Initialize
End Sub

Sub Initialize()
    m_hWnd = GetActiveWindow
    m_hDC = GetDC(m_hWnd)
    Label1.Caption = "hWnd: " & m_hWnd & " = &H" & Hex(m_hWnd)
    Label2.Caption = "hDC:  " & m_hDC & " = &H" & Hex(m_hDC)
End Sub

Private Sub Btn1_Click()
    
    Dim hr As Long
    Dim r As RECT: r = New_RECT(Btn1.Left, Btn1.Top + Btn1.Height + 30, 400, 40)
    hr = DrawFocusRect(m_hDC, r)
    
    Dim ga As String: ga = GetGreekAlphabet
    MsgBox ga
    MsgBoxW ga
    
    hr = DrawTextW(m_hDC, StrPtr(ga), Len(ga), r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    
End Sub

