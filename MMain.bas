Attribute VB_Name = "MMain"
Option Explicit

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Const DT_CENTER     As Long = &H1
Public Const DT_VCENTER    As Long = &H4
Public Const DT_SINGLELINE As Long = &H20

#If VBA7 Then
    Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function DrawTextW Lib "user32" (ByVal hdc As LongPtr, ByVal lpStr As LongPtr, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
    Public Declare PtrSafe Function DrawFocusRect Lib "user32.dll" (ByVal hDC As LongPtr, ByRef lpRect As RECT) As Long
    Public Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal wType As Long) As Long
'#End If
'#If VBC Then
#Else
    Public Enum LongPtr
        [_]
    End Enum
    Public Declare Function GetActiveWindow Lib "user32" () As LongPtr
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare Function DrawTextW Lib "user32" (ByVal hDC As LongPtr, ByVal lpStr As LongPtr, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
    Public Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As LongPtr, ByRef lpRect As RECT) As Long
    Public Declare Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal wType As Long) As Long
#End If

Sub Main()
    #If Mode_Beta Then
        InitModeBeta
    #ElseIf Mode_Debug Then
        InitModeDebug
    #Else 'every other release
        InitModeRelease
    #End If
    MyForm.Show
End Sub

Private Sub InitModeBeta()
    MsgBox "Init Mode Beta"
End Sub

Private Sub InitModeDebug()
    MsgBox "Init Mode Debug"
End Sub

Private Sub InitModeRelease()
    MsgBox "Init Mode Release"
End Sub

Function New_RECT(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As RECT
    New_RECT.Left = Left
    New_RECT.Right = Left + Width
    New_RECT.Top = Top
    New_RECT.Bottom = Top + Height
End Function

Property Get App_EXEName() As String
#If VBA6 Or VBA7 Then
    App_EXEName = Application.Name
#Else
    App_EXEName = App.EXEName
#End If
End Property

Function GetGreekAlphabet() As String
    Dim s As String
    Dim i As Long
    Dim alp As Long: alp = 913 'der Groﬂe griechische Buchstabe Alpha
    For i = alp To alp + 24
        s = s & ChrW(i)
    Next
    s = s & " "
    alp = alp + 32             'der Kleine griechische Buchstabe alpha
    For i = alp To alp + 24
        s = s & ChrW(i)
    Next
    GetGreekAlphabet = s
End Function

Public Function MsgBoxW(Prompt, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title) As VbMsgBoxResult
'Public Function MsgBoxW(Prompt, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title, Optional Helpfile, Optional Context) As VbMsgBoxResult
    Title = IIf(IsMissing(Title), App_EXEName, CStr(Title))
    MsgBoxW = MessageBoxW(0, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

