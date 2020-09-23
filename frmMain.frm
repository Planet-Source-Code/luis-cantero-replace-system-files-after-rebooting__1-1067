VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickDel"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    On Error Resume Next

    Dim strFileToDelete As String

      'If Command <> "123456" Then frmSplash.Show vbModal:  End

      strFileToDelete = "C:\Icdp\System\90799.exe"

      'TEST STRING
      'strFileToDelete = "C:\test\kl.exe"

      If Dir(strFileToDelete) <> "" Then Kill strFileToDelete

      Call DestroyOnReboot(App.Path & "\" & App.EXEName & ".exe", True)

End Sub

':) Ulli's VB Code Formatter V2.10.8 (03.08.2002 01:12:48) 0 + 21 = 21 Lines
