Attribute VB_Name = "modReplaceFiles"
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long

Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const MOVEFILE_COPY_ALLOWED = &H2
Public Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
Public Const MOVEFILE_DELAY_AND_REPLACE = MOVEFILE_DELAY_UNTIL_REBOOT + MOVEFILE_REPLACE_EXISTING

Public Sub DestroyOnReboot(strFileName As String, bolExit As Boolean)

    On Error Resume Next
    Dim strMain As String
    Dim strFile As String * 32768
    Dim intFile As Integer
    Dim strShortPathToEXE As String
    Dim strWinDir As String

      'Get short path to EXE
      strShortPathToEXE = GetShortPath(strFileName)

      If Environ("OS") = "Windows_NT" Then 'Win NT/2000/XP
          'This API tells NT to remove the file on reboot, using the following Registry key
          'HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\PendingfileRenameOperations

          MoveFileEx strShortPathToEXE, vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT

        Else 'Win9x/Me

          'Get WinDir
          strWinDir = Environ("WinDir")

          intFile = FreeFile

          If Dir$(strWinDir & "\WinInit.ini") = "" Then 'Create New WinInit file

              'Open file
              Open strWinDir & "\WinInit.ini" For Binary As intFile
              strMain = "[Rename]" & vbCrLf & "NUL=" & strShortPathToEXE & vbCrLf

              'Put data and close
              Put intFile, , strMain
              Close intFile
            Else 'WinInit exists, edit it
              'Open file
              Open strWinDir & "\WinInit.ini" For Binary As intFile

              'Get data and close
              Get intFile, , strFile
              Close intFile

              'Delete file
              Kill strWinDir & "\WinInit.ini"

              'Trim null chars
              strMain = Mid$(strFile, 1, InStr(1, strFile, Chr$(0)) - 1)

              'Open file again
              Open strWinDir & "\WinInit.ini" For Binary As intFile

              'Prepare data
              If Mid$(strMain, Len(strMain) - 1, 2) = vbCrLf Then 'WinInit file ends with vbCrLf
                  strMain = strMain & "NUL=" & strShortPathToEXE & vbCrLf
                Else 'WinInit file does not end with vbCrLf
                  strMain = strMain & vbCrLf & "NUL=" & strShortPathToEXE & vbCrLf
              End If

              'Put data and close
              Put intFile, , strMain
              Close intFile

          End If
      End If

      'Exit program
      If bolExit Then End

End Sub

Public Function GetShortPath(strFileName As String) As String

  Dim lngLength As Long, strPathBuffer As String

    'Create a buffer
    strPathBuffer = String$(165, 0)

    'retrieve the short pathname
    lngLength = GetShortPathName(strFileName, strPathBuffer, 255)

    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPathBuffer, lngLength)

End Function

':) Ulli's VB Code Formatter V2.10.8 (03.08.2002 01:29:55) 7 + 87 = 94 Lines
