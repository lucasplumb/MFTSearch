Attribute VB_Name = "modFormFuncs"
Option Explicit
Private m_driveLetters() As String

Public Property Get DriveLetters(index As Long) As String
    Let DriveLetters = m_driveLetters(index)
End Property

Public Function GetNTFSDrives() As String()
    Dim fs As Object, drive As Object, drives As Object, driveStr As String, retStr() As String, i As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set drives = fs.drives
    
    If drives.Count > 0 Then
        ReDim retStr(drives.Count - 1)
        ReDim m_driveLetters(drives.Count - 1)
    
        For Each drive In drives
            driveStr = vbNullString
            If drive.FileSystem = "NTFS" Then
                If drive.DriveType = 3 Then
                    driveStr = drive.ShareName 'network drive
                Else
                    driveStr = drive.VolumeName 'any other drive, most likely type 2 (fixed)
                End If
                If driveStr = vbNullString Then driveStr = "Local Disk"
                driveStr = driveStr & " (" & drive.driveLetter & ":)"
                m_driveLetters(i) = drive.driveLetter
                retStr(i) = driveStr
                i = i + 1
            End If
        Next
        
    End If
    GetNTFSDrives = retStr
End Function
