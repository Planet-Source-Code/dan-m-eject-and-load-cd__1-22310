Attribute VB_Name = "ioctrl"
'''''''''''''''''''''''''Eject/Load Removable Devices''''''''''''''''''''''''''''
'
' NAME
'     ioctrl.bas (Eject/Load Removable Drives)
'
' DESCRIPTION
'     Ejects and loads removable devices.  Works for any CD drive on your system.
'     Simply pass the functions the drive letter of the CD.  Note: I have not
'     tried this code with other removables other than CDs but Im assuming it
'     it should work.
'
' PLATFORM
'     WindowsNT/2000 (Unfortunately will not work with 95/98/me kernals)
'
' AUTHOR
'     Daniel Maroff 4/2001
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit


'API functions
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'API Constants
Const GENERIC_READ = &H80000000
Const FILE_FLAG_WRITE_THROUGH = &H80000000
Const GENERIC_WRITE = &H40000000
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2

Const CREATE_NEW = 1
Const CREATE_ALWAYS = 2
Const OPEN_EXISTING = 3
Const OPEN_ALWAYS = 4
Const TRUNCATE_EXISTING = 5


Const IOCTL_STORAGE_EJECT_MEDIA = 2967560
Const IOCTL_STORAGE_LOAD_MEDIA = 2967564



Public Function EjectDrive(driveLetter As String) As Boolean
  Dim hDisk As Long
  Dim dwRc As Long
  Dim sDisk As String
  
  'Generate a volume name
  sDisk = "\\.\" & UCase(Mid(driveLetter, 1, 1)) & ":"
  
  'We should get back a handle to the volume
  hDisk = CreateFile(sDisk, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
                     FILE_SHARE_WRITE, ByVal CLng(0), OPEN_EXISTING, FILE_FLAG_WRITE_THROUGH, _
                     ByVal CLng(0))
  
  If hDisk = -1 Then
    EjectDrive = False
    Exit Function
  End If
  
  'Clear any cache from the disk
  FlushFileBuffers (hDisk)
  
  'Control the device (in this case, eject it)
  If DeviceIoControl(hDisk, IOCTL_STORAGE_EJECT_MEDIA, _
                     ByVal CLng(0), 0, ByVal CLng(0), 0, _
                     dwRc, ByVal CLng(0)) = 0 Then
    EjectDrive = False
  Else
    EjectDrive = True
  End If
  
  'Always close your handles!
  CloseHandle (hDisk)
  
End Function


Public Function LoadDrive(driveLetter As String) As Boolean
  Dim hDisk As Long
  Dim dwRc As Long
  Dim sDisk As String
  
  'Generate a volume name
  sDisk = "\\.\" & UCase(Mid(driveLetter, 1, 1)) & ":"
  
  'We should get back a handle to the volume
  hDisk = CreateFile(sDisk, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
                     FILE_SHARE_WRITE, ByVal CLng(0), OPEN_EXISTING, FILE_FLAG_WRITE_THROUGH, _
                     ByVal CLng(0))
  
  If hDisk = -1 Then
    LoadDrive = False
    Exit Function
  End If
  
  'Clear any cache from the disk
  FlushFileBuffers (hDisk)
  
  'Control the device (in this case, eject it)
  If DeviceIoControl(hDisk, IOCTL_STORAGE_LOAD_MEDIA, _
                     ByVal CLng(0), 0, ByVal CLng(0), 0, _
                     dwRc, ByVal CLng(0)) = 0 Then
    LoadDrive = False
  Else
    LoadDrive = True
  End If
  
  'Always close your handles!
  CloseHandle (hDisk)
  
End Function


