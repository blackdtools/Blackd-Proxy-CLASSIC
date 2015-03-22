Attribute VB_Name = "modUpdater"
Option Explicit
'for help file
Public Const SW_NORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const cte_MakeFolderRequired As String = "1"
Public Const cte_NoUpdateRequired As String = "2"
Public Const cte_DownloadRequired As String = "3"
Public FilesToDownload(200, 3) As String
Public NumberOfFilesToDownload As Long
Public ScanDone As Boolean
Public dblTotalToDownload As Double
Public dblProgressSoFar As Double
Public dblProgressPortion As Double
