Attribute VB_Name = "Impresora"
Option Explicit
    Public Const GENERIC_WRITE = &H40000000
    Public Const GENERIC_READ = &H80000000
    Public Const FILE_ATTRIBUTE_NORMAL = &H80
    Public Const CREATE_ALWAYS = 2
    Public Const OPEN_ALWAYS = 4
    Public Const INVALID_HANDLE_VALUE = -1
    
    Public Type COMSTAT
        Filler1 As Long
        Filler2 As Long
        Filler3 As Long
        Filler4 As Long
        Filler5 As Long
        Filler6 As Long
        Filler7 As Long
        Filler8 As Long
        Filler9 As Long
        Filler10 As Long
    End Type

    Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
    
    Public Const CE_BREAK = &H10 ' break condition
    Public Const CE_PTO = &H200 ' printer timeout
    Public Const CE_IOE = &H400 ' printer I/O error
    Public Const CE_DNS = &H800 ' device not selected
    Public Const CE_OOP = &H1000 ' out of paper

