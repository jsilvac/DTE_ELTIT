Attribute VB_Name = "Module1"
Const MAX_IP = 10
Type IPINFO
dwAddr As Long
dwIndex As Long
dwMask As Long
dwBCastAddr As Long
dwReasmSize As Long
unused1 As Integer
unused2 As Integer
End Type
Type MIB_IPADDRTABLE
dEntrys As Long
mIPInfo(MAX_IP) As IPINFO
End Type
Type IP_Array
mBuffer As MIB_IPADDRTABLE
BufferLen As Long
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function ConvertAddressToString(longAddr As Long) As String
Dim myByte(3) As Byte
Dim Cnt As Long
CopyMemory myByte(0), longAddr, 4
For Cnt = 0 To 3
ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
Next Cnt
ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function
Public Function GetWanIP() As String
Dim Ret As Long, Tel As Long
Dim bBytes() As Byte
Dim TempList() As String
Dim TempIP As String
Dim Tempi As Long
Dim Listing As MIB_IPADDRTABLE
Dim L3 As String
On Error GoTo END1
GetIpAddrTable ByVal 0&, Ret, True
If Ret <= 0 Then Exit Function
ReDim bBytes(0 To Ret - 1) As Byte
ReDim TempList(0 To Ret - 1) As String
GetIpAddrTable bBytes(0), Ret, False
CopyMemory Listing.dEntrys, bBytes(0), 4
For Tel = 0 To Listing.dEntrys - 1
CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
Next Tel
TempIP = TempList(0)
For Tempi = 0 To Listing.dEntrys - 1
L3 = Left(TempList(Tempi), 3)
If L3 = "192" Then
TempIP = TempList(Tempi)
End If
Next Tempi
GetWanIP = TempIP
Exit Function
END1:
GetWanIP = ""
End Function


