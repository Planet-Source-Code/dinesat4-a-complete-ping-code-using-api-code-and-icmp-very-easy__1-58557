Attribute VB_Name = "Module1"
'THIS IS A FREE CODE, USE IT :D
'by Patricio 'DiNeSat4' Tapia (patricio.tapia@gmail.com)
'I'm VERY BAD with English :(, sorry for the comments
Const SOCKET_ERROR = 0
Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long 'Retrieve host information to a host name
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long 'Initialize the winsock DLL
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long 'Terminate the Winsock DLL
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long) 'Copy a block of memory
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long 'Create a handle with ICMP request
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean 'Close a ICMP handle
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean 'Send a ICMP echo, and returns one or more replies.
Private Type WSAdata 'TYPE declarations for the WINSOCK
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Public Function makeping(ByVal HostName As String) As Boolean
    Dim hFile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Call WSAStartup(&H101, lpWSAdata)
    If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        MsgBox "Unable to Create File Handle", vbCritical + vbOKOnly
        makeping = False 'Return FALSE
        Exit Function
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        Form1.Text1 = Form1.Text1 & vbCrLf & "Timeout..." 'Send to TExt1
        makeping = False 'Return FALSE
    End If
    If EchoReply.Status = 0 Then
        Form1.Text1 = Form1.Text1 & vbCrLf & "Reply from " + HostName + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms." 'Send to TExt1
        makeping = True 'Return TRUE
    Else
        Form1.Text1 = Form1.Text1 & vbCrLf & "Ping to " & HostName & " has failure ..." 'Send to TExt1
        makeping = False 'Return FALSE
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function

