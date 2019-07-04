VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WS Usb Writer"
   ClientHeight    =   6576
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5748
   LinkTopic       =   "Form1"
   ScaleHeight     =   6576
   ScaleWidth      =   5748
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame4 
      Height          =   3132
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   5292
      Begin VB.OptionButton Option1 
         Caption         =   "N+1"
         Height          =   252
         Index           =   0
         Left            =   1800
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1092
      End
      Begin VB.OptionButton Option1 
         Caption         =   "N-1"
         Height          =   252
         Index           =   1
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton contion 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.ListBox lstResults 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   816
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   5052
      End
      Begin VB.TextBox Txt_Count 
         Height          =   372
         Left            =   1800
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   1212
      End
      Begin VB.Shape Shape_TagDetect 
         FillStyle       =   0  '實心
         Height          =   372
         Left            =   240
         Shape           =   3  '圓形
         Top             =   1440
         Width           =   372
      End
      Begin VB.Label lbl_TagAccessResult 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3840
         TabIndex        =   12
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Shape TagReadResult 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '實心
         Height          =   372
         Left            =   2040
         Shape           =   3  '圓形
         Top             =   1440
         Width           =   372
      End
      Begin VB.Label Label8 
         Caption         =   "Counter"
         Height          =   252
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label12 
         Caption         =   "Tag Detect"
         Height          =   492
         Left            =   720
         TabIndex        =   14
         Top             =   1440
         Width           =   1812
      End
      Begin VB.Label lbl_ReadWrite 
         Caption         =   "Access Result:"
         Height          =   372
         Left            =   2640
         TabIndex        =   13
         Top             =   1440
         Width           =   1452
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1452
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   5292
      Begin VB.CommandButton BTN_M_WriteEPC 
         Caption         =   "Write"
         Height          =   372
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox txt_EPC_Data_New 
         Height          =   372
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3612
      End
      Begin VB.Label Label5 
         Caption         =   "New EPC"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1692
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1452
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5292
      Begin VB.CommandButton Command1 
         Caption         =   "Read"
         Height          =   372
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox txt_EPC_Data 
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3612
      End
      Begin VB.Label Label4 
         Caption         =   "EPC"
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1692
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**系統名稱：RFIDUSB-HID通訊調試工具
'**描    述：
'**版    本：V1.0.0
'*************************************************************************

Option Explicit
'*********************************************
'
'*********************************************
Dim bAlertable As Long
Dim Capabilities As HIDP_CAPS
Dim DataString As String
Dim DetailData As Long
Dim DetailDataBuffer() As Byte
Dim DeviceAttributes As HIDD_ATTRIBUTES
Dim DevicePathName As String
Dim DeviceInfoSet As Long
Dim ErrorString As String
Dim EventObject As Long
Dim HIDHandle As Long
Dim HIDOverlapped As OVERLAPPED
Dim LastDevice As Boolean
Dim MyDeviceDetected As Boolean
Dim MyDeviceInfoData As SP_DEVINFO_DATA
Dim MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
Dim MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
Dim Needed As Long
Dim OutputReportData(7) As Byte
Dim PreparsedData As Long
Dim ReadHandle As Long
Dim Result As Long
Dim Security As SECURITY_ATTRIBUTES
Dim Timeout As Boolean
Dim DriveContion As Boolean
Public MyVendorID As Integer
Public MyProductID As Integer
Dim gTAG As tTAG
'DriveContion = False
Sub RF_Level_Init(Optional tITEM As String = "", Optional tITEM_Value As String = "00")

    Select Case tITEM
        Case "TX_LEVEL"
            'init Tx Level
            WriteReport ("1e 07 00 69  00 02 00 01 01 15 cb cb  cb cb cb cb")
            ReadReport
            WriteReport ("1f 07 00 68  00 02 00 01 15 " + tITEM_Value + " cb cb  cb cb cb cb")  ' 0b = -11
            ReadReport
    
        Case "RX_LEVEL"
            'init Rx Level 9a 0c 00 04  00 07 00 09 01 cd 00 00  00 00 00 cb
            WriteReport ("9a 0c 00 04  00 07 00 09  01  " + tITEM_Value + "  00 00  00 00 00 cb") 'cd = -51
            ReadReport
        Case Else
        
    End Select
    
    Debug.Print tITEM + " = " + tITEM_Value
End Sub
Function FindTheHid() As Boolean

'Makes a series of API calls to locate the desired HID-class device.
'Returns True if the device is detected, False if not detected.

Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

LastDevice = False
MyDeviceDetected = False

'Values for SECURITY_ATTRIBUTES structure:

Security.lpSecurityDescriptor = 0
Security.bInheritHandle = True
Security.nLength = Len(Security)

'******************************************************************************
'一、獲得HID設備的GUID.
'HidD_GetHidGuid
'Get the GUID for all system HIDs.
'Returns: the GUID in HidGuid.
'The routine doesn't return a value in Result
'but the routine is declared as a function for consistency with the other API calls.
'******************************************************************************

Result = HidD_GetHidGuid(HidGuid)
Call DisplayResultOfAPICall("GetHidGuid")

'Display the GUID.

GUIDString = _
    Hex$(HidGuid.Data1) & "-" & _
    Hex$(HidGuid.Data2) & "-" & _
    Hex$(HidGuid.Data3) & "-"

For Count = 0 To 7

    'Ensure that each of the 8 bytes in the GUID displays two characters.
    
    If HidGuid.Data4(Count) >= &H10 Then
        GUIDString = GUIDString & Hex$(HidGuid.Data4(Count)) & " "
    Else
        GUIDString = GUIDString & "0" & Hex$(HidGuid.Data4(Count)) & " "
    End If
Next Count

'lstResults.AddItem "  系統返回的 GUID號： " & GUIDString

'******************************************************************************
'二、找出所有已連接HID設備：
'SetupDiGetClassDevs
'Returns: a handle to a device information set for all installed devices.
'Requires: the HidGuid returned in GetHidGuid.
'******************************************************************************

DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
Call DisplayResultOfAPICall("SetupDiClassDevs（找所有HID設備）")
DataString = GetDataString(DeviceInfoSet, 32)

'******************************************************************************
'三、列舉每一個HID設備:
'SetupDiEnumDeviceInterfaces
'On return, MyDeviceInterfaceData contains the handle to a
'SP_DEVICE_INTERFACE_DATA structure for a detected device.
'Requires:
'the DeviceInfoSet returned in SetupDiGetClassDevs.
'the HidGuid returned in GetHidGuid.
'An index to specify a device.
'******************************************************************************

'Begin with 0 and increment until no more devices are detected.

MemberIndex = 0

Do
    'The cbSize element of the MyDeviceInterfaceData structure must be set to
    'the structure's size in bytes. The size is 28 bytes.
    
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
    
    Call DisplayResultOfAPICall("SetupDiEnumDeviceInterfaces")
    If Result = 0 Then LastDevice = True
    
    'If a device exists, display the information returned.
    
    If Result <> 0 Then
        
        'lstResults.AddItem "  DeviceInfoSet for device " & "找要的設備#" & CStr(MemberIndex) & ": "
         'list Device info on ListBox
  
        
        '******************************************************************************
        '四、取設備的路徑
        'SetupDiGetDeviceInterfaceDetail
        'Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure
        'containing information about a device.
        'To retrieve the information, call this function twice.
        'The first time returns the size of the structure in Needed.
        'The second time returns a pointer to the data in DeviceInfoSet.
        'Requires:
        'A DeviceInfoSet returned by SetupDiGetClassDevs and
        'an SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces.
        '*******************************************************************************
        
        MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           0, _
           0, _
           Needed, _
           0)
        
        DetailData = Needed
            
        Call DisplayResultOfAPICall("SetupDiGetDeviceInterfaceDetail(取設備路徑)")
        Debug.Print "  (OK to say too small)"
        Debug.Print "  Required buffer size for the data: " & Needed
        
        'Store the structure's size.
        
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        'Use a byte array to allocate memory for
        'the MyDeviceInterfaceDetailData structure
        
        ReDim DetailDataBuffer(Needed)
        
        'Store cbSize in the first four bytes of the array.
        
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        'Call SetupDiGetDeviceInterfaceDetail again.
        'This time, pass the address of the first element of DetailDataBuffer
        'and the returned required buffer size in DetailData.
        
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
        
        Call DisplayResultOfAPICall(" Result of second call:（第二次調用） ")
        Debug.Print "  MyDeviceInterfaceDetailData.cbSize: " & _
            CStr(MyDeviceInterfaceDetailData.cbSize)
        
        'Convert the byte array to a string.
        
        DevicePathName = CStr(DetailDataBuffer())
        
        'Convert to Unicode.
        
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        
        'Strip cbSize (4 bytes) from the beginning.
        
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        Debug.Print "  Device pathname: "
        Debug.Print "    " & DevicePathName
                
        '******************************************************************************
        '五、取得設備的標示代號:
        'CreateFile
        'Returns: a handle that enables reading and writing to the device.
        'Requires:
        'The DevicePathName returned by SetupDiGetDeviceInterfaceDetail.
        '******************************************************************************
    
        HIDHandle = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            Security, _
            OPEN_EXISTING, _
            0&, _
            0)
            
        Call DisplayResultOfAPICall("CreateFile（標示代號）")
        Debug.Print "  Returned handle: " & Hex$(HIDHandle) & "h"
        
        'Now we can find out if it's the device we're looking for.
        
        '******************************************************************************
        '取得廠商與產品ID：
        'HidD_GetAttributes
        'Requests information from the device.
        'Requires: The handle returned by CreateFile.
        'Returns: an HIDD_ATTRIBUTES structure containing
        'the Vendor ID, Product ID, and Product Version Number.
        'Use this information to determine if the detected device
        'is the one we're looking for.
        '******************************************************************************
        
        'Set the Size property to the number of bytes in the structure.
        
        DeviceAttributes.Size = LenB(DeviceAttributes)
        Result = HidD_GetAttributes _
            (HIDHandle, _
            DeviceAttributes)
            
        Call DisplayResultOfAPICall("HidD_GetAttributes（取PID,VID）")
        'Call DisplayResultOfAPICall("HidD_GetAttributes（" + MyProductID + "," + MyVendorID + "）")
        If Result <> 0 Then
            Debug.Print "  HIDD_ATTRIBUTES structure filled without error."
        Else
            Debug.Print "  Error in filling HIDD_ATTRIBUTES structure."
        End If
    
        'debug.print  "  Structure size: " & DeviceAttributes.Size
        Debug.Print "  Vendor ID: " & Hex$(DeviceAttributes.VendorID)
        Debug.Print "  Product ID: " & Hex$(DeviceAttributes.ProductID)
        'debug.print  "  Version Number: " & Hex$(DeviceAttributes.VersionNumber)
        
        'Find out if the device matches the one we're looking for.
        
        If (DeviceAttributes.VendorID = MyVendorID) And _
            (DeviceAttributes.ProductID = MyProductID) Then
                
                'It's the desired device.
                
                Debug.Print "  device found!！"
                MyDeviceDetected = True
                DriveContion = True
        Else
                MyDeviceDetected = False
                
                'If it's not the one we want, close its handle.
                
                Result = CloseHandle _
                    (HIDHandle)
                DisplayResultOfAPICall ("CloseHandle（關閉此接口）")
        End If
End If
    
    'Keep looking until we find the device or there are no more left to examine.
    
    MemberIndex = MemberIndex + 1
Loop Until (LastDevice = True) Or (MyDeviceDetected = True)

'Free the memory reserved for the DeviceInfoSet returned by SetupDiGetClassDevs.

Result = SetupDiDestroyDeviceInfoList _
    (DeviceInfoSet)
Call DisplayResultOfAPICall("DestroyDeviceInfoList（釋放資源）")

If MyDeviceDetected = True Then
    FindTheHid = True
    
    'Learn the capabilities of the device
     
     Call GetDeviceCapabilities
    
    'Get another handle for the overlapped ReadFiles.
    
    ReadHandle = CreateFile _
            (DevicePathName, _
            (GENERIC_READ Or GENERIC_WRITE), _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            Security, _
            OPEN_EXISTING, _
            FILE_FLAG_OVERLAPPED, _
            0)
 
    Call DisplayResultOfAPICall("CreateFile, ReadHandle")
    Debug.Print "  Returned handle: " & Hex$(ReadHandle) & "h"
    Call PrepareForOverlappedTransfer
    'Label3.Caption = "設備聯接成功"
    contion.BackColor = vbGreen
    contion.Caption = "Connected"

Else
    Debug.Print " device not found。"
    'Label3.Caption = "設備聯接失敗"
    contion.BackColor = vbYellow
    contion.Caption = "Connect Fail"
End If

End Function

Private Function GetDataString _
    (Address As Long, _
    Bytes As Long) _
As String

'Retrieves a string of length Bytes from memory, beginning at Address.
'Adapted from Dan Appleman's "Win32 API Puzzle Book"

Dim Offset As Integer
Dim Result$
Dim ThisByte As Byte

For Offset = 0 To Bytes - 1
    Call RtlMoveMemory(ByVal VarPtr(ThisByte), ByVal Address + Offset, 1)
    If (ThisByte And &HF0) = 0 Then
        Result$ = Result$ & "0"
    End If
    Result$ = Result$ & Hex$(ThisByte) & " "
Next Offset

GetDataString = Result$
End Function

Private Function GetErrorString _
    (ByVal LastError As Long) _
As String

'Returns the error message for the last error.
'Adapted from Dan Appleman's "Win32 API Puzzle Book"

Dim Bytes As Long
Dim ErrorString As String
ErrorString = String$(129, 0)
Bytes = FormatMessage _
    (FORMAT_MESSAGE_FROM_SYSTEM, _
    0&, _
    LastError, _
    0, _
    ErrorString$, _
    128, _
    0)
    
'Subtract two characters from the message to strip the CR and LF.

If Bytes > 2 Then
    GetErrorString = Left$(ErrorString, Bytes - 2)
End If

End Function


Private Sub DisplayResultOfAPICall(FunctionName As String)

'Display the results of an API call.

Dim ErrorString As String

Debug.Print ""
ErrorString = GetErrorString(Err.LastDllError)
Debug.Print FunctionName
Debug.Print "  Result = " & ErrorString

'Scroll to the bottom of the list box.

lstResults.ListIndex = lstResults.ListCount - 1

End Sub

Private Sub btn_Clear_Click()
lstResults.Clear
'Form1.TextIR.Text = ""
End Sub

Private Sub btn_Close_Click()
    Call Shutdown
End Sub

Private Sub BTN_M_Read_Click()
Dim TagScanCount As Integer
'Call ReadReport
'Call ReadReport("ReadTag")
'WriteReport (SelectEPC)
'Sleep (100)
'Call ReadReport
Debug.Print "init RF TX"
'RF_Level_Init

If (0) Then
    WriteReport ("1d 07 00 69  00 02 00 01  01 0a cb cb  cb cb cb cb") '0a
    ReadUSB_Report
    
    WriteReport ("1e 07 00 69  00 02 00 01  01 09 cb cb  cb cb cb cb") '09
    ReadUSB_Report
    
    WriteReport ("1f 07 00 69  00 02 00 01  01 0d cb cb  cb cb cb cb") '0d
    ReadUSB_Report
End If
'Start Scan
WriteReport ("20 11 00 86  00 02 00 00  00 0d 8c 00  05 00 00 01  01 00 01 06  cb cb cb cb  cb cb cb cb  cb cb cb cb")
ReadUSB_Report ("EmptyEPCTextField")

gTAG.TagCOUNT = 0
TagScanCount = 0

Me.Shape_TagDetect.FillColor = &H0& 'black
While ((gTAG.TagCOUNT < 10) And (TagScanCount < 10))
    ReadUSB_Report ("ScanNewTag")
    TagScanCount = TagScanCount + 1
Wend

'StopScan
WriteReport ("21 0a 00 8c  00 05 00 00  01 00 00 00  00 cb cb cb")
ReadUSB_Report


End Sub
Function SelectEPC()
    Dim vSelectEPC As String
    vSelectEPC = READER_T_EPC + Form1.txt_EPC_Data.Text + "AA"
    Debug.Print vSelectEPC
    SelectEPC = vSelectEPC

End Function

Private Sub BTN_M_ReadEPC_Click()

    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_EPC)
        ReadUSB_Report ("ReadEPC")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_ReadReserved_Click()
    
    
    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_RSV)
        ReadUSB_Report ("ReadReserved")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
    
End Sub

Private Sub BTN_M_ReadTID_Click()
    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_TID)
        ReadUSB_Report ("ReadTID")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
    

End Sub

Private Sub BTN_M_ReadUser_Click()
    gTAG.ACC_RESULT = "00"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "DE"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        
        WriteReport (READER_R_USR)
        ReadUSB_Report ("ReadUser")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_Write_Click()
    WriteReport (SelectEPC)
    ReadUSB_Report
    Debug.Print "Select Tag'EPC for Writing prepare"
End Sub

Private Sub BTN_M_WriteEPC_Click()
    If txt_EPC_Data_New.Text = "" Then
        MsgBox "Please input new EPC number"
        txt_EPC_Data_New.SetFocus
        Exit Sub
    End If
    
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    gTAG.DATAFORMAT_ERR = False
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00") And (gTAG.DATAFORMAT_ERR = False))
        'Debug.Print SelectEPC
        WriteReport (SelectEPC)
        ReadUSB_Report
        DataStr = READER_W_EPC + "00000000" + Me.txt_EPC_Data_New.Text
        Debug.Print DataStr
        WriteReport (DataStr)
        Sleep (30)
        ReadUSB_Report ("ReadEPC_AfterWrote")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
    
    If txt_EPC_Data_New.Text = txt_EPC_Data.Text Then
        If Txt_Count.Text = 0 Then
            lstResults.AddItem "Start: " & txt_EPC_Data_New.Text
        End If
        
        If Option1(0).Value Then
            txt_EPC_Data_New.Text = Format(txt_EPC_Data_New.Text + 1, "000000000000000000000000")
            Txt_Count = Txt_Count.Text + 1
        ElseIf Option1(1).Value Then
            txt_EPC_Data_New.Text = Format(txt_EPC_Data_New.Text - 1, "000000000000000000000000")
            Txt_Count = Txt_Count.Text + 1
        End If
        lbl_TagAccessResult.Caption = "Down"
        
    Else
    
    End If
End Sub

Private Sub BTN_M_WriteReservedAccess_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        'DataStr = READER_W_RSV_ACCESS_PW + Me.txt_Reserved_AccessPW.Text + Me.txt_Reserved_AccessPW_New.Text
        ''''DataStr = READER_W_RSV_ACCESS_PW + "00000000" + Me.txt_Reserved_AccessPW_New.Text
        WriteReport (DataStr)
        ReadUSB_Report ("ReadReservedAfterUpdateAccessPW")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_WriteReservedKill_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        'DataStr = READER_W_RSV_ACCESS_PW + Me.txt_Reserved_AccessPW.Text + Me.txt_Reserved_AccessPW_New.Text
        ''''DataStr = READER_W_RSV_KILL_PW + "00000000" + Me.txt_Reserved_KillPW_New.Text
        WriteReport (DataStr)
        ReadUSB_Report ("ReadReservedAfterUpdateKillPW")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub BTN_M_WriteUser_Click()
    Dim DataStr As String
    gTAG.ACC_RESULT = "FF"
    gTAG.ACC_COUNT = 0
    
    'MsgBox "UserNewLength =" & Len(Me.txt_User_New.Text)
    
    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00"))
        WriteReport (SelectEPC)
        ReadUSB_Report
        ''''DataStr = READER_W47B_USR + "00000000" + Mid(Me.txt_User_New.Text, 1, 94)
        WriteReport (DataStr)
        Sleep (30)
        
        ''''DataStr = READER_W17B_USR + Mid(Me.txt_User_New.Text, 95, 34)
        WriteReport (DataStr)
        
        ReadUSB_Report ("ReadUser_AfterWrote64B")
        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
    Wend
End Sub

Private Sub btn_RF_init_Click()
''''    'init Tx Level
''''  If (1) Then 'Tx : -11 , Rx : -51
''''    Call RF_Level_Init("TX_LEVEL", "0B")
''''    Call RF_Level_Init("RX_LEVEL", "CD")
''''    Me.HS_TxPower.Value = -11
''''    Me.HS_RxSensetivity.Value = -51
''''  Else         'Tx : 0  , Rx : -70
''''    Call RF_Level_Init("TX_LEVEL", "01")
''''    Call RF_Level_Init("RX_LEVEL", "BA")
''''    Me.HS_TxPower.Value = -1
''''    Me.HS_RxSensetivity.Value = -70
''''
''''  End If
End Sub

Private Sub btn_ScanStart_Click()
WriteReport (READER_T_STARTSCAN)
ReadReport
Form1.Timer1.Enabled = True
End Sub

Private Sub btn_ScanStop_Click()
WriteReport (READER_T_STOPSCAN)
ReadReport
Form1.Timer1.Enabled = False
End Sub

Private Sub btn_Test_Click()
WriteReport (READER_T_TEST)
ReadReport
End Sub


Private Sub Command1_Click()
Dim TagScanCount As Integer
'Call ReadReport
'Call ReadReport("ReadTag")
'WriteReport (SelectEPC)
'Sleep (100)
'Call ReadReport
Debug.Print "init RF TX"
'RF_Level_Init

If (0) Then
    WriteReport ("1d 07 00 69  00 02 00 01  01 0a cb cb  cb cb cb cb") '0a
    ReadUSB_Report
    
    WriteReport ("1e 07 00 69  00 02 00 01  01 09 cb cb  cb cb cb cb") '09
    ReadUSB_Report
    
    WriteReport ("1f 07 00 69  00 02 00 01  01 0d cb cb  cb cb cb cb") '0d
    ReadUSB_Report
End If
'Start Scan
WriteReport ("20 11 00 86  00 02 00 00  00 0d 8c 00  05 00 00 01  01 00 01 06  cb cb cb cb  cb cb cb cb  cb cb cb cb")
ReadUSB_Report ("EmptyEPCTextField")

gTAG.TagCOUNT = 0
TagScanCount = 0

Me.Shape_TagDetect.FillColor = &H0& 'black
While ((gTAG.TagCOUNT < 10) And (TagScanCount < 10))
    ReadUSB_Report ("ScanNewTag")
    TagScanCount = TagScanCount + 1
Wend

'StopScan
WriteReport ("21 0a 00 8c  00 05 00 00  01 00 00 00  00 cb cb cb")
ReadUSB_Report


End Sub

'Private Sub Command2_Click()
'    Dim DataStr As String
'    gTAG.ACC_RESULT = "FF"
'    gTAG.ACC_COUNT = 0
'    gTAG.DATAFORMAT_ERR = False
'
'    While ((gTAG.ACC_COUNT < 10) And (gTAG.ACC_RESULT <> "00") And (gTAG.DATAFORMAT_ERR = False))
'        WriteReport (SelectEPC)
'        ReadUSB_Report
'        DataStr = READER_W_EPC + "00000000" + Me.txt_EPC_Data_New.Text
'        WriteReport (DataStr)
'        Sleep (30)
'        ReadUSB_Report ("ReadEPC_AfterWrote")
'        gTAG.ACC_COUNT = gTAG.ACC_COUNT + 1
'        Debug.Print "gTAG.ACC_COUNT=" + CStr(gTAG.ACC_COUNT)
'    Wend
'
'End Sub
'
Private Sub contion_Click()

CheckInput:
    If MyVendorID = 0 Or MyProductID = 0 Then
        Form2.Show 1, Me
        MyVendorID = Form2.MyVendorIDm
        MyProductID = Form2.MyProductID
        GoTo CheckInput
    End If

FindTheHid
    
   
End Sub


Private Sub Form_Load()
Form1.Show
Call contion_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Shutdown
End Sub

Private Sub GetDeviceCapabilities()

'******************************************************************************
'HidD_GetPreparsedData
'Returns: a pointer to a buffer containing information about the device's capabilities.
'Requires: A handle returned by CreateFile.
'There's no need to access the buffer directly,
'but HidP_GetCaps and other API functions require a pointer to the buffer.
'******************************************************************************

Dim ppData(29) As Byte
Dim ppDataString As Variant

'Preparsed Data is a pointer to a routine-allocated buffer.

Result = HidD_GetPreparsedData _
    (HIDHandle, _
    PreparsedData)
Call DisplayResultOfAPICall("HidD_GetPreparsedData")

'Copy the data at PreparsedData into a byte array.

Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)
Call DisplayResultOfAPICall("RtlMoveMemory")

ppDataString = ppData()

'Convert the data to Unicode.

ppDataString = StrConv(ppDataString, vbUnicode)

'******************************************************************************
'HidP_GetCaps
'Find out the device's capabilities.
'For standard devices such as joysticks, you can find out the specific
'capabilities of the device.
'For a custom device, the software will probably know what the device is capable of,
'so this call only verifies the information.
'Requires: The pointer to a buffer containing the information.
'The pointer is returned by HidD_GetPreparsedData.
'Returns: a Capabilites structure containing the information.
'******************************************************************************
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)

Call DisplayResultOfAPICall("HidP_GetCaps")
Debug.Print "  Last error: " & ErrorString
Debug.Print "  Usage: " & Hex$(Capabilities.Usage)
Debug.Print "  Usage Page: " & Hex$(Capabilities.UsagePage)
Debug.Print "  Input Report Byte Length: " & Capabilities.InputReportByteLength
Debug.Print "  Output Report Byte Length: " & Capabilities.OutputReportByteLength
Debug.Print "  Feature Report Byte Length: " & Capabilities.FeatureReportByteLength
Debug.Print "  Number of Link Collection Nodes: " & Capabilities.NumberLinkCollectionNodes
Debug.Print "  Number of Input Button Caps: " & Capabilities.NumberInputButtonCaps
Debug.Print "  Number of Input Value Caps: " & Capabilities.NumberInputValueCaps
Debug.Print "  Number of Input Data Indices: " & Capabilities.NumberInputDataIndices
Debug.Print "  Number of Output Button Caps: " & Capabilities.NumberOutputButtonCaps
Debug.Print "  Number of Output Value Caps: " & Capabilities.NumberOutputValueCaps
Debug.Print "  Number of Output Data Indices: " & Capabilities.NumberOutputDataIndices
Debug.Print "  Number of Feature Button Caps: " & Capabilities.NumberFeatureButtonCaps
Debug.Print "  Number of Feature Value Caps: " & Capabilities.NumberFeatureValueCaps
Debug.Print "  Number of Feature Data Indices: " & Capabilities.NumberFeatureDataIndices

'******************************************************************************
'HidP_GetValueCaps
'Returns a buffer containing an array of HidP_ValueCaps structures.
'Each structure defines the capabilities of one value.
'This application doesn't use this data.
'******************************************************************************

'This is a guess. The byte array holds the structures.

Dim ValueCaps(1023) As Byte

Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
   
Call DisplayResultOfAPICall("HidP_GetValueCaps")

'debug.print  "ValueCaps= " & GetDataString((VarPtr(ValueCaps(0))), 180)
'To use this data, copy the byte array into an array of structures.

'Free the buffer reserved by HidD_GetPreparsedData

Result = HidD_FreePreparsedData _
    (PreparsedData)
Call DisplayResultOfAPICall("HidD_FreePreparsedData")

End Sub


Private Sub PrepareForOverlappedTransfer()

'******************************************************************************
'CreateEvent
'Creates an event object for the overlapped structure used with ReadFile.
'Requires a security attributes structure or null,
'Manual Reset = True (ResetEvent resets the manual reset object to nonsignaled),
'Initial state = True (signaled),
'and event object name (optional)
'Returns a handle to the event object.
'******************************************************************************

If EventObject = 0 Then
    EventObject = CreateEvent _
        (Security, _
        True, _
        True, _
        "")
End If
    
Call DisplayResultOfAPICall("CreateEvent")
    
'Set the members of the overlapped structure.

HIDOverlapped.Offset = 0
HIDOverlapped.OffsetHigh = 0
HIDOverlapped.hEvent = EventObject
End Sub



Private Sub Shutdown()

'Actions that must execute when the program ends.

'Close the open handles to the device.

Result = CloseHandle _
    (HIDHandle)
Call DisplayResultOfAPICall("CloseHandle (HIDHandle)")

Result = CloseHandle _
    (ReadHandle)
Call DisplayResultOfAPICall("CloseHandle (ReadHandle)")

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub HS_RxSensetivity_Change()
''''    Dim RxLevel As String
''''    RxLevel = Hex$(CStr(256 + Form1.HS_RxSensetivity.Value))
''''
''''    ''''Form1.txt_RX_Sensetivity.Text = Form1.HS_RxSensetivity.Value
''''    Call RF_Level_Init("RX_LEVEL", RxLevel)
End Sub

Private Sub HS_TxPower_Change()
''''    Dim TxPower As String
''''    Dim bb As Byte
''''    'TxPower = CStr(0 - Form1.HS_TxPower.Value)
''''    TxPower = Hex$(CStr(0 - Form1.HS_TxPower.Value))
''''    ''''Form1.txt_TX_Power.Text = Form1.HS_TxPower.Value
''''    If (Len(TxPower) < 2) Then
''''        TxPower = "0" + TxPower
''''    End If
''''    Call RF_Level_Init("TX_LEVEL", TxPower) ' Min(19) ~ Max(0)


End Sub

Private Sub Iread_Click()
Call ReadReport
End Sub
Private Sub ReadReport(Optional RdCmdStr As String = "")

    'Read data from the device.
    
    Dim Count
    Dim NumberOfBytesRead As Long
    
    'Allocate a buffer for the report.
    'Byte 0 is the report ID.
    
    Dim ReadBuffer() As Byte
    Dim UBoundReadBuffer As Integer
    
    '******************************************************************************
    'ReadFile
    'Returns: the report in ReadBuffer.
    'Requires: a device handle returned by CreateFile
    '(for overlapped I/O, CreateFile must be called with FILE_FLAG_OVERLAPPED),
    'the Input report length in bytes returned by HidP_GetCaps,
    'and an overlapped structure whose hEvent member is set to an event object.
    '******************************************************************************
    
    Dim ByteValue As String
    Dim ByteValueOutput As String
    
    If MyDeviceDetected = False And DriveContion = False Then
          Debug.Print "讀出失敗請先聯接"
          Exit Sub
          End If
          
    
          
    'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.
    
    ReDim ReadBuffer(Capabilities.InputReportByteLength - 1)
    
    'Scroll to the bottom of the list box.
    
    lstResults.ListIndex = lstResults.ListCount - 1
    
    'Do an overlapped ReadFile.
    'The function returns immediately, even if the data hasn't been received yet.
    
    Result = ReadFile _
        (ReadHandle, _
        ReadBuffer(0), _
        CLng(Capabilities.InputReportByteLength), _
        NumberOfBytesRead, _
        HIDOverlapped)
    Call DisplayResultOfAPICall("ReadFile")
    
    Debug.Print "waiting for ReadFile"
    
    'Scroll to the bottom of the list box.
    
    lstResults.ListIndex = lstResults.ListCount - 1
    bAlertable = True
    
    '******************************************************************************
    'WaitForSingleObject
    'Used with overlapped ReadFile.
    'Returns when ReadFile has received the requested amount of data or on timeout.
    'Requires an event object created with CreateEvent
    'and a timeout value in milliseconds.
    '******************************************************************************
    Result = WaitForSingleObject _
        (EventObject, _
        100)
    Call DisplayResultOfAPICall("WaitForSingleObject")
    
    'Find out if ReadFile completed or timeout.
    
    Select Case Result
        Case WAIT_OBJECT_0
            
            'ReadFile has completed
            
            Debug.Print "ReadFile completed successfully."
        Case WAIT_TIMEOUT
            
            'Timeout
            
            Debug.Print "Readfile timeout"
            
            'Cancel the operation
            
            '*************************************************************
            'CancelIo
            'Cancels the ReadFile
            'Requires the device handle.
            'Returns non-zero on success.
            '*************************************************************
            Result = CancelIo _
                (ReadHandle)
            Debug.Print "************ReadFile timeout*************"
            Debug.Print "CancelIO"
            Call DisplayResultOfAPICall("CancelIo")
            
            'The timeout may have been because the device was removed,
            'so close any open handles and
            'set MyDeviceDetected=False to cause the application to
            'look for the device on the next attempt.
            
            'CloseHandle (HIDHandle)
            'Call DisplayResultOfAPICall("CloseHandle (HIDHandle)")
            'CloseHandle (ReadHandle)
            'Call DisplayResultOfAPICall("CloseHandle (ReadHandle)")
            'MyDeviceDetected = False
        Case Else
            Debug.Print "Readfile undefined error"
            MyDeviceDetected = False
    End Select
        
    Debug.Print " Report ID: " & ReadBuffer(0)
    Debug.Print " Report Data:"
    
    ''Form1.TextIR.Text = ""
    'ByteValueOutput = ""
    
    For Count = 1 To UBound(ReadBuffer)
        
        'Add a leading 0 to values 0 - Fh.
        
        If Len(Hex$(ReadBuffer(Count))) < 2 Then
            ByteValue = "0" & Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        Else
            ByteValue = Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        End If
        
        'If (Len(ByteValueOutput) Mod 128) = 0 Then
            'ByteValueOutput = ByteValueOutput + vbCrLf
            'ByteValueOutput = ByteValueOutput + Chr(13) + Chr(10)
        'End If
    
    Next Count
    

    'print log to TextIR text box
    
    
        
        Select Case RdCmdStr
          Case "ReadTag"
                    gTAG.EPC = Mid(ByteValueOutput, 39, 24)
                    Form1.txt_EPC_Data.Text = gTAG.EPC

          Case Else
                    'Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
        End Select
    
    '******************************************************************************
    'ResetEvent
    'Sets the event object in the overlapped structure to non-signaled.
    'Requires a handle to the event object.
    'Returns non-zero on success.
    '******************************************************************************
    
    Call ResetEvent(EventObject)
    Call DisplayResultOfAPICall("ResetEvent")

End Sub
Private Sub ReadUSB_Report(Optional RdCmdStr As String = "")

    'Read data from the device.
    
    Dim Count
    Dim NumberOfBytesRead As Long
    
    'Allocate a buffer for the report.
    'Byte 0 is the report ID.
    
    Dim ReadBuffer() As Byte
    Dim UBoundReadBuffer As Integer
    
    '******************************************************************************
    'ReadFile
    'Returns: the report in ReadBuffer.
    'Requires: a device handle returned by CreateFile
    '(for overlapped I/O, CreateFile must be called with FILE_FLAG_OVERLAPPED),
    'the Input report length in bytes returned by HidP_GetCaps,
    'and an overlapped structure whose hEvent member is set to an event object.
    '******************************************************************************
    
    Dim ByteValue As String
    Dim ByteValueOutput As String
    
    Dim TagReadResult As String '標籤讀取結果
    Debug.Print RdCmdStr
    If MyDeviceDetected = False And DriveContion = False Then
          Debug.Print "讀出失敗請先聯接"
          Exit Sub
    End If
          
    
          
    'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.
    
    ReDim ReadBuffer(Capabilities.InputReportByteLength - 1)
    
    
    'Do an overlapped ReadFile.
    'The function returns immediately, even if the data hasn't been received yet.
    
    Result = ReadFile _
        (ReadHandle, _
        ReadBuffer(0), _
        CLng(Capabilities.InputReportByteLength), _
        NumberOfBytesRead, _
        HIDOverlapped)
    
    Debug.Print "waiting for ReadFile"
    
    'Scroll to the bottom of the list box.
    
    bAlertable = True
    
    '******************************************************************************
    'WaitForSingleObject
    'Used with overlapped ReadFile.
    'Returns when ReadFile has received the requested amount of data or on timeout.
    'Requires an event object created with CreateEvent
    'and a timeout value in milliseconds.
    '******************************************************************************
    Result = WaitForSingleObject _
        (EventObject, _
        100)
    Call DisplayResultOfAPICall("WaitForSingleObject")
    
    'Find out if ReadFile completed or timeout.
    
    Select Case Result
        Case WAIT_OBJECT_0
            
            'ReadFile has completed
            
            Debug.Print "ReadFile completed successfully."
        Case WAIT_TIMEOUT
            
            'Timeout
            
            Debug.Print "Readfile timeout"
            
            'Cancel the operation
            
            '*************************************************************
            'CancelIo
            'Cancels the ReadFile
            'Requires the device handle.
            'Returns non-zero on success.
            '*************************************************************
            Result = CancelIo _
                (ReadHandle)
            Debug.Print "************ReadFile timeout*************"
            Debug.Print "CancelIO"
            Call DisplayResultOfAPICall("CancelIo")
            
            'The timeout may have been because the device was removed,
            'so close any open handles and
            'set MyDeviceDetected=False to cause the application to
            'look for the device on the next attempt.
            
            'CloseHandle (HIDHandle)
            'Call DisplayResultOfAPICall("CloseHandle (HIDHandle)")
            'CloseHandle (ReadHandle)
            'Call DisplayResultOfAPICall("CloseHandle (ReadHandle)")
            'MyDeviceDetected = False
        Case Else
            Debug.Print "Readfile undefined error"
            MyDeviceDetected = False
    End Select
        
    Debug.Print " Report  : Result : " & Result & " ,ID: " & ReadBuffer(0)
    'Debug.Print " Report Data:"
    
    ''Form1.TextIR.Text = ""
    'ByteValueOutput = ""
    
    For Count = 1 To UBound(ReadBuffer)
        
        'Add a leading 0 to values 0 - Fh.
        
        If Len(Hex$(ReadBuffer(Count))) < 2 Then
            ByteValue = "0" & Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        Else
            ByteValue = Hex$(ReadBuffer(Count))
            ByteValueOutput = ByteValueOutput + ByteValue
        End If
    Next Count
    
        TagReadResult = "00"
        Me.TagReadResult.FillColor = &HC0C0C0
    
        Select Case RdCmdStr
          Case "EmptyEPCTextField"
                    Form1.txt_EPC_Data.Text = ""
          Case "ReadReserved"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'Form1.txt_Reserved_AccessPW.Text = ""
                    'Form1.txt_Reserved_KillPW.Text = ""
                    
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.RESERVED = Mid(ByteValueOutput, 17, 16)
                        'Form1.txt_Reserved_KillPW.Text = Mid(ByteValueOutput, 17, 8)
                        'Form1.txt_Reserved_AccessPW.Text = Mid(ByteValueOutput, 25, 8)
                        'Form1.txt_Reserved_KillPW.Text = Mid(ByteValueOutput, 17, 8)
                        'Form1.txt_Reserved_AccessPW.Text = Mid(ByteValueOutput, 25, 8)
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        
                    End If
                    Debug.Print "Tag Reserved Result=" + gTAG.ACC_RESULT
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadReservedAfterUpdateAccessPW"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'Form1.txt_Reserved_AccessPW.Text = ""
                    'Form1.txt_Reserved_KillPW.Text = ""
                    
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "00") Then
                       
                        gTAG.RESERVED = Mid(ByteValueOutput, 17, 16)
                        'Form1.txt_Reserved_AccessPW.Text = Form1.txt_Reserved_AccessPW_New
                        'Form1.txt_Reserved_KillPW.Text = Form1.txt_Reserved_KillPW_New.Text
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        
                    End If
                    Debug.Print "Tag Reserved Result=" + gTAG.ACC_RESULT
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadReservedAfterUpdateKillPW"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'Form1.txt_Reserved_AccessPW.Text = ""
                    'Form1.txt_Reserved_KillPW.Text = ""
                    
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "00") Then
                       
                        gTAG.RESERVED = Mid(ByteValueOutput, 17, 16)
                        'Form1.txt_Reserved_AccessPW.Text = Form1.txt_Reserved_AccessPW_New
                        'Form1.txt_Reserved_KillPW.Text = Form1.txt_Reserved_KillPW_New.Text
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        
                    End If
                    Debug.Print "Tag Reserved Result=" + gTAG.ACC_RESULT
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
                    
          Case "ReadEPC"
                    '2816000800DE001193B23400E2003000390701110610D48103055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'Form1.txt_EPC_Data = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.EPC = Mid(ByteValueOutput, 25, 24)
                        Form1.txt_EPC_Data = gTAG.EPC
                        Me.TagReadResult.FillColor = &HFF00&
                    Else
                        
                    End If
                    Debug.Print "Tag EPC Result=" + gTAG.ACC_RESULT
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadEPC_AfterWrote"
                    '2816000800DE001193B23400E2003000390701110610D48103055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    'Form1.txt_EPC_Data = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "00") Then
                       
                        'f30700070000000206a5bbbb11111111033400e2003000390701110610d4827c

                        gTAG.EPC = Mid(ByteValueOutput, 39, 24)
                        Form1.txt_EPC_Data = Form1.txt_EPC_Data_New
                        Me.TagReadResult.FillColor = &HFF00&
                        Debug.Print "ByteValueOutput = " + ByteValueOutput
                        'Debug.Print "Result check:" & Mid(ByteValueOutput, 3, 8)
                        'Can't just confirm  byte_6, also confirm byte2 , byte4 should as 07 07
                        If (Mid(ByteValueOutput, 3, 8) <> "07000700") Then
                            gTAG.ACC_RESULT = "FF"
                        End If
                    Else
                        
                    End If
                    Debug.Print "Tag EPC Result=" + gTAG.ACC_RESULT
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          Case "ReadTID"
                    '2D1E000800DE0019E2003412012EF8000686D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    ''Form1.txt_TID = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.TID = Mid(ByteValueOutput, 17, 48)
                        ''Form1.txt_TID = gTAG.TID
                        Me.TagReadResult.FillColor = &HFF00&
                    Else
                        
                    End If
                    Debug.Print "Tag TID Result=" + gTAG.ACC_RESULT
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          Case "ReadUser"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    ''Form1.txt_User = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.USER = Mid(ByteValueOutput, 17, 112)
                        ''Form1.txt_User = gTAG.USER
                        Me.TagReadResult.FillColor = &HFF00&
                        Debug.Print "Mid(ByteValueOutput, 15, 2)= " + Mid(ByteValueOutput, 15, 2)
                        gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 15, 2))
                        ReadUSB_Report ("ReadUserPart2")
                    Else
                        
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadUserPart2"
                    gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 3, 2))
                    If (gTAG.DATA_LENGTH > 0) Then
                        'gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 3, 2))
                        gTAG.USER = gTAG.USER + Mid(ByteValueOutput, 7, 16)
                        ''Form1.txt_User = gTAG.USER
                        Me.TagReadResult.FillColor = &HFF00&
                    Else
                        
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    ''Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
                    
          Case "ReadUser_AfterWrote"
                    '240E000800DE000900000000000000000386D481110C012530055FFBFFFFDC400300000000000000000000000000000000000000000000000000000000000000
                    ''Form1.txt_User = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    If (gTAG.ACC_RESULT = "DE") Then
                       
                        gTAG.USER = Mid(ByteValueOutput, 17, 112)
                        ''Form1.txt_User = gTAG.USER
                        Me.TagReadResult.FillColor = &HFF00&
                        Debug.Print "Mid(ByteValueOutput, 15, 2)= " + Mid(ByteValueOutput, 15, 2)
                        gTAG.DATA_LENGTH = CInt(Mid(ByteValueOutput, 15, 2))
                        ReadUSB_Report ("ReadUserPart2")
                    Else
                        
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    'Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
          
          Case "ReadUser_AfterWrote64B"
                    ''Form1.txt_User = ""
                    gTAG.ACC_RESULT = Mid(ByteValueOutput, 11, 2)
                    
                    If (Mid(ByteValueOutput, 3, 8) = "07000700") Then
                        'Form1.txt_User = 'Form1.txt_User_New
                        Me.TagReadResult.FillColor = &HFF00&
                        
                    Else
                        gTAG.ACC_RESULT = "FF"
                    End If
                    Debug.Print "Tag User Result=" + gTAG.ACC_RESULT + ",DATA_LENGTH=" + Mid(ByteValueOutput, 15, 2)
                    'Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = gTAG.ACC_RESULT
                    
                    
          Case "ScanNewTag"
                    '0F1C0005000000170100018FDA16D20D0E3400E2003000390701110610D481000000000000000000000000000000000000000000000000000000000000000000
                    'gTAG.TagCOUNT = CInt(Mid(ByteValueOutput, 21, 2))
                    
                    
                    If (Mid(ByteValueOutput, 7, 2) = "05") Then
                        If (CInt(Mid(ByteValueOutput, 21, 2)) > 0) Then
                             gTAG.EPC = Mid(ByteValueOutput, 39, 24)
                             Form1.txt_EPC_Data.Text = gTAG.EPC
                             ''''''''Form1.txt_EPC_Data_New.Text = gTAG.EPC
                             Me.Shape_TagDetect.FillColor = &HFF00& ' green
                         End If
                    Else
                            
                    End If
                    'Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = TagReadResult

          Case Else
                    'Form1.TextIR.Text = 'Form1.TextIR.Text + "< " + ByteValueOutput + vbCrLf
                    Me.lbl_TagAccessResult.Caption = ""
        End Select
    
    '******************************************************************************
    'ResetEvent
    'Sets the event object in the overlapped structure to non-signaled.
    'Requires a handle to the event object.
    'Returns non-zero on success.
    '******************************************************************************
    
    Call ResetEvent(EventObject)
    Call DisplayResultOfAPICall("ResetEvent")

End Sub


Private Sub Iwrite_Click()

End Sub

Private Sub WriteReport(Optional TxStr As String = "")

'Send data to the device.

Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim ReadBuffer() As Byte
Dim SendBuffer() As Byte
Dim temp As String
Dim Cwritlong As Integer
Dim Wcont As Integer
Dim SendTxStr As String
'******************************************************************************
'WriteFile
'Sends a report to the device.
'Returns: success or failure.
'Requires: the handle returned by CreateFile and
'The output report byte length returned by HidP_GetCaps
'******************************************************************************

If MyDeviceDetected = False And DriveContion = False Then
      Debug.Print "device write fail , please check device is plugin ready!"
      Exit Sub
    ElseIf MyDeviceDetected = False And DriveContion = True Then
    MyDeviceDetected = FindTheHid
    End If
    
If MyDeviceDetected = True Then
'The SendBuffer array begins at 0, so subtract 1 from the number of bytes.
ReDim SendBuffer(Capabilities.OutputReportByteLength - 1)

'temp = TextIW.Text
temp = TxStr
temp = Replace(temp, " ", "")
Cwritlong = Len(temp) / 2
If Cwritlong < Capabilities.OutputReportByteLength - 1 Then
    For Wcont = 1 To Capabilities.OutputReportByteLength - 1 - Cwritlong
    temp = temp + "00"
    Next Wcont
End If
''Form1.TextIR.Text = 'Form1.TextIR.Text + "> " + temp + vbCrLf
'The first byte is the Report ID

SendBuffer(0) = 0

'The next bytes are data
On Error GoTo ERROR_Handle

For Count = 0 To Capabilities.OutputReportByteLength - 2
    '從文本框中取出數放到發送中
    SendBuffer(Count + 1) = "&H" & Trim(Mid(temp, Count * 2 + 1, 2))
Next Count

NumberOfBytesWritten = 0


Result = WriteFile _
    (HIDHandle, _
    SendBuffer(0), _
    CLng(Capabilities.OutputReportByteLength), _
    NumberOfBytesWritten, _
    0)
Call DisplayResultOfAPICall("WriteFile")

Debug.Print " OutputReportByteLength = " & Capabilities.OutputReportByteLength
Debug.Print " NumberOfBytesWritten = " & NumberOfBytesWritten
Debug.Print " Report ID: " & SendBuffer(0)
Debug.Print " Report Data:"


For Count = 1 To UBound(SendBuffer)
    Debug.Print Count & " " & Hex$(SendBuffer(Count))
    'SendTxStr = SendTxStr + Hex$(SendBuffer(Count))
Next Count
    'Debug.Print SendTxStr
End If

Exit Sub

ERROR_Handle:
    'MsgBox Err.Number & "Data Error , Please make sure data match hex format 00 ~ FF  !"
    If Err.Number = 13 Then
        lstResults.AddItem "Data Error , Please make sure input data match Hex format 00 ~ FF  !"
        gTAG.DATAFORMAT_ERR = True ' set gTAG.DATAFORMAT_ERR = True for exit write loop
    Else
        lstResults.AddItem "Err Code : " & Err.Number & "  !"
    End If
End Sub

Private Sub M_About_Click()
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
    'Form1.TextIR.Text = ""
    ReadReport
End Sub

Function Base64Encode(Str() As Byte) As String
On Error GoTo over
Dim buf() As Byte, length As Long, mods As Long
Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
mods = (UBound(Str) + 1) Mod 3
length = UBound(Str) + 1 - mods
ReDim buf(length / 3 * 4 + IIf(mods <> 0, 3, 0))
Dim i As Long
For i = 0 To length - 1 Step 3
buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
Next
If mods = 1 Then
buf(length / 3 * 4) = (Str(length) And &HFC) / &H4
buf(length / 3 * 4 + 1) = (Str(length) And &H3) * &H10
buf(length / 3 * 4 + 2) = 64
buf(length / 3 * 4 + 3) = 64
ElseIf mods = 2 Then
buf(length / 3 * 4) = (Str(length) And &HFC) / &H4
buf(length / 3 * 4 + 1) = (Str(length) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
buf(length / 3 * 4 + 2) = (Str(length) And &HF) * &H4
buf(length / 3 * 4 + 3) = 64
End If
For i = 0 To UBound(buf)
Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
Next
over:
End Function

Function Base64Uncode(B64 As String) As Byte()
On Error GoTo over
Dim OutStr() As Byte, i As Long, j As Long
Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)
Dim length As Long, mods As Long
mods = Len(B64) Mod 4
length = Len(B64) - mods
ReDim OutStr(length / 4 * 3 - 1 + Switch(mods = 2, 1, mods = 3, 2))
For i = 1 To length Step 4
Dim buf(3) As Byte
For j = 0 To 3
buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1
Next
OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
Next
If mods = 2 Then
OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &H30) / 16
ElseIf mods = 3 Then
OutStr(length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &H30) / 16
OutStr(length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(B64, length + 3, 1)) - 1) And &H3C) / &H4
End If
Base64Uncode = OutStr
over:
End Function



