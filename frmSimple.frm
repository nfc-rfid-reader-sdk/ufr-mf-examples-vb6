VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   Caption         =   "uFR Simple"
   ClientHeight    =   10455
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   8235
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmLinearWrite 
      Caption         =   "Linear Write"
      Height          =   2535
      Left            =   120
      TabIndex        =   60
      Top             =   7320
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox txtLinearWriteData 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   480
         Width           =   7455
      End
      Begin VB.TextBox txtLinearWriteAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   64
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtLinearWriteLength 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   63
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtLinearWriteBytesWritten 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   62
         Top             =   2000
         Width           =   495
      End
      Begin VB.CommandButton btnLinearWrite 
         Caption         =   "LINEAR WRITE"
         Height          =   495
         Left            =   4080
         TabIndex        =   61
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label Label19 
         Caption         =   "Write Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Linear address:"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Data length:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   " Bytes written:"
         Height          =   255
         Left            =   1995
         TabIndex        =   66
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFunctionStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   52
      Top             =   10000
      Width           =   5535
   End
   Begin VB.Frame frmLinearRead 
      Caption         =   "Linear Read"
      Height          =   2535
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Width           =   7935
      Begin VB.CommandButton btnLinearRead 
         Caption         =   "LINEAR READ"
         Height          =   495
         Left            =   4080
         TabIndex        =   59
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtLinearReadReadBytes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   58
         Top             =   2000
         Width           =   495
      End
      Begin VB.TextBox txtLinearReadLength 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   56
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtLinearReadAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   55
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtLinearReadData 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label Label15 
         Caption         =   " Read bytes:"
         Height          =   255
         Left            =   2000
         TabIndex        =   57
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Data length:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Linear address:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Read Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox cmbLinearFunc 
      Height          =   315
      ItemData        =   "frmSimple.frx":0000
      Left            =   1920
      List            =   "frmSimple.frx":000A
      TabIndex        =   47
      Text            =   "Linear Read"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.ComboBox cmbCardReader 
      Height          =   315
      ItemData        =   "frmSimple.frx":0029
      Left            =   1800
      List            =   "frmSimple.frx":0033
      TabIndex        =   33
      Text            =   "New card keys"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Frame frmNewReaderKey 
      Caption         =   "New  reader key"
      Height          =   1335
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton btnStoreKeyIntoReader 
         Caption         =   "WRITE KEY INTO READER"
         Height          =   850
         Left            =   2520
         TabIndex        =   45
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtReaderKeyIndex 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   43
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtReaderKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   42
         Text            =   "FFFFFFFFFFFF"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblKeyIndex 
         Caption         =   "Key index:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Key:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   240
      TabIndex        =   27
      Top             =   4440
      Width           =   7935
      Begin VB.OptionButton optAuth1B 
         Caption         =   "AUTH 1B"
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAuth1A 
         Caption         =   "AUTH 1A"
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5760
      Top             =   240
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   8055
      Begin VB.CommandButton btnReaderUiSignal 
         Caption         =   "READER UI SIGNAL"
         Height          =   735
         Left            =   3360
         TabIndex        =   26
         Top             =   1550
         Width           =   4455
      End
      Begin VB.ComboBox cmbSoundMode 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSimple.frx":0056
         Left            =   1320
         List            =   "frmSimple.frx":006C
         TabIndex        =   25
         Text            =   "None"
         Top             =   2000
         Width           =   1695
      End
      Begin VB.ComboBox cmbLightMode 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSimple.frx":00AE
         Left            =   1320
         List            =   "frmSimple.frx":00C1
         TabIndex        =   24
         Text            =   "None"
         Top             =   1520
         Width           =   1695
      End
      Begin VB.TextBox txtCardUID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Top             =   690
         Width           =   3135
      End
      Begin VB.TextBox txtUIDSize 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   20
         Top             =   350
         Width           =   975
      End
      Begin VB.TextBox txtCardType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   350
         Width           =   855
      End
      Begin VB.TextBox txtRSerial 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtRType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   350
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Sound mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2000
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Light mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1520
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Top             =   1080
         Width           =   8040
      End
      Begin VB.Label Label4 
         Caption         =   "UID size:"
         Height          =   255
         Left            =   5880
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Card seral:"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   700
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Card type:"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Reader Serial:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   700
         Width           =   1000
      End
      Begin VB.Label txtReaderType 
         Caption         =   "ReaderType:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frmAdvancedOptions 
      Caption         =   "Advanced options"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   8055
      Begin VB.TextBox txtOpenArg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   10
         Top             =   350
         Width           =   1695
      End
      Begin VB.TextBox txtPortInterface 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   8
         Top             =   350
         Width           =   255
      End
      Begin VB.TextBox txtPortName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   350
         Width           =   1455
      End
      Begin VB.TextBox txtReaderTypeEx 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblOpenArg 
         Caption         =   "Arg:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblPortInterface 
         Caption         =   "Port interface:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPortName 
         Caption         =   "Port name:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblReaderTypeEx 
         Caption         =   "Reader type:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CheckBox checkAdvanced 
      Appearance      =   0  'Flat
      Caption         =   "Use Advanced options"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton btnReaderOpen 
      Appearance      =   0  'Flat
      Caption         =   "Reader Open"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame frmNewCardKeys 
      Caption         =   "New card keys"
      Height          =   1335
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   7935
      Begin VB.TextBox txtFormatCardSectorsFormatted 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   40
         Top             =   540
         Width           =   375
      End
      Begin VB.CommandButton btnFormatCard 
         Caption         =   "FORMAT CARD"
         Height          =   850
         Left            =   2520
         TabIndex        =   38
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtFormatCardKeyB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   37
         Text            =   "FFFFFFFFFFFF"
         Top             =   800
         Width           =   1335
      End
      Begin VB.TextBox txtFormatCardKeyA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   36
         Text            =   "FFFFFFFFFFFF"
         Top             =   350
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Sector formatted:"
         Height          =   285
         Left            =   5520
         TabIndex        =   39
         Top             =   550
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Key B:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Key A:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label12 
      Caption         =   "FUNCTION STATUS:"
      DataField       =   "STATUS:"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   10080
      Width           =   1935
   End
   Begin VB.Label lblLinearFunc 
      Caption         =   "Linear Read"
      Height          =   195
      Left            =   240
      TabIndex        =   46
      Top             =   7020
      Width           =   1455
   End
   Begin VB.Label lblKeysMode 
      Caption         =   "New card keys"
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   5160
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cmbLightMode.ListIndex = 0
    cmbSoundMode.ListIndex = 0
End Sub

Private Function ErrorCode(ByVal code As Integer) As String
    Dim code_str As String
    code_str = ""
    
    If code < 15 Then
        code_str = "&H0" + CStr(Hex$(code))
    Else
        code_str = "&H" + CStr(Hex$(code))
    End If
    
    ErrorCode = code_str

End Function

Private Function HexToBytes(ByVal HexString As String) As Byte()

    Dim Bytes() As Byte
    Dim HexPos As Integer
    Dim HexDigit As Integer
    Dim BytePos As Integer
    Dim Digits As Integer

    ReDim Bytes(Len(HexString) \ 2)
    For HexPos = 1 To Len(HexString)
        HexDigit = InStr("0123456789ABCDEF", _
                         UCase$(Mid$(HexString, HexPos, 1))) - 1
        If HexDigit >= 0 Then
            If BytePos > UBound(Bytes) Then
                ReDim Preserve Bytes(UBound(Bytes) + 4)
            End If
            Bytes(BytePos) = Bytes(BytePos) * &H10 + HexDigit
            Digits = Digits + 1
        End If
        If Digits = 2 Or HexDigit < 0 Then
            If Digits > 0 Then BytePos = BytePos + 1
            Digits = 0
        End If
    Next
    If Digits = 0 Then BytePos = BytePos - 1
    If BytePos < 0 Then
    Else
        ReDim Preserve Bytes(BytePos)
    End If
    HexToBytes = Bytes
End Function

Function ByteArrayToHexStr(b() As Byte) As String
   Dim n As Long, i As Long
   
   ByteArrayToHexStr = Space$(3 * (UBound(b) - LBound(b)) + 2)
   n = 1
   For i = LBound(b) To UBound(b)
      Mid$(ByteArrayToHexStr, n, 2) = Right$("00" & Hex$(b(i)), 2)
      n = n + 3
   Next
End Function

Public Function HexByte2Char(ByVal Value As Byte) As String
  HexByte2Char = IIf(Value < &H10, "0", "") & Hex$(Value)
End Function

Private Sub btnReaderOpen_Click()

    Dim status As Integer
    
    Dim reader_sn As Long
    
    Dim r_type As Long

    If checkAdvanced.Value = 1 Then
        Dim reader_type_str As String
        Dim port_name As String
        Dim port_interface_str As String
        Dim arg As String
        Dim port_interface As Integer
        Dim reader_type As Integer
        
        reader_type_str = txtReaderTypeEx.Text
        port_name = txtPortName.Text
        port_interface_str = txtPortInterface.Text
        arg = txtOpenArg.Text
        
        
        reader_type = Val(reader_type_str)
        port_interface = Asc(port_interface_str)
        
        status = ReaderOpenEx(reader_type, port_name, port_interface, arg)
    Else
        status = ReaderOpen()
    End If
    
    
    If status = 0 Then
        status = ReaderUISignal(1, 1)
        status = GetReaderType(r_type)
        status = GetReaderSerialNumber(reader_sn)
        
        txtRType.Text = CStr(Hex$(r_type))
        txtRSerial.Text = "UN" + CStr(reader_sn)
        txtFunctionStatus = "Reader opened !"
        Timer1.Enabled = True
    Else
        txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
        Timer1.Enabled = False
        txtRType.Text = ""
        txtRSerial.Text = ""
    End If
    
End Sub

Private Sub Timer1_Timer()

    Dim status As Integer
    Dim sak As Byte
    Dim uid_size As Byte
    Dim uid(11) As Byte
    txtCardUID.Text = ""
    txtCardType = ""
    txtUIDSize.Text = ""
    Dim str(1) As String
    
    status = GetCardIdEx(sak, uid(LBound(uid)), uid_size)
  
    If status = 0 Then
        Dim short_uid() As Byte
        ReDim short_uid(uid_size - 1) As Byte
        For i = 1 To uid_size
        short_uid(i - 1) = uid(i - 1)
        Next i
        
        Dim uid_str As String
        
        uid_str = ByteArrayToHexStr(short_uid)
        txtCardUID.Text = uid_str
        txtCardType = HexByte2Char(sak)
        txtUIDSize.Text = HexByte2Char(uid_size)
        
    ElseIf status = 8 Then
        txtCardUID.Text = "NO CARD"
        txtCardType = ""
        txtUIDSize.Text = ""
    Else
        txtCardUID.Text = ""
        txtCardType = ""
        txtUIDSize.Text = ""
        txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
    End If

End Sub

Private Sub btnReaderUiSignal_Click()
    
    Dim light_mode As Byte
    Dim sound_mode As Byte
    Dim status As Integer
    
    light_mode = cmbLightMode.ListIndex
    sound_mode = cmbSoundMode.ListIndex
    
    txtPortName.Text = light_mode
    txtReaderTypeEx.Text = sound_mode
    status = ReaderUISignal(light_mode, sound_mode)

    If status = 0 Then
        txtFunctionStatus.Text = "UFR_OK"
    Else
        txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
    End If
End Sub


Private Sub btnFormatCard_Click()
    Dim keyA_str As String
    Dim keyB_str As String
    
    Dim status As Integer
    Dim sectors_formatted As Byte
    Dim auth_mode As Byte
    
    sectors_formatted = 0
    
    If optAuth1A.Value = True Then
        auth_mode = &H60
    Else
        auth_mode = &H61
    End If
    
    keyA_str = txtFormatCardKeyA.Text
    keyB_str = txtFormatCardKeyB.Text
        
    Dim keyA() As Byte
    Dim keyB() As Byte
    
    keyA = HexToBytes(keyA_str)
    keyB = HexToBytes(keyB_str)
    
    status = LinearFormatCard(keyA(0), 0, 1, &H69, keyB(0), sectors_formatted, auth_mode, 0)
    
    If status = 0 Then
        txtFunctionStatus = "Card formatted !"
        txtFormatCardSectorsFormatted.Text = CStr(sectors_formatted)
        
    Else
        txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
        txtFormatCardSectorsFormatted.Text = ""
    End If
    
End Sub

Private Sub btnStoreKeyIntoReader_Click()

    Dim status As Integer
    Dim key_str As String
    Dim key() As Byte
    Dim key_index As Byte
    
    key_str = txtReaderKey.Text
    key_index = Val(txtReaderKeyIndex.Text)
    key = HexToBytes(key_str)
    
    status = ReaderKeyWrite(key(0), 0)
    
    If status = 0 Then
        txtFunctionStatus = "Key stored !"
        
    Else
        txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
    End If
   
End Sub

Private Sub btnLinearRead_Click()
    
    Dim status As Integer
    Dim address As Integer
    Dim length As Integer
    Dim returned As Integer
    
    address = 0
    length = 0
    
    If Not txtLinearReadAddress.Text = "" Then
        address = Val(txtLinearReadAddress.Text)
    Else
        MsgBox ("Linear Read requires start address !")
        txtLinearReadAddress.SetFocus
        Exit Sub
    End If
    
    If Not txtLinearReadLength.Text = "" Then
        length = Val(txtLinearReadLength.Text)
    Else
        MsgBox ("Linear Read requires data length to read !")
        txtLinearReadLength.SetFocus
        Exit Sub
    End If
    
    
    Dim data() As Byte
    ReDim data(length - 1) As Byte
    
    If optAuth1A.Value = True Then
        auth_mode = &H60
    Else
        auth_mode = &H61
    End If
    
     status = LinearRead(data(0), address, length, returned, auth_mode, 0)
     
    If status = 0 Then
        txtFunctionStatus.Text = "Linear read done ! "
        txtLinearReadData.Text = ByteArrayToHexStr(data)
        txtLinearReadReadBytes = CStr(returned)
    Else
        txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
    End If
    
End Sub

Private Sub btnLinearWrite_Click()
Dim status As Integer
    
    Dim address As Integer
    Dim length As Integer
    Dim written As Integer
    
    address = 0
    
    length = Val(txtLinearWriteLength.Text)
    
    If txtLinearWriteData.Text = "" Then
        MsgBox ("Linear write requires data to write !")
        txtLinearWriteData.SetFocus
        Exit Sub
    End If
    
    Dim data_str As String
    Dim data() As Byte
    
    data_str = txtLinearWriteData.Text
    
    data = HexToBytes(data_str)
    
    If Not txtLinearWriteAddress.Text = "" Then
        address = Val(txtLinearWriteAddress.Text)
    Else
        MsgBox ("Linear Write requires start address !")
        txtLinearWriteAddress.SetFocus
        Exit Sub
    End If
    
    If optAuth1A.Value = True Then
        auth_mode = &H60
    Else
        auth_mode = &H61
    End If
    
     status = LinearWrite(data(0), address, length, written, auth_mode, 0)
     
    If status = 0 Then
       txtFunctionStatus.Text = "Linear write done ! "
       txtLinearWriteBytesWritten.Text = CStr(written)
    Else
         txtFunctionStatus.Text = "Error code: " + ErrorCode(status)
    End If

End Sub

Private Sub checkAdvanced_Click()

    If checkAdvanced.Value = 1 Then
        lblReaderTypeEx.Enabled = True
        txtReaderTypeEx.Enabled = True
        lblPortName.Enabled = True
        txtPortName.Enabled = True
        lblPortInterface.Enabled = True
        txtPortInterface.Enabled = True
        lblOpenArg.Enabled = True
        txtOpenArg.Enabled = True
        frmAdvancedOptions.Enabled = True
    Else
        lblReaderTypeEx.Enabled = False
        txtReaderTypeEx.Enabled = False
        lblPortName.Enabled = False
        txtPortName.Enabled = False
        lblPortInterface.Enabled = False
        txtPortInterface.Enabled = False
        lblOpenArg.Enabled = False
        txtOpenArg.Enabled = False
        frmAdvancedOptions.Enabled = False
    End If

End Sub

Private Sub cmbCardReader_Click()
    If cmbCardReader.List(cmbCardReader.ListIndex) = "New card keys" Then
        frmNewCardKeys.Visible = True
        frmNewReaderKey.Visible = False
        lblKeysMode.Caption = "New card keys"
    Else
        frmNewCardKeys.Visible = False
        frmNewReaderKey.Visible = True
        lblKeysMode.Caption = "New reader key"
    End If
End Sub

Private Sub cmbLinearFunc_Click()
    If cmbLinearFunc.List(cmbLinearFunc.ListIndex) = "Linear Read" Then
        frmLinearRead.Visible = True
        frmLinearWrite.Visible = False
        lblLinearFunc.Caption = "Linear Read"
        
    Else
        lblLinearFunc.Caption = "Linear Write"
        frmLinearRead.Visible = False
        frmLinearWrite.Visible = True
    End If
End Sub

Private Sub txtLinearWriteData_Change()
    txtLinearWriteLength.Text = CStr(Round(Len(txtLinearWriteData.Text) / 2))
    End Sub
