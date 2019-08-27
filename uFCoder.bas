Attribute VB_Name = "uFCoder"
Public Enum DlCardType
'DLOGIC CARD TYPES

    DL_MIFARE_ULTRALIGHT = &H1
    DL_MIFARE_ULTRALIGHT_EV1_11 = &H2
    DL_MIFARE_ULTRALIGHT_EV1_21 = &H3
    DL_MIFARE_ULTRALIGHT_C = &H4
    DL_NTAG_203 = &H5
    DL_NTAG_210 = &H6
    DL_NTAG_212 = &H7
    DL_NTAG_213 = &H8
    DL_NTAG_215 = &H9
    DL_NTAG_216 = &HA
    DL_MIFARE_MINI = &H20
    DL_MIFARE_CLASSIC_1K = &H21
    DL_MIFARE_CLASSIC_4K = &H22
    DL_MIFARE_PLUS_S_2K = &H23
    DL_MIFARE_PLUS_S_4K = &H24
    DL_MIFARE_PLUS_X_2K = &H25
    DL_MIFARE_PLUS_X_4K = &H26
    DL_MIFARE_DESFIRE = &H27
    DL_MIFARE_DESFIRE_EV1_2K = &H28
    DL_MIFARE_DESFIRE_EV1_4K = &H29
    DL_MIFARE_DESFIRE_EV1_8K = &H2A
    
End Enum

Public Enum ERRORCODES

    DL_OK = &H0
    COMMUNICATION_ERROR = &H1
    CHKSUM_ERROR = &H2
    READING_ERROR = &H3
    WRITING_ERROR = &H4
    BUFFER_OVERFLOW = &H5
    MAX_ADDRESS_EXCEEDED = &H6
    MAX_KEY_INDEX_EXCEEDED = &H7
    NO_CARD = &H8
    COMMAND_NOT_SUPPORTED = &H9
    FORBIDEN_DIRECT_WRITE_IN_SECTOR_TRAILER = &HA
    ADDRESSED_BLOCK_IS_NOT_SECTOR_TRAILER = &HB
    WRONG_ADDRESS_MODE = &HC
    WRONG_ACCESS_BITS_VALUES = &HD
    AUTH_ERROR = &HE
    PARAMETERS_ERROR = &HF
    MAX_SIZE_EXCEEDED = &H10
    UNSUPPORTED_CARD_TYPE = &H11

    COMMUNICATION_BREAK = &H50
    NO_MEMORY_ERROR = &H51
    CAN_NOT_OPEN_READER = &H52
    READER_NOT_SUPPORTED = &H53
    READER_OPENING_ERROR = &H54
    READER_PORT_NOT_OPENED = &H55
    CANT_CLOSE_READER_PORT = &H56

    WRITE_VERIFICATION_ERROR = &H70
    BUFFER_SIZE_EXCEEDED = &H71
    VALUE_BLOCK_INVALID = &H72
    VALUE_BLOCK_ADDR_INVALID = &H73
    VALUE_BLOCK_MANIPULATION_ERROR = &H74
    WRONG_UI_MODE = &H75
    KEYS_LOCKED = &H76
    KEYS_UNLOCKED = &H77
    WRONG_PASSWORD = &H78
    CAN_NOT_LOCK_DEVICE = &H79
    CAN_NOT_UNLOCK_DEVICE = &H7A
    DEVICE_EEPROM_BUSY = &H7B
    RTC_SET_ERROR = &H7C
    ANTICOLLISION_DISABLED = &H7D
    NO_CARDS_ENUMERRATED = &H7E
    CARD_ALREADY_SELECTED = &H7F

    FT_STATUS_ERROR_1 = &HA0
    FT_STATUS_ERROR_2 = &HA1
    FT_STATUS_ERROR_3 = &HA2
    FT_STATUS_ERROR_4 = &HA3
    FT_STATUS_ERROR_5 = &HA4
    FT_STATUS_ERROR_6 = &HA5
    FT_STATUS_ERROR_7 = &HA6
    FT_STATUS_ERROR_8 = &HA7
    FT_STATUS_ERROR_9 = &HA8

End Enum

Declare Function ReaderOpen Lib "ufr-lib/windows/x86/uFCoder-x86.dll" () As Integer

Declare Function ReaderOpenEx Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByVal readerType As Integer, ByVal portName As String, ByVal portInterface As Integer, ByVal additionalArgument As String) As Integer

Declare Function ReaderUISignal Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByVal sound As Byte, ByVal light As Byte) As Integer

Declare Function GetReaderSerialNumber Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef serial_number As Long) As Integer

Declare Function GetReaderType Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef reader_type As Long) As Integer


Declare Function GetCardIdEx Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef sak As Byte, ByRef uid As Byte, ByRef uid_size As Byte) As Integer

Declare Function LinearFormatCard Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef new_key_A As Byte, ByVal blocks_access_bits As Byte, ByVal sector_trailer_access_bits As Byte, ByVal sector_trailers_byte9 As Byte, ByRef new_key_B As Byte, ByRef sectors_formatted As Byte, ByVal auth_mode As Byte, ByVal key_index As Byte) As Integer

Declare Function ReaderKeyWrite Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef reader_key As Byte, ByVal key_index As Byte) As Integer

Declare Function LinearRead Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef data As Byte, ByVal linear_address As Integer, ByVal length As Integer, ByRef bytes_returned As Integer, ByVal auth_mode As Byte, ByVal key_index As Byte) As Integer

Declare Function LinearWrite Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef data As Byte, ByVal linear_address As Integer, ByVal length As Integer, ByRef bytes_written As Integer, ByVal auth_mode As Byte, ByVal key_index As Byte) As Integer

Declare Function BlockRead Lib "ufr-lib/windows/x86/uFCoder-x86.dll" (ByRef data As Byte, ByVal block_address As Byte, ByVal auth_mode As Byte, ByVal key_index As Byte) As Integer












