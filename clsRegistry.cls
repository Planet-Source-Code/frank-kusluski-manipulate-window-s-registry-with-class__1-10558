VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare private variables
Private m_lngKeyValue As Long
Private m_strKeyValue As String
Private m_strKeyHandle As Long

'Declare enumerations
Public Enum RegDataType
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum KeyTypes
    REG_SZ = 1
    REG_DWORD = 4
End Enum

'Declare private constants
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

'Declare registry APIs
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) _
   As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As _
   Long, phkResult As Long, lpdwDisposition As Long) As Long
   
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult _
   As Long) As Long
   
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal Reserved As Long, ByVal dwType As Long, ByVal _
   lpValue As String, ByVal cbData As Long) As Long
   
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal Reserved As Long, ByVal dwType As Long, lpValue _
   As Long, ByVal cbData As Long) As Long
   
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
   "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
   "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpKeyName As String) As Long
   
Public Property Get KeyHandle() As Long
    KeyHandle = m_strKeyHandle
End Property

Public Property Get StringKeyValue() As String
    StringKeyValue = m_strKeyValue
End Property

Public Property Get LongKeyValue() As Long
    LongKeyValue = m_lngKeyValue
End Property

Public Function CreateKey(ByVal strSectionName As String, ByVal strKey As _
   String, Optional ByVal lngRegDataType As RegDataType = _
   HKEY_LOCAL_MACHINE) As Long
   
   'Create a new key
   CreateKey = CreateKeyEx(strSectionName & strKey, lngRegDataType)
End Function
Private Function CreateKeyEx(strSectionName As String, Optional _
   lngRegDataType As Long = HKEY_LOCAL_MACHINE) As Long
   
   'Declare variables
   Dim lngHwd As Long
   Dim lngRC As Long
   
   'Create the key and close it
   lngRC = RegCreateKeyEx(lngRegDataType, strSectionName, 0&, _
      vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngHwd, _
      lngRC)
   RegCloseKey (lngHwd)
   
   'Return with the return code
   CreateKeyEx = lngRC
End Function

Public Function SetKeyValue(ByVal strSectionName As String, ByVal _
   strKeyName As String, ByVal varKeyData As Variant, ByVal lngKeyType _
   As KeyTypes, Optional ByVal lngRegDataType As RegDataType = _
   HKEY_LOCAL_MACHINE) As Long
   
   'Declare variables
   Dim lngRC As Long
   Dim lngHwd As Long

   'Open the key, set the value and close the key
   lngRC = RegOpenKeyEx(lngRegDataType, strSectionName, 0, _
      KEY_ALL_ACCESS, lngHwd)
   lngRC = SetValueEx(lngHwd, strKeyName, lngKeyType, varKeyData)
   RegCloseKey (lngHwd)
   
   'Return with the return code
   SetKeyValue = lngRC
End Function

Private Function SetValueEx(ByVal lngHwdKey As Long, ByVal strValueName _
   As String, ByVal lngType As Long, ByVal varValueData As Variant) _
   As Long
   
   'Declare variables
   Dim lngValue As Long
   Dim strValue As String
   
   'Determine the size and type of data to be written
   Select Case lngType
   
      'Strings
      Case REG_SZ
         strValue = varValueData & Chr$(0)
         SetValueEx = RegSetValueExString(lngHwdKey, strValueName, _
            0&, lngType, strValue, Len(strValue))
      
      'DWORDs
      Case REG_DWORD
         lngValue = varValueData
         SetValueEx = RegSetValueExLong(lngHwdKey, strValueName, _
            0&, lngType, lngValue, 4)
            
   End Select
End Function

Public Function QueryKeyValue(ByVal strSectionName As String, ByVal _
   strKeyName As String, Optional ByVal lngRegDataType As RegDataType _
   = HKEY_LOCAL_MACHINE) As Long
   
   'Declare variables
   Dim lngRC As Long
   Dim lngHwd As Long
   Dim varValue As Variant
   
   'Get the value currently in the key and close it
   lngRC = RegOpenKeyEx(lngRegDataType, strSectionName, 0, _
      KEY_ALL_ACCESS, lngHwd)
      
   m_strKeyHandle = lngHwd  'set key handle
   
   lngRC = QueryValueEx(lngHwd, strKeyName, varValue)
   RegCloseKey (lngHwd)
   
   'Return with the return code
   QueryKeyValue = lngRC
End Function

Private Function QueryValueEx(ByVal lngHwdKey As Long, ByVal strKeyName _
   As String, ByVal varValue As Variant) As Long
   
   'Declare variables
   Dim lngDataLen As Long
   Dim lngRC As Long
   Dim lngType As Long
   Dim lngTemp As Long
   Dim strTemp As String

   On Error GoTo QueryErr

   'Determine the size and type of data to be read
   lngRC = RegQueryValueExNULL(lngHwdKey, strKeyName, 0&, lngType, _
      0&, lngDataLen)
   If lngRC <> 0 Then Error 5

   Select Case lngType
   
      'Strings
      Case REG_SZ:
         strTemp = String(lngDataLen, 0)
         lngRC = RegQueryValueExString(lngHwdKey, strKeyName, 0&, _
            lngType, strTemp, lngDataLen)
         If lngRC = 0 Then
            m_strKeyValue = Left$(strTemp, lngDataLen - 1)
         Else
            m_strKeyValue = Empty
         End If
         
      'DWORDs
      Case REG_DWORD:
         lngRC = RegQueryValueExLong(lngHwdKey, strKeyName, 0&, _
            lngType, lngTemp, lngDataLen)
         If lngRC = 0 Then
            m_lngKeyValue = lngTemp
         Else
            m_lngKeyValue = -1
         End If
         
      'All other data types
      Case Else
         lngRC = -1
         
   End Select

   On Error GoTo 0
   QueryValueEx = lngRC
   Exit Function
   
QueryErr:
   lngRC = -1
   QueryValueEx = lngRC
End Function

Public Function DeleteKey(lngRegDataType As Long, strKey As String) As Long
Dim lngValue As Long
lngValue = RegDeleteKey(lngRegDataType, strKey)
DeleteKey = lngValue
End Function

Public Function DeleteKeyValue(ByVal strSectionName As String, ByVal _
   strKeyName As String, Optional ByVal lngRegDataType As RegDataType _
   = HKEY_LOCAL_MACHINE) As Long
   
   'Declare variables
   Dim lngRC As Long
   Dim lngHwd As Long
   Dim varValue As Variant
   
   'Get the value currently in the key and close it
   lngRC = RegOpenKeyEx(lngRegDataType, strSectionName, 0, _
      KEY_ALL_ACCESS, lngHwd)
      
   lngRC = RegDeleteValue(lngHwd, strKeyName)
   
   RegCloseKey (lngHwd)
   
   'Return with the return code
   DeleteKeyValue = lngRC
End Function

Private Sub Class_Initialize()
'MsgBox "class init"
End Sub

Private Sub Class_Terminate()
'MsgBox "class terminate"
End Sub
