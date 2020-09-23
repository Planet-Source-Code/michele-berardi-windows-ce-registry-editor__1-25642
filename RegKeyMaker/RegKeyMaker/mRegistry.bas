Attribute VB_Name = "mRegistry"
Option Explicit

Public Const NO_ERROR = 0                           ' Function returned successfully.
Public Const ERROR_NONE = 0

Public Const HKEY_CURRENT_USER = &H80000001         ' Reference a section of the registry.
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Const KEY_ALL_ACCESS = &H3F                  ' This is a simplified version of the declaration in WINCEAPI.TXT
Public Const REG_DWORD = 4                          ' 32-bit number.
Public Const REG_BINARY = 3 ' redefine this!!!
Public Const REG_OPTION_NON_VOLATILE = 0            ' Key is preserved when system is rebooted.
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_SZ = 1                             ' Unicode nul terminated string.
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const KEY_ENUMERATE_SUB_KEYS = &H8


Public Declare Function RegEnumValue Lib "Coredll" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long


Declare Function RegCloseKey Lib "Coredll" ( _
  ByVal hKey As Long) _
  As Long

Declare Function RegCreateKeyEx Lib "Coredll" Alias "RegCreateKeyExW" ( _
  ByVal hKey As Long, _
  ByVal lpSubKey As String, _
  ByVal Reserved As Long, _
  ByVal lpClass As String, _
  ByVal dwOptions As Long, _
  ByVal samDesired As Long, _
  ByVal lpSecurityAttributes As Long, _
  phkResult As Long, _
  lpdwDisposition As Long) _
  As Long

Declare Function RegDeleteKey Lib "Coredll" Alias "RegDeleteKeyW" ( _
  ByVal hKey As Long, _
  ByVal lpSubKey As String) _
  As Long

'
' New Function added on  31 . 03 . 2001
'

Public Declare Function RegDeleteValue Lib "Coredll" Alias "RegDeleteValueW" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegEnumKeyEx Lib "Coredll" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpname As String, lpcbname As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Long) As Long

'Public Declare Function RegEnumValue Lib "Coredll" Alias "RegEnumValueW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Public Declare Function RegQueryInfoKey Lib "Coredll" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Long) As Long

'
'
'

Declare Function RegOpenKeyEx Lib "Coredll" Alias "RegOpenKeyExW" ( _
  ByVal hKey As Long, _
  ByVal lpSubKey As String, _
  ByVal ulOptions As Long, _
  ByVal samDesired As Long, _
  phkResult As Long) _
  As Long

Declare Function RegQueryValueEx Lib "Coredll" Alias "RegQueryValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
  lpType As Long, _
  ByVal lpData As Long, _
  lpcbData As Long) _
  As Long

Declare Function RegQueryValueExLong Lib "Coredll" Alias "RegQueryValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
  lpType As Long, _
  lpData As Long, _
  lpcbData As Long) As Long

Declare Function RegSetValueExLong Lib "Coredll" Alias "RegSetValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal Reserved As Long, _
  ByVal dwType As Long, _
        lpValue As Long, _
  ByVal cbData As Long) _
  As Long

Declare Function RegQueryValueExString Lib "Coredll" Alias "RegQueryValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
        lpType As Long, _
  ByVal lpData As String, _
        lpcbData As Long) _
  As Long

Declare Function RegQueryValueExMsz Lib "Coredll" Alias "RegQueryValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
        lpType As Long, _
  ByVal lpData As String, _
        lpcbData As Long) _
  As Long

Declare Function RegSetValueExString Lib "Coredll" Alias "RegSetValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal Reserved As Long, _
  ByVal dwType As Long, _
  ByVal lpValue As String, _
  ByVal cbData As Long) _
  As Long
  
  
Declare Function RegQueryValueExBin Lib "Coredll" Alias "RegQueryValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal lpReserved As Long, _
        lpType As Long, _
  ByVal lpData As String, _
        lpcbData As Long) As Long
'
' to be corrected ...
'
Declare Function RegSetValueExBin Lib "Coredll" Alias "RegSetValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal Reserved As Long, _
  ByVal dwType As Long, _
  ByVal lpValue As String, _
  ByVal cbData As Long) _
  As Long
    '
  ' not exactly defined for msz size vars...
  '
  Declare Function RegSetValueExMsz Lib "Coredll" Alias "RegSetValueExW" ( _
  ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal Reserved As Long, _
  ByVal dwType As Long, _
        lpValue As Long, _
  ByVal cbData As Long) _
  As Long
  
  



Public Function CreateNewKey(lSection As Long, _
                             sNewKeyName As String)
       
    Dim hNewKey As Long         '-- Handle to the new key
    Dim lretval As Long         '-- Result of the RegCreateKeyEx function
    
    '-- Create Registry Key
    '-- If key already exists, nothing happens
    lretval = RegCreateKeyEx(lSection, sNewKeyName, CLng(0), _
              vbNullString, REG_OPTION_NON_VOLATILE, _
              KEY_ALL_ACCESS, _
              CLng(0), hNewKey, lretval)
    
    '-- Return Handle to Key
    CreateNewKey = hNewKey
    
    '-- Close Registry Handle
    RegCloseKey (hNewKey)

End Function

Public Function QueryValue(lSection As Long, _
                           sKeyName As String, _
                           sValueName As String)

       Dim lretval  As Long         'result of the API functions
       Dim hKey     As Long         'handle of opened key
       Dim vValue   As Variant      'setting of queried value

       '-- Open Registry Key
       lretval = RegOpenKeyEx(lSection, sKeyName, 0, _
                KEY_ALL_ACCESS, hKey)
       
       '-- Query Registry Value
       lretval = QueryValueEx(hKey, sValueName, vValue)
       
       '-- Return Value
       QueryValue = vValue
       
       '-- Close Reg Key
       RegCloseKey (hKey)

End Function

'SetValueEx and QueryValueEx Wrapper Functions:
Public Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           lType As Long, vValue As Variant) As Long
       
    Dim lValue As Long
    Dim sValue As String
    
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, CLng(0), _
                                           lType, CStr(sValue), Len(sValue) * 2)
        Case REG_DWORD
            'sValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, CLng(0), _
                                           lType, vValue, 4)
        
        Case REG_BINARY
            
'            sValue = vValue
'
' SETTABIN!
'
'

  SetValueEx = RegSetValueExBin(hKey, sValueName, CLng(0), _
                                           lType, vValue, Len(vValue))
                                           
Case REG_MULTI_SZ

            sValue = vValue & Chr(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, CLng(0), _
                                           lType, CStr(sValue), Len(sValue) * 2)

                                             
    End Select
    
End Function

Function QueryValueEx(ByVal lhKey As Long, _
                      ByVal szValueName As String, _
                      vValue As Variant) As Long
       
    Dim cch      As Long
    Dim lrc      As Long
    Dim lType    As Long
    Dim lValue   As Long
    Dim sValue   As String
     
    '-- Determine the size and type of data to be read
    lrc = RegQueryValueEx(lhKey, szValueName, CLng(0), lType, CLng(0), cch)
    
    Select Case lType
        '-- For strings
        Case REG_SZ:
        
        fRegistryTest.txtSubktype.Text = "str"
        
            sValue = String(cch / 2 - 1, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, CLng(0), lType, _
                                        sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        
Case REG_BINARY:
    
    'LEGGIBIN!
    
    fRegistryTest.txtSubktype.Text = "bin"
    
            sValue = String(cch, 0)
            lrc = RegQueryValueExBin(lhKey, szValueName, CLng(0), lType, _
                                        sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = sValue
            Else
                vValue = Empty
            End If
         
        '-- For DWORDS
        Case REG_DWORD:
            
       fRegistryTest.txtSubktype.Text = "dwd"
        
            lrc = RegQueryValueExLong(lhKey, szValueName, CLng(0), lType, _
                                      lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
    
                   
           
'
'
'
'
       Case REG_MULTI_SZ:
       
 fRegistryTest.txtSubktype.Text = "msz"
       
            sValue = String(cch / 2 - 1, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, CLng(0), lType, _
                                        sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left(sValue, cch - 1)
            Else
                vValue = Empty
            End If
       
        Case Else
            '-- All other data types not supported
            lrc = -1
    End Select

End Function
Public Sub SetKeyValue(lSection As Long, _
                       sKeyName As String, _
                       sValueName As String, _
                       vValueSetting As Variant, _
                       lValueType As Long)
   
    Dim lretval  As Long         '-- Result of the SetValueEx function
    Dim hKey     As Long         '-- Handle of open key
       
    '-- Open the specified key
    lretval = RegOpenKeyEx(lSection, sKeyName, _
                           0, _
                           KEY_ALL_ACCESS, hKey)
                              
    '-- Set New Value
    lretval = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    
    '-- Close Reg Key
    RegCloseKey (hKey)
End Sub

 Public Function GetFileDialog(Intestazione As String, filefilter As String, DefExt As String, Operation As String) As String
'Dim fileflags As FileOpenConstants
 fRegistryTest.CommonDialog1.DialogTitle = Intestazione
 fRegistryTest.CommonDialog1.DefaultExt = DefExt
 fRegistryTest.CommonDialog1.InitDir = "\"
  fRegistryTest.CommonDialog1.Filter = filefilter
 fRegistryTest.CommonDialog1.FilterIndex = 0
 
 If Operation = "load" Then
 ' fileflags = cdlOFNFileMustExist
 fRegistryTest.CommonDialog1.Flags = cdlOFNFileMustExist
 fRegistryTest.CommonDialog1.ShowOpen
 
 Else
 
 '
 ' passando un parametro non null
 ' e premendo cancel lui sovreascrive
 ' comunque il <nomefile> passato
 '
 
 fRegistryTest.CommonDialog1.FileName = ""
 ' fileflags = cdlOFNCreatePrompt + cdlOFNOverwritePrompt
 fRegistryTest.CommonDialog1.Flags = cdlOFNCreatePrompt + cdlOFNOverwritePrompt
 fRegistryTest.CommonDialog1.ShowSave
 End If

GetFileDialog = fRegistryTest.CommonDialog1.FileName

On Error Resume Next





End Function






