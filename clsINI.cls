VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Function            : INI Manipulation
' Original by         : Eric Dalquist
' Last Updated by     : enZo c",)
' Last Updated        : November 14, 2002

Option Explicit

'The two API calls to read and write from the INI file
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

' API declaration, returns the default windows path.
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'the INI file's path
Private strINIFile As String
'catching the API's result
Private lngkeyLen As Long
'holds the buffer size intially it holds 255
Private intBufferSize As Integer

Private Sub Class_Initialize()
    'when the object is created we set the default file name
    ResetINIFile
End Sub

'get windows directory
Public Function GetWindowsDir() As String
Dim RetStr As String, RetLen As Long
    RetStr = Space(1024)
    RetLen = GetWindowsDirectory(RetStr, Len(RetStr))
    RetStr = Left(RetStr, RetLen)
    ' Add a "\" at the end of the windows directory if not present.
    If Not Right(RetStr, 1) = "\" Then RetStr = RetStr & "\"
    GetWindowsDir = RetStr
End Function

'****************************** Start Let Properties ******************************
Public Property Let INIFile(ByVal INIFileIn As String)
    'set the INI path
    strINIFile = INIFileIn
End Property

Public Property Let BufferSize(ByVal intSize As Integer)
    intBufferSize = intSize
End Property
'****************************** End Let Properties ******************************

'****************************** Start Get Properties ******************************
Public Property Get INIFile() As String
    'return the INI path
    INIFile = strINIFile
End Property

Public Property Get BufferSize() As Integer
    BufferSize = intBufferSize
End Property
'****************************** End Get Properties ******************************

Public Sub ResetINIFile()
    strINIFile = GetWindowsDir & "\twwINImanipulator.ini"
    intBufferSize = 1024
End Sub

'****************************** Start Create Methods ******************************
Public Sub CreateKey(ByVal strSection As String, ByVal strKey As String)
    'call the write profile string but since we only want to create a key but
    'not give it a value we pass "" (vbNullString) for the value
  lngkeyLen = WritePrivateProfileString(strSection, strKey, "", strINIFile)
End Sub

Public Sub CreateKeyValue(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)
    'call the write profile string with all the parameters
    lngkeyLen = WritePrivateProfileString(strSection, strKey, strValue, strINIFile)
End Sub
'****************************** End Create Methods ******************************


'****************************** Start Get Methods ******************************
Public Function GetKeyValue(ByVal Section As String, ByVal Key As String) As String
    Dim RetVal As String
    'It just creates a
    'string of that length in memory for the API call to write the results of it's action
    RetVal = String(intBufferSize, 0)
    lngkeyLen = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), strINIFile)
    
    'return a null string if the return has 0 length
    If lngkeyLen = 0 Then
        GetKeyValue = ""
    Else ' or the key's value
        GetKeyValue = Left(RetVal, lngkeyLen)
    End If
End Function

Public Function GetKeysInSection(ByVal Section As String) As String
    Dim RetVal As String

    RetVal = String(intBufferSize, 0)
    lngkeyLen = GetPrivateProfileString(Section, vbNullString, "", RetVal, Len(RetVal), strINIFile)
    
    'if no keys are found then return "" (vbNullString)
    If lngkeyLen = 0 Then
        GetKeysInSection = ""
    Else
        'if the retval is > 0 then return the results
        'since we are getting multiple keys but returning them as one string the
        'programer should use the split() function in the returned value with
        'chr(0) being the delimiter
        GetKeysInSection = Left(RetVal, lngkeyLen - 1)
    End If
End Function

Public Function GetSections() As String
    Dim RetVal As String
    
    RetVal = String(intBufferSize, 0)
    lngkeyLen = GetPrivateProfileString(vbNullString, vbNullString, "", RetVal, Len(RetVal), strINIFile)
    
    'if no sections are found then return "" (vbNullString)
    If lngkeyLen = 0 Then
        GetSections = ""
    Else
        'if the retval is > 0 then return the results
        'since we are getting multiple sections but returning them as one string the
        'programer should use the split() function in the returned value with
        'chr(0) being the delimiter
        GetSections = Left(RetVal, lngkeyLen - 1)
    End If
End Function
'****************************** End Get Methods ******************************

'****************************** Start Delete Methods ******************************
Public Sub DeleteKeyValue(ByVal Section As String, ByVal Key As String)
    'to delete the key's value just write a null string to it
    lngkeyLen = WritePrivateProfileString(Section, Key, "", strINIFile)
End Sub

Public Sub DeleteKey(ByVal Section As String, ByVal Key As String)
    'to delete a key use the actual vbnullstring object ... for some reason
    'it behaves differently that passing the API ""
    lngkeyLen = WritePrivateProfileString(Section, Key, vbNullString, strINIFile)
End Sub

Public Sub DeleteSection(ByVal Section As String)
    'to delete a section it's the same thing as a key but we write a vbNullString
    'to the key name and value
    lngkeyLen = WritePrivateProfileString(Section, vbNullString, vbNullString, strINIFile)
End Sub
'****************************** End Delete Methods ******************************

'****************************** Start Rename Methods ******************************
Public Sub RenameKey(ByVal Section As String, ByVal CurrentKey As String, ByVal NewKey As String)
    
    CreateKeyValue Section, NewKey, GetKeyValue(Section, CurrentKey)
    'delete the old key
    DeleteKey Section, CurrentKey
    
End Sub

Public Sub RenameSection(ByVal CurrentSection As String, ByVal NewSection As String)
    Dim KeyList() As String
    Dim KeyCount As Integer
    Dim index As Integer
            
    'we get all the key names into an array
    KeyList = Split(GetKeysInSection(CurrentSection), Chr(0))
    'get the number of keys
    KeyCount = UBound(KeyList)
    
    'for each key in the array
    For index = 0 To KeyCount
        'create the new key in the new section
        CreateKeyValue NewSection, KeyList(index), GetKeyValue(CurrentSection, KeyList(index))
    Next index
    
    'delete the old section
    DeleteSection CurrentSection
End Sub
'****************************** End Rename Methods ******************************

