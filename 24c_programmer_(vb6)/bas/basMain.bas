Attribute VB_Name = "basMain"
Option Explicit

Public bByte        As Integer      '\\ contain selected data
Public sHex         As String       '\\ contain hex data but like string
Public bIsChanged   As Boolean      '\\ to check if data is changed
Public sLine        As String       '\\ contain selected line of data
Public bDirty       As Boolean      '\\ check if data is changed
Public i2cDelay     As Integer      '\\
Public iStartDelay  As Integer      '\\
Public iVerAfter    As Integer      '\\
Public iVerDuring   As Integer      '\\
Public iComPort     As Integer      '\\
'Public iAutoClose   As Integer      '\\ auto close program window when finished
Public Enum EepromOperation
  Read_Memory = 0
  Write_Memory = 1
  Blank_Check = 2
  Verify_Data = 3
End Enum

Public EepromAction As EepromOperation

Public iMemory      As Integer      '\\ memory type



