Attribute VB_Name = "modArrayPointer"
Option Explicit

'I got this stuff as i was asking for it on an forum

Public Type SAFEARRAYBOUND ' 8 bytes
  cElements As Long
  lLbound   As Long
End Type

Public Type SAFEARRAYHEADER ' 20 bytes (for one dimensional arrays
  dimensions    As Integer
  fFeatures     As Integer
  DataSize      As Long
  cLocks        As Long
  dataPointer   As Long
  sab(1)        As SAFEARRAYBOUND
End Type

Const FADF_AUTO = &H1&        '// Array is allocated on the stack.
Const FADF_STATIC = &H2&      '// Array is statically allocated.
Const FADF_EMBEDDED = &H4&    '// Array is embedded in a structure.
Const FADF_FIXEDSIZE = &H10&  '// Array may not be resized or reallocated.
Const FADF_BSTR = &H100&      '// An array of BSTRs.
Const FADF_UNKNOWN = &H200&   '// An array of IUnknown*.
Const FADF_DISPATCH = &H400&  '// An array of IDispatch*.
Const FADF_VARIANT = &H800&   '// An array of VARIANTs.
Const FADF_RESERVED = &HF0E8& '// Bits reserved for future use.

Public Enum eDATASIZE
  byteArray = 1
  integerArray = 2  ' or Boolean Data Type
  longArray = 4
  singleArray = 4
  doubleArray = 8
End Enum

'Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Var() As Any) As Long
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function RedimArray(ByVal DataSize As Long, ByVal lNumElements As Long, ByRef sa As SAFEARRAYHEADER, ByVal lDataPointer As Long, ByVal lArrayPointer As Long, Optional LoBound As Long = 0) As Long
  If lNumElements > 0 And lDataPointer <> 0 And lArrayPointer <> 0 Then
    With sa
      .DataSize = DataSize                              ' byte = 1 byte, integer = 2 bytes etc
      .dimensions = 1 '2                                ' one dimensional
      .dataPointer = lDataPointer                       ' to unicode string data (or other?)
      .sab(0).lLbound = LoBound                         ' lower bound
      .sab(0).cElements = lNumElements                  ' number of elements
      '.sab(1).cElements = lNumElements
      '.sab(1).lLbound = LoBound
      CopyMemory ByVal lArrayPointer, VarPtr(sa), 4& ' fake VB out
      RedimArray = True
    End With
  End If
End Function

Public Sub DestroyArray(ByVal lArrayPointer As Long)
  Dim lZero As Long
 
  DXCopyMemory ByVal lArrayPointer, lZero, 4         ' put the array back to its original state
End Sub
