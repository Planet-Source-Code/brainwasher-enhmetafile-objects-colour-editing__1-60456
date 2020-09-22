Attribute VB_Name = "Md_Metafile"
Option Explicit
'
'___________________________________________________________________________
' Program name      : EdCol_enhMetafile.
' Description       : A simple way to edit the object's colours in an enhanced
'                     metafile (EMF).
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.05.10
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.05.10
'___________________________________________________________________________
' TODO :
'       -
'       -
'___________________________________________________________________________
'
' The main module where we'll use the EnhMetaFileProc function.
'___________________________________________________________________________
'

Public Type METARECORD
    nSize As Long
    rdFunction As Integer       '//Low byte FunctionId, Hi byte Para count
    '//rdParm(1) As Integer
    lpParm As Integer           '// Parm[] Pointer to staring address of parameter array
End Type

Public Type ENHMETARECORD
    iType As Long
    nSize As Long
    lpParm As Long              '// Parm[] Pointer to staring address of parameter array
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function DeleteMetaFile Lib "gdi32.dll" (ByVal hMF As Long) As Long
Public Declare Function EnumMetaFile Lib "gdi32.dll" (ByVal hdc As Long, ByVal hMetafile As Long, ByVal lpMFEnumProc As Long, ByVal lParam As Long) As Long
Public Declare Function GetMetaFile Lib "gdi32.dll" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Public Declare Function PlayMetaFileRecord Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpHandletable As Long, ByRef lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function EnumEnhMetaFile Lib "gdi32.dll" (ByVal hdc As Long, ByVal hEMF As Long, ByVal lpEnhMetaFunc As Long, ByRef lpData As Any, ByRef lpRect As RECT) As Long
Public Declare Function GetEnhMetaFile Lib "gdi32.dll" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
Public Declare Function PlayEnhMetaFileRecord Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpHandletable As Long, ByRef lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function PlayEnhMetaFile Lib "gdi32.dll" (ByVal hdc As Long, ByVal hEMF As Long, ByRef lpRect As Any) As Long
Public Declare Function PlayMetaFile Lib "gdi32.dll" (ByVal hdc As Long, ByVal hMF As Long) As Long
Public Declare Function DeleteEnhMetaFile Lib "gdi32.dll" (ByVal hEMF As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Global Temp_Array() As String

Public Function EnhMetaFileProc(ByVal ClientHDC As Long, ByRef lpHandTab As Long, ByRef MetaRec As ENHMETARECORD, ByVal nHandles As Long, ByVal OptData As Long) As Integer
    Dim Address As Long
    Dim ResultArray(5) As Long
    Dim TabCurrent() As String
    Dim i As Integer
    Dim boolExists As Boolean
    
    Select Case MetaRec.iType
        Case 39                     'CREATEBRUSHINDIRECT
            
            CopyMemory ResultArray(0), MetaRec.lpParm, 16   'Copy from memory the lpParm to the ResultArray
            
            boolExists = False
            For i = 0 To UBound(Temp_Array) - 1                 'Parse the Temp_Array.
                TabCurrent = Split(Temp_Array(i), ";")
                If ResultArray(2) = CLng(TabCurrent(0)) Then    'If colour matches then we'll process later on.
                    boolExists = True
                    Exit For
                End If
            Next i
            
            If boolExists Then
                ResultArray(2) = CLng(TabCurrent(1))            'Substitute colours.
            Else
                If Fm_Metafile.Opt_Remaining(1).Value = True Then
                    'If we want to change all the remaining colours.
                    ResultArray(2) = Fm_Metafile.Pic_RemainingColours.BackColor
                End If
            End If
            Address = VarPtr(MetaRec.lpParm)
            CopyMemory ByVal Address, ResultArray(0), 16
    End Select
    EnhMetaFileProc = 1      '//1=Continue , 0=exit
End Function





