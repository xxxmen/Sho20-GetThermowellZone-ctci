Attribute VB_Name = "CommonSubs"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright 2001 Intergraph
'   All Rights Reserved
'
'   CommonSubs.bas
'   Common procedures added by CommandWizzard
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
#If DEBUG_COMPILE Then
Private DBS As New DebugSupport
#End If


Private Const MODULE = "GetInst"

Public Sub DDB()
#If DEBUG_COMPILE Then
    DBS.DEBUG_DEEP_BEGIN
#End If
End Sub

Public Sub DDE()
#If DEBUG_COMPILE Then
    DBS.DEBUG_DEEP_END
#End If
End Sub

Public Sub SET_DEBUG_SOURCE(ByVal source As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_SOURCE = source
#End If
End Sub

Public Sub DEBUG_MSG(message As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_MSG message
#End If
End Sub

Public Sub DEBUG_ERROR_MSG(message As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_ERROR_MSG message
#End If
End Sub

Public Sub DEBUG_MSG_ONCE(cookie As Integer, message As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_MSG_ONCE cookie, message
#End If
End Sub

Public Sub DEBUG_MSG_ONCE_RESET()
#If DEBUG_COMPILE Then
    DBS.DEBUG_MSG_ONCE_RESET
#End If
End Sub

Public Sub DEBUG_DUMP_OBJECT(element As Object, info As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_DUMP_OBJECT element, info
#End If
End Sub

Public Sub DEBUG_DUMP_LIST(List As Variant, info As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_DUMP_LIST List, info
#End If
End Sub

Public Sub DEBUG_DUMP_MATRIX(matrix As IJDT4x4, info As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_DUMP_MATRIX matrix, info
#End If
End Sub

Public Sub DEBUG_DUMP_POSITION(position As Object, info As String)
#If DEBUG_COMPILE Then
    DBS.DEBUG_DUMP_POSITION position, info
#End If
End Sub

Public Sub DEBUG_REFCOUNT_OBJECT(element As Object)
#If DEBUG_COMPILE Then
    DBS.DEBUG_REFCOUNT_OBJECT element
#End If
End Sub


