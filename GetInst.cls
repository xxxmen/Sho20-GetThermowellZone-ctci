VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    ' NOOP
'
'   Copyright 2011 Intergraph
'   All Rights Reserved
'
'   GetInst.cls
'   ProgID:         GetInstFromFireZone.GetInst
'   Author:         <Your Name>
'   Creation Date:  Friday, Apr 29, 2011
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Implements IJCommand2

Private Const MODULE = "GetInst::"

    ' Declare class variables.
Private m_fOnIdleEnableState As Boolean
Private m_fIdleTimerDisabled As Boolean
Private m_oLocalizer As IJLocalizer
Private m_frmCmdForm As frmGetInst
Private m_iCmdID As Long
Private m_oCmnAppErrHandler As IJCommonError



'Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

            
            


Private Sub Class_Initialize()
    Set m_oCmnAppErrHandler = New CommonError
    On Error GoTo Errx

    ' create and initialize our localizer
    On Error Resume Next ' If localizer fails, don't stop
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path + "\" + App.EXEName
    Exit Sub
Errx:
    ' TODO - Wizard-generated code reports this error as "Warning". Modify as needed
    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrWarning, MODULE & "Initialize", _
                                            Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()

    Set m_oLocalizer = Nothing

    Set m_oCmnAppErrHandler = Nothing
End Sub


' EnableSpecial may call for significant processing. If so, error handler is meaningful
Private Property Get IJCommand2_EnableSpecial() As Boolean
    On Error GoTo Errx
    ' NOOP
    Exit Property
Errx:
    ' TODO - Wizard-generated code reports this error as "Warning". Modify as needed
    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrWarning, MODULE & "IJCommand2_EnableSpecial", _
                                            Err.Number, Err.Description
End Property


' Other properties simply return values. In general, error handler not needed

Private Property Get IJCommand2_EnableUIFlags() As Long
    IJCommand2_EnableUIFlags = 18
End Property

Private Property Get IJCommand2_EnableViewType() As String
    IJCommand2_EnableViewType = "IJDGraphicView"
End Property

Private Property Get IJCommand2_Modal() As Boolean
    IJCommand2_Modal = True
End Property

Private Sub IJCommand2_Stop()

End Sub

Private Property Get IJCommand2_Suspendable() As Boolean
    IJCommand2_Suspendable = False
Errx:
End Property

Private Property Get IJCommand2_Running() As Boolean
'Property Not Used
End Property


'   IJCommand2 Methods

Private Sub IJCommand2_Start(ByVal CommandID As Long, ByVal cmdArg As Variant)
    On Error GoTo Errx
    
    m_iCmdID = CommandID
    Dim oTrader As New Trader

    ' Disable Idle Timer
    Dim oIdleGen As IJIdleGenerator
    Set oIdleGen = oTrader.Service(TKIdleGenerator, vbNullString)
    m_fOnIdleEnableState = oIdleGen.Enable
    oIdleGen.Enable = False
    m_fIdleTimerDisabled = True
    ' Create form & display form since it's a Modal command.
    Set m_frmCmdForm = New frmGetInst
    m_frmCmdForm.Show vbModal

    Exit Sub
Errx:
    ' TODO - Wizard-generated code reports error as "Critical". Modify as needed
    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrCritical, MODULE & "IJCommand2_Start", _
                                            Err.Number, Err.Description
End Sub


'Private Sub IJCommand2_Stop()
'    On Error GoTo Errx
'    Dim oTrader As New Trader
'
''    Unload m_frmCmdForm
''    Set m_frmCmdForm = Nothing
'    If m_fIdleTimerDisabled Then
'        Dim oIdleGen As IJIdleGenerator
'        Set oIdleGen = oTrader.Service(TKIdleGenerator, vbNullString)
'        oIdleGen.Enable = m_fOnIdleEnableState
'    End If
'
'    Exit Sub
'Errx:
'    ' TODO - Wizard-generated code reports error as "Critical". Modify as needed
'    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrCritical, MODULE & "IJCommand2_Stop", _
'                                            Err.Number, Err.Description
'End Sub


Private Sub IJCommand2_Suspend()
    On Error GoTo Errx

    Exit Sub
Errx:
    ' TODO - Wizard-generated code reports error as "Critical". Modify as needed
    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrCritical, MODULE & "IJCommand2_Suspend", _
                                            Err.Number, Err.Description
End Sub


Private Sub IJCommand2_Resume()
    On Error GoTo Errx

    Exit Sub
Errx:
    ' TODO - Wizard-generated code reports error as "Critical". Modify as needed
    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrCritical, MODULE & "IJCommand2_Resume", _
                                            Err.Number, Err.Description
End Sub


Private Sub IJCommand2_OnIdle()
    On Error GoTo Errx
    Exit Sub
Errx:
    ' TODO - Wizard-generated code reports error as "Critical". Modify as needed
    m_oCmnAppErrHandler.ReportEnhancedError ERRORLEVEL_ErrCritical, MODULE & "IJCommand2_OnIdle", _
                                            Err.Number, Err.Description
End Sub


'End Of Template - V2
