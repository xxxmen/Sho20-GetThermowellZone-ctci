VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGetInst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GetInst"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7992
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   7992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnUpdateDia 
      Caption         =   "Update Dia."
      Height          =   435
      Left            =   6600
      TabIndex        =   9
      Top             =   3360
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   435
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1155
   End
   Begin VB.TextBox TextBox1 
      Height          =   372
      Left            =   6600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1080
      Width           =   1212
   End
   Begin MSComctlLib.ListView ListInstView 
      Height          =   3732
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   6252
      _ExtentX        =   11028
      _ExtentY        =   6583
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   6252
      _ExtentX        =   11028
      _ExtentY        =   868
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton ComUpload 
      Caption         =   "Import"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   2040
      Width           =   1155
   End
   Begin VB.CommandButton Comcancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   435
      Left            =   6600
      TabIndex        =   3
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton ComReport 
      Caption         =   "Report"
      Height          =   435
      Left            =   6600
      TabIndex        =   2
      Top             =   3960
      Width           =   1155
   End
   Begin VB.CommandButton btnDo 
      Caption         =   "Update"
      Height          =   435
      Left            =   6600
      TabIndex        =   0
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Instruments of Firezone:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGetInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oTrader As New Trader
Private m_oWS As IJDWorkingSet
Private m_oMTODom
Private m_oConn As IJDConnection
Private m_oConnCache As IJConnectionCache
Private m_oUOM As IJUnitsOfMeasure
Private m_oGfxViewMgr As IJDGraphicViews
Private m_oHiliter As IJHiliter
Private xpath As String
Public m_oDom As Object
''''''''''''''''''''''''''''''''''''''''DO REPORT
Private xlApp As Excel.Application
Private ListBook As Excel.Workbook
Private shtReport As Excel.Worksheet
'''''''''''''''''''''''''''''''''''''''''''''

Const ALTERNATE = 1
Const WINDING = 2
''''''''''''''''''''''''''''''TopMost 變數
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
            ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
            
Private Sub btnClose_Click()
    Unload Me
End Sub

Public Function GetDBConnectionByName(ByVal strDBName As String) As IJDConnection

   Dim oAppCtx As IJApplicationContext, strDBConn As String
   Set oAppCtx = m_oTrader.Service(TKApplicationContext, vbNullString)
   strDBConn = oAppCtx.DBTypeConfiguration.get_DataBaseFromDBType(strDBName)
   Set GetDBConnectionByName = m_oWS.Item(strDBConn)

End Function

Private Sub btnCancel_Click()
    
End Sub

Private Sub btnDo_Click()
    ListInstView.ListItems.Clear
    Dim m_oTxnmgr As IJTransactionMgr
    Set m_oTxnmgr = m_oTrader.Service(TKTransactionMgr, "")
    
    
    Set m_oUOM = m_oTrader.Service(TKUnitsOfMeasure, vbNullString)
    Set m_oWS = m_oTrader.Service(TKWorkingSet, vbNullString)
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''get TG
    Dim oCmd As IJADOCommand
    Set oCmd = New JCommand
    oCmd.QueryLanguage = LANGUAGE_SQL
    oCmd.CommandType = adCmdText
'    oCmd.CommandText = "SELECT oid from ROUTEPipeInstrumentOcc WHERE "
   
    oCmd.CommandText = "SELECT * FROM ROUTEPipeInstrumentOcc"
                      
                      
    oCmd.ActiveConnection = GetDBConnectionByName("Model").Name
    Dim oMkrEles As IJMonikerElements
    Set oMkrEles = oCmd.SelectObjects()
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''get pipe inst
'
'    Dim oCmdpipeInst As IJADOCommand
'    Set oCmdpipeInst = New JCommand
'
'    oCmdpipeInst.QueryLanguage = LANGUAGE_SQL
'    oCmdpipeInst.CommandType = adCmdText
'    oCmdpipeInst.CommandText = "SELECT oid from JRteInstrument "
'    oCmdpipeInst.ActiveConnection = GetDBConnectionByName("Model").Name
'    Dim oMkrElesInOnLineinst As IJMonikerElements
'    Set oMkrElesInOnLineinst = oCmdpipeInst.SelectObjects()
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''get eeraceway
'
'    Dim oCmdTray As IJADOCommand
'    Set oCmdTray = New JCommand
'
'    oCmdTray.QueryLanguage = LANGUAGE_SQL
'    oCmdTray.CommandType = adCmdText
'    oCmdTray.CommandText = "SELECT oid from JRteCableway "
'    oCmdTray.ActiveConnection = GetDBConnectionByName("Model").Name
'    Dim oMkrElesTray As IJMonikerElements
'    Set oMkrElesTray = oCmdTray.SelectObjects()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    Dim oObj As Object
    Dim oObjPermission As IJDObject
    Dim longPermissingroup As Long
    Dim listIndex As Integer
    listIndex = 1
    
    longPermissingroup = 117
    'longPermissingroup = 512
    
    Dim i As Integer
    Dim x As Double
    Dim y As Double
    Dim Z As Double
    Dim Xmm As Double
    Dim Ymm As Double
    Dim Zmm As Double


    '''''''''''''''''''''''''''''''''''''''''''''''get inon line  inst
     If oMkrEles.Count > 0 Then
' MsgBox oMkrEles.Count
            For i = 1 To oMkrEles.Count
                    Set oObj = oMkrEles.Bind(i)

                    If TypeOf oObj Is IJRteInstrumentOccur Then


                            Dim oPipeInst As IJRteInstrumentOccur
                            Set oObjPermission = oPipeInst
                            Dim oIJRtePathGenPart As IJRtePathGenPart
                            Dim oPipeInstName As IJNamedItem
                            Set oPipeInst = oObj
                            Set oPipeInstName = oPipeInst
                           
                                '''''''''GET RELATION FROM IJRtePathGenPart TO IJRtePathFeat
                                Dim AssocRel As IJDAssocRelation, Toc As IJDTargetObjectCol
                                Set AssocRel = oObj
                                Set Toc = AssocRel.CollectionRelations("IJRtePathGenPart", "DefiningFeature")    '''''role=collectionname=relation name(ends with relation collection)
                                Dim ele As IXMLDOMElement

                                Dim j As Integer
    
                                    For j = 1 To Toc.Count
                                            ProgressBar1.Value = 10
                                            Dim oPipePathFeat As IJRtePathFeat
                                            Set oPipePathFeat = Toc.Item(j)
                                            oPipePathFeat.GetLocation x, y, Z
                                            Xmm = x * 1000
                                            Ymm = y * 1000
                                            Zmm = Z * 1000
            '                                CheckInstInFireZone Xmm, Ymm, Zmm, ele, oPipeInstName.Name
            
                                           xpath = "SmartPlant3D/THERMOWELL[@TAG='" & Replace(oPipeInstName.Name, " ", "") & "']" ''''SP3D 與資料來源的tag名稱相同
                                           Set ele = Me.m_oDom.selectSingleNode(xpath)
                                           
                                            If Not ele Is Nothing Then
                                            'MsgBox 1234
                                                Dim objoid As String
            '                                    MsgBox ele.Attributes.getNamedItem("OID").Text '''''''zone's oid
                                        ''判斷從SP3D抓來的物件名稱是否有包含以下字串
                                                If (InStr(1, oPipeInstName.Name, "TE") Or InStr(1, oPipeInstName.Name, "TI") Or InStr(1, oPipeInstName.Name, "TIA") Or InStr(1, oPipeInstName.Name, "TIC") Or InStr(1, oPipeInstName.Name, "TICA") Or InStr(1, oPipeInstName.Name, "TIZA") Or InStr(1, oPipeInstName.Name, "TT") Or InStr(1, oPipeInstName.Name, "TW") Or InStr(1, oPipeInstName.Name, "TZT")) > 0 Then
                                                'MsgBox InStr(1, oPipeInstName.Name, TextBox1.Text)
                                                'If InStr(1, oPipeInstName.Name, TextBox1.Text) > 0 Then
                                                'MsgBox 123
                                                   Dim oAttributes As IJDAttributes
                                                   Dim oAttribute As IJDAttribute
                                                  Set oAttributes = oPipeInst
                                                  Dim MaintenanceL As Double
                                          
                                                  On Error Resume Next
                                                  MaintenanceL = oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value
    
                                                
                                         
                                                  oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value = CVar(CDbl(ele.Attributes.getNamedItem("LENGTH").Text))
                                        
                                                    ProgressBar1.Value = 50
                                                    ListInstView.ListItems.Add listIndex, , oPipeInstName.Name
                                                    ListInstView.ListItems(listIndex).SubItems(1) = oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value
                                                    objoid = GetObjectOid(oPipeInstName.Name, oObj) '''''''getoid
                                                    ListInstView.ListItems(listIndex).SubItems(2) = objoid
                                    'MsgBox objoid
                                    
                                                    If CVar(CDbl(ele.Attributes.getNamedItem("LENGTH").Text)) = oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value Then
                                                        ListInstView.ListItems(listIndex).SubItems(3) = "OK"
                                                        Else
                                                        ListInstView.ListItems(listIndex).SubItems(3) = "Failed"
                                                    End If
                                                    
                                                    m_oTxnmgr.Compute
                                                    m_oTxnmgr.Commit "THERMOWELL"
                                                    
                '                                    listIndexAfterAddEquip = listIndexAfterAddEquip + 1
                                                End If
                                       End If
                                    Next
                        
'
                    End If

            Next
            
            
         ProgressBar1.Value = 80
    End If
     ProgressBar1.Value = 100


End Sub
Private Function GetObjectOid(Equipname As String, oObj As IJDObject) As String

    Dim sMkr As IUnknown, strOid As String, sObjInfo As IJDObjectInfo
    Set sObjInfo = m_oConn
    Set sMkr = m_oConn.GetObjectName(oObj) ' Get Mkr
    strOid = sObjInfo.GetDbIdentifierFromMoniker(sMkr) ' Get Oid
    
    GetObjectOid = strOid
    
    Set oObj = Nothing
    Set sObjInfo = Nothing
    Set sMkr = Nothing ' Get Mkr
    
End Function


Private Sub CheckInstInFireZone(ByRef x As Double, y As Double, Z As Double, ByRef node As IXMLDOMElement, Equipname As String)
        Dim dom, xFile As String, ele As IXMLDOMElement, nodes As IXMLDOMNodeList      '''''''''''ele是每筆的資料，nodes是xml整份的list
        Dim xx As Double, yy As Double, Zmin As Double, Zmax As Double, rr As Double, Radius As Double

        Set dom = CreateObject("Microsoft.XMLDOM")
        xFile = GetSymbolShareName & "\CTCI\XML\FireProofingZone.XML"
        dom.Load xFile

        Set nodes = dom.selectNodes("FireProofingList/Zone")
        Set node = Nothing
        For Each ele In nodes
        
               Zmin = CLng(ele.Attributes.getNamedItem("Zmin").Text)
               Zmax = CLng(ele.Attributes.getNamedItem("Zmax").Text)

               If Zmin < Z And Zmax > Z Then
                   xx = CLng(ele.Attributes.getNamedItem("X").Text) - x  ''''''firezone的x減掉儀器的x
                   yy = CLng(ele.Attributes.getNamedItem("Y").Text) - y
                   Radius = CLng(ele.Attributes.getNamedItem("Radius").Text)
                   rr = Sqr(xx * xx + yy * yy)
                   If rr <= Radius Then
                     Set node = ele
                     Exit For
                   End If
        
               End If
         Next
         Set nodes = Nothing
         Set dom = Nothing
End Sub

Public Function GetSymbolShareName() As String
   Dim oJContext As IJContext
   Set oJContext = GetJContext()
   Dim strContextCabServerPath As String
   strContextCabServerPath = oJContext.GetVariable("CAB_SERVER")
   GetSymbolShareName = strContextCabServerPath
End Function


Private Sub BtnUpdateDia_Click()
    ListInstView.ListItems.Clear
    Dim m_oTxnmgr As IJTransactionMgr
    Set m_oTxnmgr = m_oTrader.Service(TKTransactionMgr, "")
    
    
    Set m_oUOM = m_oTrader.Service(TKUnitsOfMeasure, vbNullString)
    Set m_oWS = m_oTrader.Service(TKWorkingSet, vbNullString)
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''get TG
    Dim oCmd As IJADOCommand
    Set oCmd = New JCommand
    oCmd.QueryLanguage = LANGUAGE_SQL
    oCmd.CommandType = adCmdText
'    oCmd.CommandText = "SELECT oid from ROUTEPipeInstrumentOcc WHERE "
   
    oCmd.CommandText = "SELECT * FROM ROUTEPipeInstrumentOcc"
                      
                      
    oCmd.ActiveConnection = GetDBConnectionByName("Model").Name
    Dim oMkrEles As IJMonikerElements
    Set oMkrEles = oCmd.SelectObjects()
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''get pipe inst
'
'    Dim oCmdpipeInst As IJADOCommand
'    Set oCmdpipeInst = New JCommand
'
'    oCmdpipeInst.QueryLanguage = LANGUAGE_SQL
'    oCmdpipeInst.CommandType = adCmdText
'    oCmdpipeInst.CommandText = "SELECT oid from JRteInstrument "
'    oCmdpipeInst.ActiveConnection = GetDBConnectionByName("Model").Name
'    Dim oMkrElesInOnLineinst As IJMonikerElements
'    Set oMkrElesInOnLineinst = oCmdpipeInst.SelectObjects()
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''get eeraceway
'
'    Dim oCmdTray As IJADOCommand
'    Set oCmdTray = New JCommand
'
'    oCmdTray.QueryLanguage = LANGUAGE_SQL
'    oCmdTray.CommandType = adCmdText
'    oCmdTray.CommandText = "SELECT oid from JRteCableway "
'    oCmdTray.ActiveConnection = GetDBConnectionByName("Model").Name
'    Dim oMkrElesTray As IJMonikerElements
'    Set oMkrElesTray = oCmdTray.SelectObjects()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    Dim oObj As Object
    Dim oObjPermission As IJDObject
    Dim longPermissingroup As Long
    Dim listIndex As Integer
    listIndex = 1
    
    longPermissingroup = 117
    'longPermissingroup = 512
    
    Dim i As Integer
    Dim x As Double
    Dim y As Double
    Dim Z As Double
    Dim Xmm As Double
    Dim Ymm As Double
    Dim Zmm As Double


    '''''''''''''''''''''''''''''''''''''''''''''''get inon line  inst
     If oMkrEles.Count > 0 Then
' MsgBox oMkrEles.Count
            For i = 1 To oMkrEles.Count
                    Set oObj = oMkrEles.Bind(i)

                    If TypeOf oObj Is IJRteInstrumentOccur Then


                            Dim oPipeInst As IJRteInstrumentOccur
                            Set oObjPermission = oPipeInst
                            Dim oIJRtePathGenPart As IJRtePathGenPart
                            Dim oPipeInstName As IJNamedItem
                            Set oPipeInst = oObj
                            Set oPipeInstName = oPipeInst
                           
                                '''''''''GET RELATION FROM IJRtePathGenPart TO IJRtePathFeat
                                Dim AssocRel As IJDAssocRelation, Toc As IJDTargetObjectCol
                                Set AssocRel = oObj
                                Set Toc = AssocRel.CollectionRelations("IJRtePathGenPart", "DefiningFeature")    '''''role=collectionname=relation name(ends with relation collection)
                                Dim ele As IXMLDOMElement

                                Dim j As Integer
    
                                    For j = 1 To Toc.Count
                                            ProgressBar1.Value = 10
                                            Dim oPipePathFeat As IJRtePathFeat
                                            Set oPipePathFeat = Toc.Item(j)
                                            oPipePathFeat.GetLocation x, y, Z
                                            Xmm = x * 1000
                                            Ymm = y * 1000
                                            Zmm = Z * 1000
            '                                CheckInstInFireZone Xmm, Ymm, Zmm, ele, oPipeInstName.Name
            
                                           xpath = "SmartPlant3D/THERMOWELL[@TAG='" & Replace(oPipeInstName.Name, " ", "") & "']" ''''SP3D 與資料來源的tag名稱相同
                                           Set ele = Me.m_oDom.selectSingleNode(xpath)
                                           
                                            If Not ele Is Nothing Then
                                            'MsgBox 1234
                                                Dim objoid As String
            '                                    MsgBox ele.Attributes.getNamedItem("OID").Text '''''''zone's oid
                                        ''判斷從SP3D抓來的物件名稱是否有包含以下字串
                                                If (InStr(1, oPipeInstName.Name, "TE") Or InStr(1, oPipeInstName.Name, "TI") Or InStr(1, oPipeInstName.Name, "TIA") Or InStr(1, oPipeInstName.Name, "TIC") Or InStr(1, oPipeInstName.Name, "TICA") Or InStr(1, oPipeInstName.Name, "TIZA") Or InStr(1, oPipeInstName.Name, "TT") Or InStr(1, oPipeInstName.Name, "TW") Or InStr(1, oPipeInstName.Name, "TZT")) > 0 Then
                                                'MsgBox InStr(1, oPipeInstName.Name, TextBox1.Text)
                                                'If InStr(1, oPipeInstName.Name, TextBox1.Text) > 0 Then
                                                'MsgBox 123
                                                   Dim oAttributes As IJDAttributes
                                                   Dim oAttribute As IJDAttribute
                                                  Set oAttributes = oPipeInst
                                                  Dim MaintenanceL As Double
                                                  Dim MaintenanceD As Double
                                                  
                                          
                                                  On Error Resume Next
                                                  MaintenanceL = oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value
                                                  MaintenanceD = oAttributes.CollectionOfAttributes("IJUAInstrumentDimensions").Item("InstrumentDiameter").Value
                                                  
                                                  
                                                
                                         
                                                  oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value = CVar(CDbl(ele.Attributes.getNamedItem("LENGTH").Text))
                                                  oAttributes.CollectionOfAttributes("IJUAInstrumentDimensions").Item("InstrumentDiameter").Value = CVar(CDbl(ele.Attributes.getNamedItem("DIAMETER").Text))
                                                    
                                                    ProgressBar1.Value = 50
                                                    ListInstView.ListItems.Add listIndex, , oPipeInstName.Name
                                                    ListInstView.ListItems(listIndex).SubItems(1) = oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value
                                                    ListInstView.ListItems(listIndex).SubItems(2) = oAttributes.CollectionOfAttributes("IJUAInstrumentDimensions").Item("InstrumentDiameter").Value
                                                    objoid = GetObjectOid(oPipeInstName.Name, oObj) '''''''getoid
                                                    ListInstView.ListItems(listIndex).SubItems(2) = objoid
                                    'MsgBox objoid
                                    
                                                    If (CVar(CDbl(ele.Attributes.getNamedItem("LENGTH").Text)) = oAttributes.CollectionOfAttributes("IJUACTCIMaintenanceLength").Item("MaintenanceLength").Value) Then
                                                        ListInstView.ListItems(listIndex).SubItems(3) = "OK"
                                                        Else
                                                        ListInstView.ListItems(listIndex).SubItems(3) = "Failed"
                                                    End If
                                                    
                                                    m_oTxnmgr.Compute
                                                    m_oTxnmgr.Commit "THERMOWELL"
                                                    
                '                                    listIndexAfterAddEquip = listIndexAfterAddEquip + 1
                                                End If
                                       End If
                                    Next
                        
'
                    End If

            Next
            
            
         ProgressBar1.Value = 80
    End If
     ProgressBar1.Value = 100
End Sub

Private Sub ComReport_Click()
    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.Visible = True
    xlApp.Workbooks.Add
    
    Set ListBook = xlApp.Workbooks(1)
    Set shtReport = ListBook.Sheets(1)
    
    shtReport.Cells(1, 1) = "TAG.NO"
    shtReport.Cells(1, 2) = "MaintenanceLength"
    shtReport.Cells(1, 3) = "Status"
    
    Dim i As Integer
    Dim strFileName As String
    i = 1
    
    For i = 1 To ListInstView.ListItems.Count
        shtReport.Cells(i + 1, 1) = ListInstView.ListItems.Item(i).Text
        shtReport.Cells(i + 1, 2) = ListInstView.ListItems(i).SubItems(1)
        shtReport.Cells(i + 1, 3) = ListInstView.ListItems(i).SubItems(3)
    Next
    
    strFileName = "c:\54477\" & "Thermowell Inst" & ".xlsx"
    
    On Error Resume Next
    xlApp.Application.DisplayAlerts = False
    ListBook.SaveAs FileName:=strFileName
    '
    ListBook.Close
    Set ListBook = Nothing
    xlApp.Quit
    MsgBox "Finish"

End Sub

Private Sub ComUpload_Click()


    Dim fs As New Scripting.FileSystemObject
    Dim file As String
    Dim strConn As String
    Dim rf As ADODB.Recordset
    Dim cn As ADODB.Connection
    
    Me.MousePointer = 11
'先清空及設定部分值
    
    Set m_oDom = CreateObject("Microsoft.XMLDOM")

'    file = "D:\sp3dsp3d\GetThermowellZone\THERMO.mdb" 'Environ$("TEMP") + "\" + "SmartPlant3DUtilityVB.xml" mdb黨的路徑及名
     file = "c:\54477\THERMO.mdb"
    
    If fs.FileExists(file) = True Then
        Dim xml As String
        xml = "<?xml version='1.0' encoding='utf-16' ?><SmartPlant3D></SmartPlant3D>"
        m_oDom.loadXML xml
        

        Set cn = New ADODB.Connection

        With cn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & file & "; Jet OLEDB:Engine Type=5;"
            .Open
        End With

        Set rf = New ADODB.Recordset
        strConn = "SELECT * FROM thermowellL"
        rf.Open strConn, cn, 1, 3
        rf.MoveFirst
        Dim i As Integer
        i = 1
        Dim parent As IXMLDOMElement, ele As IXMLDOMElement
        Set parent = m_oDom.selectSingleNode("SmartPlant3D")

        
        Do While Not rf.EOF


            If i >= CInt(1) Then
            
            
                Dim str As String
                
                Set ele = m_oDom.createElement("THERMOWELL")
                
                If IsNull(rf.Fields(0)) Then
                    ele.setAttribute "TAG", ""
                Else
                    ele.setAttribute "TAG", rf.Fields(0)
                End If

                If IsNull(rf.Fields(1)) Then
                    ele.setAttribute "LENGTH", ""
                Else
                    ele.setAttribute "LENGTH", rf.Fields(1) / 1000#
                End If
                
                If IsNull(rf.Fields(2)) Then
                    ele.setAttribute "DIAMETER", ""
                Else
                    ele.setAttribute "DIAMETER", rf.Fields(2) / 1000#
                End If
                
                

'                    If IsNull(rf.Fields(2)) Then
'                        ele.setAttribute "TOSID", ""
'                    Else
'                        ele.setAttribute "TOSID", rf.Fields(2)
'                    End If
'                    If IsNull(rf.Fields(3)) Then
'                        ele.setAttribute "TOSSubID", ""
'                    Else
'                        ele.setAttribute "TOSSubID", rf.Fields(3)
'                    End If

                parent.appendChild ele
                Set ele = Nothing
            End If

'            With ListInstView.Col
'             .Cols = 3
'             .Rows = rf.RecordCount + 1
'             .TextMatrix(0, 0) = "PID Tag"
'             .TextMatrix(0, 1) = "PID TOS"
'             .TextMatrix(0, 2) = "PID SUB TOS"
'            End With



            ListInstView.ListItems.Add , , rf.Fields(0).Value
            ListInstView.ListItems(i).SubItems(1) = Val(rf.Fields(1).Value) / 1000#
            ListInstView.ListItems(i).SubItems(2) = Val(rf.Fields(2).Value) / 1000#
                        
            
            'MsgBox ListInstView.ListItems(i).SubItems(1)
            
            ' ListInstView.ListItems(i).SubItems(1) = rf.Fields(1).Value
''            MSFlexGridpid.TextMatrix(i, 1) = rf.Fields(2)
''            MSFlexGridpid.TextMatrix(i, 2) = rf.Fields(3)
'
''            ListPIDcount.AddItem rf.Fields(0)
'
            i = i + 1
            rf.MoveNext
'
          
                    
        Loop
        rf.Close
        cn.Close
        Set cn = Nothing

        m_oDom.Save "c:\54477\THERMOWELL.xml"
        Set ele = Nothing
        Set parent = Nothing
    Else
        Exit Sub
    End If
    Me.MousePointer = 0

MsgBox "Finish"

End Sub

Private Sub Form_Load()
    Set m_oWS = m_oTrader.Service(TKWorkingSet, vbNullString)
    Set m_oUOM = m_oTrader.Service(TKUnitsOfMeasure, vbNullString)
    Set m_oConn = m_oWS.ActiveConnection
    Set m_oConnCache = m_oConn
    
    Set m_oGfxViewMgr = m_oTrader.Service(TKGraphicViewMgr, vbNullString)
    Set m_oHiliter = m_oGfxViewMgr.CreateHiliter
    m_oHiliter.Color = vbYellow
    m_oHiliter.Weight = 2
        
    ListInstView.ColumnHeaders.Add , , "Instrument Tag"
    ListInstView.ColumnHeaders.Add , , "Maintenance Length"
    ListInstView.ColumnHeaders.Add , , "Maintenance Diameter"
    ListInstView.ColumnHeaders.Add , , "OID"
    ListInstView.ColumnHeaders.Add , , "Status"
    
    
'    ListInstView.ColumnHeaders(3).Width = 1
    Set m_oWS = m_oTrader.Service(TKWorkingSet, vbNullString)
    Set m_oUOM = m_oTrader.Service(TKUnitsOfMeasure, vbNullString)
    Set m_oConn = m_oWS.ActiveConnection
    Set m_oConnCache = m_oConn
    
    Set m_oGfxViewMgr = m_oTrader.Service(TKGraphicViewMgr, vbNullString)
    Set m_oHiliter = m_oGfxViewMgr.CreateHiliter
    m_oHiliter.Color = vbYellow
    m_oHiliter.Weight = 2

End Sub

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long

    SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'    MsgBox hwnd

End Function




Private Sub ListInstView_DblClick()
    m_oHiliter.Elements.Clear   '''''''''''''''hiliter 一定要把oid轉成object才能加
    
    Dim oGraphicView As IJDGraphicView
    Dim ocamera As IJCamera
    
    Dim oActiveConn As IJDConnection, oObjInfo As IJDObjectInfo
    Dim oMkr As IUnknown
    Set oActiveConn = m_oWS.ActiveConnection
    Set oObjInfo = oActiveConn
    Set oMkr = oObjInfo.GetMonikerFromDbIdentifier(ListInstView.SelectedItem.SubItems(2))
    Dim GetObject As IJDObject
    Set GetObject = oActiveConn.GetObject(oMkr)
    m_oHiliter.Elements.Add GetObject '''''''''''''''''''''hiliter
    
      
    Dim oCmdMgr As IJCommandManager2, oTrader As New Trader
    Set oCmdMgr = oTrader.Service(TKCommandMgr, vbNullString)
    Dim oSelset As IJSelectSet
    Set oSelset = oTrader.Service(TKSelectSet, vbNullString)
  
    oSelset.Elements.Add GetObject

    MsgBox "1"
  
'    Set oGraphicView = oTrader.Service(TKGraphicViewsControl, vbNullString)

    Dim lCmCmdLong As Long
'    lCmCmdLong = 0
    On Error Resume Next
   
    oCmdMgr.StartCommand "Gscad3dviewCmds.cviewfitcmd", HighPriority, lCmCmdLong, vbNullString
''   oGraphicView.Camera.Fit
'    Dim factor As Double
'    factor = 50
'    Set ocamera = oGraphicView.Camera
    
'    oGraphicView.Camera.Zoom (factor)
'    oGraphicView.Camera.Fit
'MsgBox "1" ''''''''''''''''''''''''''top most
'
    Dim lR As Long
    lR = SetTopMostWindow(Me.hwnd, True)

'MsgBox "2"
 m_oHiliter.Elements.Clear
 oSelset.Elements.Clear
End Sub

