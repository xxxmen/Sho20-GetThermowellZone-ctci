Attribute VB_Name = "CoreTraderKeys"
'*******************************************************************
'  Copyright (C) 1998-2003, Intergraph Corporation.  All rights reserved.
'
'  File : CoreTraderKeys.bas
'
'  Description:
'   Public Consts for Trader'd objects supplied or used by the Core
'   "TK" prefix stands for "TraderKey"
'   "TKT" is prefix for "Trader Key Type" to be used in defining
'   2nd arguments for Traded objects
'
'  Change History:
'   05.03.98    cgp     Created
'   19.03.98    cgp     Add ValueMgr, PreferencesMgr
'   27-Mar-1998 JJH     added Material Stylemanager
'   10.04.98    cgp     added TaskInfo
'   29.04.98    cgp     added TaskInfo consts, also DUI consts
'   29.04.1998  kek     added SelectCmdFilter
'   30.04.98    cgp     added TKElementsFilter & subtype EFLocateFilter
'                       In 2.0, Core & Apps must switch to this and
'                       TKSelectCmdFilter removed (slipped in because of IGUG)
'   12.05.98    cgp     Add TKMetaHiliterFactory
'   22 may 98   cgp     Add other tk's
'   7 jul 98    cgp     Add TKIdleGenerator
'   15 jul 98   aic     Add TKFileMenu
'   15 Jul 98   cgp     add DebugLog
'   20 Jul 98   kek     add HTMLHelper
'   30 Jul 98   kek     add TKTaskSwitchMgr
'    6 Aug 98   kek     add TKSelCmdRibbon
'   14.aug.98   cgp     add TKElementsSorter, ESSmartLocateSorter, TKSmartLocateProgID
'   25.aug.98   cgp     add EyeView DUI location reference
'   26.aug.98   cgp     add TKViewSet, VSGraphic
'   22.sep.98   kek     add TKDraftView
'   22.sep.98   cgp     remove TKElementsSorter, ESSmartLocateSorter, EFSmartLocateFilter
'    3.oct.98   kek     add TKContextConfiguration
'   20.oct.98   kek     add TKDefModCmdRibbon
'   28.oct.98   kek     add TKElementsMultiCaster
'   28.oct.98   kek     removed TKElementsMultiCaster added TKSelectMultiCaster
'   18.nov.98   lb      added TKApplicationContext - removed TKContextConfiguration
'   23.nov.98   kek     add VMKLocateFilters
'    3.dec.98   kek     remove TKHTMLHelper - functions moved to DynamicUI
'    4.jan.99   elb     add TKUnitsOfMeasure
'   26.jan.99   kek     add TKRAD2DView
'   9.feb.99    cgp     add TKMetaLocatorProgID, TKProxyLocatorProgID
'   12.feb.99   cgp     put TKProxyLocatorProgID back to version-independent string
'   25.feb.99   cgp     add TKHelp
'   18.mar.99   elb     add TKATPTestApp
'    8.apr.99   kek     add TKUserObjectProperties
'   14.apr.99   kek     add PRFColorHighlight
'   17.aug.00   ryz     add TKPostSelectSelectSetUpdate
'   12.mar.2001 cgp     add TKToolbarMgr
'
'*******************************************************************

Option Explicit


Public Const TKToolbarMgr = "ToolbarMgr"

Public Const TKHelp = "Help"

Public Const TKViewSet = "ViewSet"
Public Const VSGraphic = "Graphic"

Public Const TKSmartLocateProgID = "IMSSmartLocate.SmartLocate"
Public Const TKMetaLocatorProgID = "IMSMetaLocator.MetaLocator"
Public Const TKProxyLocatorProgID = "ProxyLocator.ProxyLocator"
Public Const TKLocatorProgID = "IMSLocator.Locator"

Public Const TKDebugLog = "DebugLog"
Public Const TKFileMenu = "FileMenu"
Public Const TKIdleGenerator = "IdleGenerator"

Public Const TKGraphicViewsControl = "GraphicViewsControl"
Public Const TKEnvironmentMgr = "EnvironmentMgr"

Public Const TKMenuService = "MenuService"

'MSGSCADHost and GSCADHost are deprecated and should not be used.
'Applications should use TKTaskHost
Public Const MSGSCADHost = "GSCADHost"
Public Const TKTaskHost = "TaskHost"

Public Const TKMainFrame = "MainFrame"
Public Const TKDynamicUI = "DynamicUI"
Public Const TKModifyCommandBroker = "ModifyCommandBroker"

Public Const TKMetaHiliterFactory = "MetaHiliterFactory"

Public Const TKDraftView = "DraftViewControl"

Public Const TKElementsFilter = "ElementsFilter"
Public Const EFLocateFilter = "Locate"      ' sub type

'Following are to be used w/ DynamicUI service.
'They represent well-known IDREF's
Public Const DUIEyeView = "EyeView"
Public Const DUIRibbonBar = "Ribbonbar"
Public Const DUIStdToolbar = "Toolbar"
Public Const DUIStatusBar = "Statusbar"
Public Const DUIGraphicViews = "GraphicViews"

Public Const TKTaskInfo = "TaskInfo"

'Following are standard, required TaskInfo keys
Public Const TITaskName = "TaskName"
Public Const TITaskVersion = "TaskVersion"

Public Const TKValueMgr = "ValueMgr"
Public Const TKPreferences = "Preferences"
Public Const TKTTaskPreferences = "Task"
Public Const TKTSessionPreferences = "Session"

Public Const TKCommandMgr = "CommandMgr"
Public Const TKCmdContinuationMgr = "CommandContinuationMgr"
Public Const TKTransactionMgr = "TransactionMgr"
Public Const TKSessionMgr = "SessionMgr"
Public Const TKGraphicViewMgr = "GraphicViewMgr"
Public Const TKStyleMgr = "StyleMgr"

Public Const TKSelectSet = "SelectSet"
Public Const TKWorkingSet = "WorkingSet"
Public Const TKStartWorkingSet = "StartWorkingSet"
Public Const TKAppInfo = "AppInfo"
Public Const TKErrorHandler = "ErrorHandler"
Public Const TKActiveView = "ActiveView"
Public Const TKRAD2DView = "RAD2DView"
Public Const TKStatusBar = "StatusBar"
Public Const TKSelCmdRibbon = "SelectCmdRibbonBar"
Public Const TKDefModCmdRibbon = "DefaultModifyCmdRibbonBar"
Public Const TKSelectMultiCaster = "SelectMultiCaster"
Public Const TKLocateMultiCaster = "LocateMultiCaster"
Public Const TKUserObjectProperties = "User Object Properties"
Public Const TKDropSource = "DropSource"

Public Const TKMaterialStyleManager = "StyleManager"
Public Const TKTMaterialStyleManager = "Material"

Public Const TKTaskSwitchMgr = "TaskSwitchMgr"

Public Const TKApplicationContext = "ApplicationContext"

'ValueManager keys
Public Const VMKLocateFilters = "LocateFilters"

'Preference Keys - moved to CorePreferenceKeys.bas

'Key for the Units Of Measure Service
Public Const TKUnitsOfMeasure = "UnitsOfMeasure"

'Key for the ATPTestApp
Public Const TKATPTestApp = "ATPTestApp"

Public Const TKPostSelectSelectSetUpdate = "PostSelectSelectSetUpdate"
Public Const TKObjectAttributeSvc = "ObjectAttributeSvc"
Public Const TKToolTipService = "ToolTipService"
Public Const TKToDoListService = "ToDoListService"

