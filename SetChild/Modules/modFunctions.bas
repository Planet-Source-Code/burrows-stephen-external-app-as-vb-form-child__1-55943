Attribute VB_Name = "modFunctions"
'=======================================================================
'Title              :modFunctions
'System             :SetChild
'=======================================================================
'Copyright          :© Albion Software
'Date               :01/09/2004
'Author             :© BombDrop
'Technical Reviewer :
'Purpose            :
'=======================================================================
Option Explicit


'Used In System Menu Manipulation
 Const MF_BYCOMMAND As Long = &H0&
 Const MF_GRAYED As Long = &H1&
 Const SC_CLOSE As Long = &HF060&
 Const MF_ENABLED As Long = &H0&
 Const FOOLVB As Long = -10

'Application Class Names
Const ClassNameMSWord = "OpusApp"
Const ClassNameMSExcel = "XLMAIN"
Const ClassNameMSIExplorer = "IEFrame"
Const ClassNameMSVBasicIDE = "wndclass_desked_gsk"
Const ClassNameMSNotePad = "Notepad"
Const ClassNameMSVBApp = "ThunderForm"
Const ClassNameMSAccess = "OMain"
Const ClassNameMSPowePoint95 = "PP7FrameClass"
Const ClassNameMSPowePoint97 = "PP97FrameClass"
Const ClassNameMSPowePoint2000 = "PP9FrameClass"
Const ClassNameMSPowePointXP = "PP10FrameClass"
Const ClassNameMSFrontPage = "FrontPageExplorerWindow40"
Const ClassNameMSOutLook = "rctrl_renwd32"

'Used for Application Caption to aid in finding Child
Public Const AppTitle = "Test"

'Enumeration of Applications
Public Enum AppClass
    [MS Notepad]
    [MS Word]
    [MS Excel]
    [MS PowerPoint 95]
    [MS PowerPoint 97]
    [MS PowerPoint 2000]
    [MS PowerPoint XP]
    [MS Access]
    [MS Outlook]
    [Visual Bassic Application]
    [Visual Basic IDE]
    [MS Internet Explorer]
    [MS FrontPage]
    End Enum

'Used to find the application Child
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Used to Set the Parent of the Child
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long
    
'Used to get System menu of Child
 Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, _
    ByVal bRevert As Long) As Long

'Used to Modify the Child System menu
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" ( _
    ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
    
'Used to Redraw the Child System Menu
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long


       



'=======================================================================
'Procedure  :DisableClose (Sub)
'Date       :01/09/2004
'InPut      :ApplicationHandle As Long
'Returns    :N/A
'Author     :© BombDrop
'Purpose    :Will Disable the "X" button on the application passed to it
'=======================================================================
Private Sub DisableClose(ByRef ApplicationHandle As Long)

    Dim lngMemuHanle As Long

    'Get system menu handle for passed application
    lngMemuHanle = GetSystemMenu(ApplicationHandle, 0)

    If lngMemuHanle Then
        'Modify the menu
        Call ModifyMenu(lngMemuHanle, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED, _
            FOOLVB, "Close")

        'Redraw the menu
        Call DrawMenuBar(ApplicationHandle)
    End If 'lngMemuHanle

End Sub

Public Sub SetAsChild(ByVal ApplicationType As AppClass, ByRef Parent As Long)
    'Get handel of Word Apllication
Dim lngHandle   As Long
Dim lngFrame    As Long
    
    Select Case ApplicationType

        Case [MS Notepad]
            lngHandle = FindWindow(ClassNameMSNotePad, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS Word]
            lngHandle = FindWindow(ClassNameMSWord, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS Excel]
            lngHandle = FindWindow(ClassNameMSExcel, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS PowerPoint 95]
            lngHandle = FindWindow(ClassNameMSPowePoint95, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS PowerPoint 97]
            lngHandle = FindWindow(ClassNameMSPowePoint97, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS PowerPoint 2000]
            lngHandle = FindWindow(ClassNameMSPowePoint2000, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS PowerPoint XP]
            lngHandle = FindWindow(ClassNameMSPowePointXP, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS Access]
            lngHandle = FindWindow(ClassNameMSAccess, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS Outlook]
            lngHandle = FindWindow(ClassNameMSOutLook, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [Visual Bassic Application]
            lngHandle = FindWindow(ClassNameMSVBApp, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [Visual Basic IDE]
            lngHandle = FindWindow(ClassNameMSVBasicIDE, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS Internet Explorer]
            lngHandle = FindWindow(ClassNameMSIExplorer, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
        Case [MS FrontPage]
            lngHandle = FindWindow(ClassNameMSFrontPage, AppTitle)
            'Set the Word Application as a Child to the Form
            lngFrame = SetParent(lngHandle, Parent)

            DisableClose (lngHandle)
    End Select 'ApplicationType



End Sub
