VERSION 5.00
Begin VB.Form frmParent 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================
'Title              :frmParent
'System             :SetChild
'=======================================================================
'Copyright          :© Albion Software
'Date               :01/09/2004
'Author             :© BombDrop
'Technical Reviewer :
'Purpose            :To show how a VB form get be a parent of an other
'                   :application. This example show using MS Word
'=======================================================================


Private objWordApp  As Word.Application
Private objDoc      As Word.Document

Private Sub Form_Load()
   
    Set objWordApp = New Word.Application

    objWordApp.Caption = AppTitle
    objWordApp.Visible = True
    
    'Set form as parent to callen application
    Call modFunctions.SetAsChild([MS Word], Me.hWnd)
        
    'Add a document to the application
    Set objDoc = Word.Documents.Add
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Destory Application Object
    If Not objWordApp Is Nothing Then
        objWordApp.Quit
        Set objWordApp = Nothing
    End If
End Sub

Private Sub Form_Resize()
With objWordApp
    .Height = Me.ScaleHeight
    .Width = Me.ScaleWidth
    .Move 0, 0
End With 'objWordApp
End Sub

