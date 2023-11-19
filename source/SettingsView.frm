VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsView 
   Caption         =   "Íàñòðîéêè"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "SettingsView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public RawTemplatesFolder As FolderBrowserHandler
Public PreparedTemplatesFolder As FolderBrowserHandler
Public ImagesFolder As FolderBrowserHandler
Public OutputFolder As FolderBrowserHandler

'===============================================================================

Private Sub UserForm_Initialize()
    Set RawTemplatesFolder = _
        FolderBrowserHandler.New_( _
            RawTemplatesFolderBox, _
            RawTemplatesFolderBrowse _
        )
    Set PreparedTemplatesFolder = _
        FolderBrowserHandler.New_( _
            PreparedTemplatesFolderBox, _
            PreparedTemplatesFolderBrowse _
        )
    Set ImagesFolder = _
        FolderBrowserHandler.New_( _
            ImagesFolderBox, _
            ImagesFolderBrowse _
        )
    Set OutputFolder = _
        FolderBrowserHandler.New_( _
            OutputFolderBox, _
            OutputFolderBrowse _
        )
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub ButtonClose_Click()
    FormClose
End Sub

'===============================================================================

Private Sub FormClose()
    Me.Hide
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Ñancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Ñancel = True
        FormClose
    End If
End Sub
