VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' Config v2023.10.09
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

#Const ReadOnly = False

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public RawTemplatesFolder As String
Public PreparedTemplatesFolder As String
Public ImagesFolder As String
Public OutputFolder As String

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.New_
    
        iniFilePrepare .Self, Action
        
        '.Section = "DEFAULT"
        .KeyVar "RawTemplatesFolder", RawTemplatesFolder
        .KeyVar "PreparedTemplatesFolder", PreparedTemplatesFolder
        .KeyVar "ImagesFolder", ImagesFolder
        .KeyVar "OutputFolder", OutputFolder
    
    End With
    
End Sub

'===============================================================================
'           * * * Контент, не предназначенный для изменения * * *
'===============================================================================
' # Constructors

Public Function Bind_() As Config
    Set Bind_ = New Config
    Bind_.ForceLoad
End Function

#If Not ReadOnly Then
Private Sub Class_Terminate()
    ForceSave
End Sub
#End If

'===============================================================================
' # Public

Public Sub ForceLoad()
    Assign iniRead
End Sub

#If Not ReadOnly Then
Public Sub ForceSave()
    Assign iniWrite
End Sub
#End If

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            #If DefaultContent Then
            MakeDefaultFile .File
            If Action = iniRead Then
                .Action = iniRead
            Else
                .Action = iniWrite
            End If
            #Else
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
            #End If
        End Select
    End With

End Sub

Private Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Private Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
