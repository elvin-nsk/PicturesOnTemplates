Attribute VB_Name = "PicturesOnTemplates"
'===============================================================================
'   ������          : PicturesOnTemplates
'   ������          : 2023.11.19
'   �����           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "PicturesOnTemplates"

'===============================================================================

Sub Prepare()

    If RELEASE Then
        On Error GoTo Catch
        Optimization = True
    End If
    
    With New Main
        .Prepare
    End With
    
Finally:
    Optimization = False
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub SetOnTemplates()

    If RELEASE Then
        On Error GoTo Catch
        Optimization = True
    End If
    
    With New Main
        .SetOnTemplates
    End With
    
Finally:
    Optimization = False
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub Settings()

    If RELEASE Then On Error GoTo Catch
    
    With New Main
        .Settings
    End With
    
Finally:
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub
