VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Declarations

Private Type ImageSize
    Height As Long
    Width As Long
End Type

Private Type This
    Cfg As Config
    Logger As SimpleLogger
End Type
Private This As This

'===============================================================================
' # Constructors

Private Sub Class_Initialize()
    With This
        Set .Cfg = Config.Bind_
        Set .Logger = New SimpleLogger
    End With
End Sub

Private Sub Class_Terminate()
    With This
        Set .Cfg = Nothing
        .Logger.Check
    End With
End Sub

'===============================================================================
' # Handlers


'===============================================================================
' # Entry points

Public Sub Prepare()
    
    With This
        If Not CheckFoldersForPrepare Then Exit Sub
        
        Dim TemplateFiles As Collection
        Set TemplateFiles = GetValidImagesFromFolder(.Cfg.RawTemplatesFolder)
        If Not CheckImagesCount(TemplateFiles, .Cfg.RawTemplatesFolder) Then Exit Sub
        
        Dim ImageFiles As Collection
        Set ImageFiles = GetValidImagesFromFolder(.Cfg.ImagesFolder)
        If Not CheckImagesCount(ImageFiles, .Cfg.ImagesFolder) Then Exit Sub
        
        Dim ImageSize As ImageSize
        ImageSize = GetImageSizeFromFile(ImageFiles(1))
        
        Dim File As Variant
        For Each File In TemplateFiles
            PrepareImageAndSave File, ImageSize
        Next File
    End With
End Sub

Public Sub SetOnTemplates()
End Sub

Public Sub Settings()
    With New SettingsView
        .RawTemplatesFolder = This.Cfg.RawTemplatesFolder
        .PreparedTemplatesFolder = This.Cfg.PreparedTemplatesFolder
        .ImagesFolder = This.Cfg.ImagesFolder
        .OutputFolder = This.Cfg.OutputFolder
        
        .Show vbModal
        
        This.Cfg.RawTemplatesFolder = .RawTemplatesFolder
        This.Cfg.PreparedTemplatesFolder = .PreparedTemplatesFolder
        This.Cfg.ImagesFolder = .ImagesFolder
        This.Cfg.OutputFolder = .OutputFolder
    End With
End Sub

'===============================================================================
' # Helpers

Private Sub PrepareImageAndSave( _
                ByVal TemplateFile As String, _
                ByRef ImageSize As ImageSize _
            )
    CreateDocument
    Dim Template As Shape
    Dim Frame As Shape
    Dim ExportFile As New FileSpec
    With ActiveDocument
        .ActiveLayer.Import TemplateFile
        Set Template = ActiveShape
        Set Frame = _
            .ActiveLayer.CreateRectangle2( _
                0, 0, _
                ToDocUnits(ImageSize.Width), _
                ToDocUnits(ImageSize.Height) _
            )
        Frame.CenterX = .ActivePage.CenterX
        Frame.CenterY = .ActivePage.CenterY
        Frame.OrderFrontOf Template
        Frame.Fill.ApplyUniformFill CreateRGBColor(255, 255, 255)
        Frame.CreateDropShadow _
            cdrDropShadowFlat, 50, 15, 0, 0, CreateRGBColor(0, 0, 0), _
            MergeMode:=cdrMergeMultiply
        ExportFile.Inject TemplateFile
        ExportFile.Path = This.Cfg.PreparedTemplatesFolder
        ExportFile.Ext = "cdr"
        .SaveAs ExportFile
        .Close
    End With
End Sub

Private Function GetImageSizeFromFile(ByVal ImageFile As String) As ImageSize
    Dim ImageSize As ImageSize
    CreateDocument
    With ActiveDocument
        .ActiveLayer.Import ImageFile
        With ActiveShape
            ImageSize.Height = .Bitmap.SizeHeight
            ImageSize.Width = .Bitmap.SizeWidth
        End With
        .Close
    End With
    GetImageSizeFromFile = ImageSize
End Function

Private Function CheckFoldersForPrepare() As Boolean
    CheckFoldersForPrepare = True
    With FSO
        If Not .FolderExists(This.Cfg.RawTemplatesFolder) Then
            CheckFoldersForPrepare = False
            This.Logger.Add "Не найдена папка исходных шаблонов " _
          & This.Cfg.RawTemplatesFolder
        End If
        If Not .FolderExists(This.Cfg.PreparedTemplatesFolder) Then
            CheckFoldersForPrepare = False
            This.Logger.Add "Не найдена папка подготовленных шаблонов " _
          & This.Cfg.PreparedTemplatesFolder
        End If
        If Not .FolderExists(This.Cfg.ImagesFolder) Then
            CheckFoldersForPrepare = False
            This.Logger.Add "Не найдена папка исходных изображений " _
          & This.Cfg.ImagesFolder
        End If
    End With
End Function

Private Function CheckFoldersForOutput() As Boolean
    CheckFoldersForOutput = CheckFoldersForPrepare
    If Not FSO.FolderExists(This.Cfg.OutputFolder) Then
        CheckFoldersForOutput = False
        This.Logger.Add "Не найдена папка " & This.Cfg.PreparedTemplatesFolder
    End If
End Function

Private Function CheckImagesCount( _
                     ByVal Images As Collection, _
                     ByVal Folder As String _
                 ) As Boolean
    If Images.Count = 0 Then
        This.Logger.Add "Не найдены изображения в папке " & Folder
    Else
        CheckImagesCount = True
    End If
End Function

Private Property Get GetValidImagesFromFolder( _
                         ByVal Folder As String _
                     ) As Collection
    With FSO
        Dim ValidFiles As New Collection
        Dim File As File
        For Each File In .GetFolder(Folder).Files
            If ValidImageExt(File.Name) Then
                ValidFiles.Add File
            End If
        Next File
    End With
    Set GetValidImagesFromFolder = ValidFiles
End Property

Private Property Get ValidImageExt(ByVal SomeFile As String) As Boolean
    Dim File As New FileSpec
    File.Inject SomeFile
    If File.Ext = "png" _
    Or File.Ext = "jpeg" _
    Or File.Ext = "jpg" Then
        ValidImageExt = True
    End If
End Property

Private Property Get ToDocUnits(ByVal SizeInPixels As Long) As Double
    ToDocUnits = ConvertUnits(SizeInPixels, cdrPixel, ActiveDocument.Unit)
End Property
