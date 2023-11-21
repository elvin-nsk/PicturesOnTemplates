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
        Set .Cfg = New Config
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
        
        Dim RawTemplateFiles As Collection
        Set RawTemplateFiles = GetValidImagesFromFolder(.Cfg.RawTemplatesFolder)
        If Not CheckImagesCount(RawTemplateFiles, .Cfg.RawTemplatesFolder) Then _
            Exit Sub
        
        Dim ImageFiles As Collection
        Set ImageFiles = GetValidImagesFromFolder(.Cfg.ImagesFolder)
        If Not CheckImagesCount(ImageFiles, .Cfg.ImagesFolder) Then Exit Sub
        
        Dim ImageSize As ImageSize
        ImageSize = GetImageSizeFromFile(ImageFiles(1))
        
        Dim PBar As New ProgressBar
        PBar.Inject ProgressBarNumeric, RawTemplateFiles.Count
        PBar.Cancelable = True
        
        Dim File As Variant
        For Each File In RawTemplateFiles
            PBar.Update
            PrepareImageAndSave File, ImageSize
            If PBar.Canceled Then Exit Sub
        Next File
    End With
End Sub

Public Sub SetOnTemplates()
    With This
        If Not CheckFoldersForOutput Then Exit Sub
        
        Dim ImageFiles As Collection
        Set ImageFiles = GetValidImagesFromFolder(.Cfg.ImagesFolder)
        If Not CheckImagesCount(ImageFiles, .Cfg.ImagesFolder) Then Exit Sub
        
        Dim TemplateFiles As Collection
        Set TemplateFiles = _
            GetValidTemplatesFromFolder(.Cfg.PreparedTemplatesFolder)
        If Not _
            CheckImagesCount(TemplateFiles, .Cfg.PreparedTemplatesFolder) Then _
            Exit Sub
        
        Dim TemplateFile As Variant
        
        Dim PBar As New ProgressBar
        PBar.Inject ProgressBarNumeric, ImageFiles.Count
        PBar.Cancelable = True
        
        Dim Path As String
        Dim ImageSize As ImageSize
        Dim ImageIndex As Long
        For ImageIndex = 1 To ImageFiles.Count
            PBar.Update
            For Each TemplateFile In TemplateFiles
                SetOnTemplateAndExport _
                    TemplateFile, ImageFiles(ImageIndex), _
                    ImageIndex, Path, ImageSize
            Next TemplateFile
            FSO.CopyFile _
                ImageFiles(ImageIndex), _
                Path & "1-" & ImageSize.Width & "x" & ImageSize.Height _
              & "-" & GetFormattedIndex(ImageIndex) & ".jpg"
            If PBar.Canceled Then Exit Sub
        Next ImageIndex
    End With
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
        .ColorContext.BlendingColorModel = clrColorModelRGB
        .ActiveLayer.Import TemplateFile
        Set Template = ActiveShape
        ResizeImageToDocumentResolution Template
        Template.CenterX = .ActivePage.CenterX
        Template.CenterY = .ActivePage.CenterY
        Set Frame = _
            .ActiveLayer.CreateRectangle2( _
                0, 0, _
                PixelsToDocUnits(ImageSize.Width), _
                PixelsToDocUnits(ImageSize.Height) _
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

Private Sub SetOnTemplateAndExport( _
                ByVal TemplateFile As String, _
                ByVal ImageFile As String, _
                ByVal Index As Long, _
                ByRef outFolder As String, _
                ByRef outImageSize As ImageSize _
            )
    OpenDocument TemplateFile
    With ActiveDocument
        
        .ActiveLayer.Import ImageFile
        Dim Image As Shape
        Set Image = ActiveShape
        Dim Frames As Collection
        Set Frames = GetFrames()
        Dim Frame As Shape
        Set Frame = Frames(1)
        Image.SetSize Frame.SizeWidth, Frame.SizeHeight
        Image.CenterX = Frame.CenterX
        Image.CenterY = Frame.CenterY
        Image.OrderFrontOf Frame
        'здесь фреймы делаются прозрачными, пока убрал
        'SetFramesTransparent Frames
        
        outImageSize.Height = Image.Bitmap.SizeHeight
        outImageSize.Width = Image.Bitmap.SizeWidth
                
        Dim File As New FileSpec
        File.Inject TemplateFile
        File.Ext = "jpg"
        Dim Path As String
        Path = _
            This.Cfg.OutputFolder _
          & outImageSize.Width & "x" & outImageSize.Height _
          & "\" _
          & GetFormattedIndex(Index)
        File.Path = Path
        outFolder = File.Path
        MakePath outFolder
        
        Dim ExpFlt As ExportFilter
        With _
            .ExportBitmap( _
                File, cdrJPEG, cdrCurrentPage, cdrRGBColorImage, , , _
                .ResolutionX, .ResolutionX _
            )
            .Compression = 10
            .Optimized = True
            .Finish
        End With
        
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

Private Property Get GetFrames() As Collection
    Set GetFrames = New Collection
    Dim Shape As Shape
    For Each Shape In ActivePage.Shapes
        If Shape.Type = cdrRectangleShape Then
            GetFrames.Add Shape
        End If
    Next Shape
End Property

Private Sub SetFramesTransparent(ByVal Frames As Collection)
    Dim Frame As Shape
    For Each Frame In Frames
        Frame.Fill.ApplyNoFill
        Frame.Outline.SetNoOutline
    Next Frame
End Sub

Private Property Get GetFormattedIndex(ByVal Index As Long) As String
    GetFormattedIndex = VBA.Format$(Index, "0000000")
End Property

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

Private Property Get GetValidTemplatesFromFolder( _
                         ByVal Folder As String _
                     ) As Collection
    With FSO
        Dim ValidFiles As New Collection
        Dim File As File
        For Each File In .GetFolder(Folder).Files
            If ValidTemplateExt(File.Name) Then
                ValidFiles.Add File
            End If
        Next File
    End With
    Set GetValidTemplatesFromFolder = ValidFiles
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

Private Property Get ValidTemplateExt(ByVal SomeFile As String) As Boolean
    Dim File As New FileSpec
    File.Inject SomeFile
    If File.Ext = "cdr" Then
        ValidTemplateExt = True
    End If
End Property
