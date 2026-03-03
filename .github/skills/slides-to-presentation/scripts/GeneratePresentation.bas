Attribute VB_Name = "GeneratePresentation"
'==============================================================================
' GeneratePresentation.bas
'
' VBA Macro for Microsoft PowerPoint
' Generates a presentation from a JSON data file containing slide content
' and speaker notes.
'
' Prerequisites:
'   1. Generate the JSON file using generate_pptx.py:
'      python generate_pptx.py --slides slides.md --notes speaker-notes.md --export-json slides_data.json
'   2. Import this module into PowerPoint VBA (Alt+F11 > File > Import)
'   3. Run the macro (Alt+F8 > GeneratePresentation > Run)
'
' Note: Requires "Microsoft Scripting Runtime" reference for Dictionary.
'       The macro will attempt to add it automatically.
'==============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Configuration Constants
' ---------------------------------------------------------------------------
Private Const TITLE_FONT_NAME As String = "Calibri"
Private Const TITLE_FONT_SIZE As Long = 36
Private Const TITLE_FONT_COLOR As Long = &H90502E  ' #2E5090 in BGR

Private Const SUBTITLE_FONT_NAME As String = "Calibri"
Private Const SUBTITLE_FONT_SIZE As Long = 24
Private Const SUBTITLE_FONT_COLOR As Long = &H555555

Private Const BODY_FONT_NAME As String = "Calibri"
Private Const BODY_FONT_SIZE As Long = 20
Private Const BODY_FONT_COLOR As Long = &H333333

Private Const CODE_FONT_NAME As String = "Consolas"
Private Const CODE_FONT_SIZE As Long = 14
Private Const CODE_FONT_COLOR As Long = &H333333

Private Const ACCENT_COLOR As Long = &H4F53D9  ' #D9534F in BGR

' Image constraints (in points; 1 inch = 72 points)
Private Const IMG_MAX_WIDTH As Single = 648  ' 9 inches
Private Const IMG_MAX_HEIGHT As Single = 360 ' 5 inches
Private Const IMG_LEFT As Single = 72        ' 1 inch from left
Private Const IMG_TOP As Single = 144        ' 2 inches from top

' Slide dimensions (widescreen 16:9)
Private Const SLIDE_WIDTH As Single = 960    ' 13.333 inches
Private Const SLIDE_HEIGHT As Single = 540   ' 7.5 inches

' Template path (leave empty for blank presentation)
Private Const TEMPLATE_PATH As String = ""


' ---------------------------------------------------------------------------
' Main Entry Point
' ---------------------------------------------------------------------------
Public Sub GeneratePresentation()

    Dim jsonPath As String
    Dim jsonText As String
    Dim slides As Collection
    Dim prs As Presentation
    Dim i As Long

    ' Prompt user to select JSON file
    jsonPath = SelectJSONFile()
    If jsonPath = "" Then
        MsgBox "No file selected. Macro cancelled.", vbInformation
        Exit Sub
    End If

    ' Read JSON file
    jsonText = ReadFileContent(jsonPath)
    If jsonText = "" Then
        MsgBox "Could not read file: " & jsonPath, vbCritical
        Exit Sub
    End If

    ' Parse JSON
    Set slides = ParseSlidesJSON(jsonText)
    If slides Is Nothing Or slides.Count = 0 Then
        MsgBox "No slides found in JSON file.", vbWarning
        Exit Sub
    End If

    ' Determine diagrams directory (same folder as JSON)
    Dim diagramsDir As String
    diagramsDir = GetParentFolder(jsonPath) & "\diagrams"

    ' Create presentation
    If TEMPLATE_PATH <> "" And Dir(TEMPLATE_PATH) <> "" Then
        Set prs = Presentations.Open(TEMPLATE_PATH)
    Else
        Set prs = Presentations.Add(msoTrue)
    End If

    ' Set slide size to widescreen
    prs.PageSetup.SlideWidth = SLIDE_WIDTH
    prs.PageSetup.SlideHeight = SLIDE_HEIGHT

    ' Generate slides
    Application.StatusBar = "Generating presentation..."

    For i = 1 To slides.Count
        Dim slideData As Object
        Set slideData = slides(i)

        Application.StatusBar = "Creating slide " & i & " of " & slides.Count & "..."

        If IsLeadSlide(slideData) Then
            CreateTitleSlide prs, slideData, i
        Else
            CreateContentSlide prs, slideData, diagramsDir, i
        End If
    Next i

    ' Remove the initial blank slide if the template added one
    If prs.Slides.Count > slides.Count Then
        prs.Slides(prs.Slides.Count).Delete
    End If

    Application.StatusBar = "Presentation generated: " & slides.Count & " slides"
    MsgBox "Presentation generated successfully!" & vbCrLf & _
           slides.Count & " slides created.", vbInformation

End Sub


' ---------------------------------------------------------------------------
' Slide Creation
' ---------------------------------------------------------------------------

Private Sub CreateTitleSlide(prs As Presentation, slideData As Object, index As Long)

    Dim sld As Slide
    Set sld = prs.Slides.Add(index, ppLayoutTitle)

    ' Title
    If sld.Shapes.HasTitle Then
        With sld.Shapes.Title.TextFrame.TextRange
            .Text = GetStringValue(slideData, "title")
            .Font.Name = TITLE_FONT_NAME
            .Font.Size = 44
            .Font.Bold = msoTrue
            .Font.Color.RGB = TITLE_FONT_COLOR
        End With
    End If

    ' Subtitle
    If sld.Shapes.Placeholders.Count > 1 Then
        Dim bodyText As String
        bodyText = JoinBodyLines(slideData)
        If bodyText = "" Then bodyText = GetStringValue(slideData, "subtitle")

        With sld.Shapes.Placeholders(2).TextFrame.TextRange
            .Text = bodyText
            .Font.Name = SUBTITLE_FONT_NAME
            .Font.Size = SUBTITLE_FONT_SIZE
            .Font.Color.RGB = SUBTITLE_FONT_COLOR
        End With
    End If

    ' Speaker notes
    AddSpeakerNotes sld, slideData

End Sub


Private Sub CreateContentSlide(prs As Presentation, slideData As Object, _
                                diagramsDir As String, index As Long)

    Dim sld As Slide
    Dim hasImages As Boolean
    hasImages = HasSlideImages(slideData)

    ' Use blank layout for image-heavy slides, content layout otherwise
    If hasImages And CountBodyLines(slideData) <= 2 Then
        Set sld = prs.Slides.Add(index, ppLayoutBlank)
        AddTitleShape sld, slideData
    Else
        Set sld = prs.Slides.Add(index, ppLayoutText)
        ' Title
        If sld.Shapes.HasTitle Then
            Dim titleText As String
            titleText = GetStringValue(slideData, "title")
            Dim subtitle As String
            subtitle = GetStringValue(slideData, "subtitle")
            If subtitle <> "" Then titleText = titleText & " " & Chr(8212) & " " & subtitle

            With sld.Shapes.Title.TextFrame.TextRange
                .Text = titleText
                .Font.Name = TITLE_FONT_NAME
                .Font.Size = TITLE_FONT_SIZE
                .Font.Bold = msoTrue
                .Font.Color.RGB = TITLE_FONT_COLOR
            End With
        End If
    End If

    ' Body content
    If CountBodyLines(slideData) > 0 Then
        AddBodyContent sld, slideData
    End If

    ' Code blocks
    AddCodeBlocks sld, slideData

    ' Images
    If hasImages Then
        AddImages sld, slideData, diagramsDir
    End If

    ' Speaker notes
    AddSpeakerNotes sld, slideData

End Sub


' ---------------------------------------------------------------------------
' Content Helpers
' ---------------------------------------------------------------------------

Private Sub AddTitleShape(sld As Slide, slideData As Object)
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
              36, 18, SLIDE_WIDTH - 72, 60)

    With shp.TextFrame.TextRange
        .Text = GetStringValue(slideData, "title")
        .Font.Name = TITLE_FONT_NAME
        .Font.Size = TITLE_FONT_SIZE
        .Font.Bold = msoTrue
        .Font.Color.RGB = TITLE_FONT_COLOR
    End With
End Sub


Private Sub AddBodyContent(sld As Slide, slideData As Object)
    ' Find or create a body placeholder
    Dim bodyShape As Shape
    Set bodyShape = FindBodyPlaceholder(sld)

    If bodyShape Is Nothing Then
        ' Create a text box
        Set bodyShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        72, 100, SLIDE_WIDTH - 144, SLIDE_HEIGHT - 180)
    End If

    Dim tf As TextFrame
    Set tf = bodyShape.TextFrame
    tf.WordWrap = msoTrue

    Dim bodyLines As Object
    Set bodyLines = GetArrayValue(slideData, "body_lines")

    If bodyLines Is Nothing Then Exit Sub

    Dim lineItem As Variant
    Dim firstLine As Boolean
    firstLine = True

    Dim lineIdx As Long
    For lineIdx = 1 To bodyLines.Count
        Dim lineText As String
        lineText = CStr(bodyLines(lineIdx))

        If Len(Trim(lineText)) = 0 Then GoTo NextLine

        Dim para As TextRange
        If firstLine Then
            Set para = tf.TextRange
            para.Text = lineText
            firstLine = False
        Else
            Set para = tf.TextRange.InsertAfter(vbCrLf & lineText)
        End If

        ' Format the paragraph
        With tf.TextRange.Paragraphs(tf.TextRange.Paragraphs.Count)
            .Font.Name = BODY_FONT_NAME
            .Font.Size = BODY_FONT_SIZE
            .Font.Color.RGB = BODY_FONT_COLOR
            .ParagraphFormat.SpaceAfter = 6
            .ParagraphFormat.Bullet.Type = ppBulletUnnumbered
        End With

NextLine:
    Next lineIdx

End Sub


Private Sub AddCodeBlocks(sld As Slide, slideData As Object)
    Dim codeBlocks As Object
    Set codeBlocks = GetArrayValue(slideData, "code_blocks")
    If codeBlocks Is Nothing Then Exit Sub
    If codeBlocks.Count = 0 Then Exit Sub

    Dim codeText As String
    codeText = CStr(codeBlocks(1))  ' First code block only

    ' Truncate if too long
    If Len(codeText) > 800 Then
        codeText = Left(codeText, 800) & vbCrLf & "..."
    End If

    ' Position below existing content
    Dim topPos As Single
    topPos = GetLowestShapeBottom(sld) + 12

    If topPos > SLIDE_HEIGHT - 100 Then Exit Sub  ' No room

    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
              72, topPos, SLIDE_WIDTH - 144, SLIDE_HEIGHT - topPos - 36)

    With shp.TextFrame
        .WordWrap = msoTrue
        .TextRange.Text = codeText
        .TextRange.Font.Name = CODE_FONT_NAME
        .TextRange.Font.Size = CODE_FONT_SIZE
        .TextRange.Font.Color.RGB = CODE_FONT_COLOR
    End With

    ' Light background for code
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(245, 245, 245)
        .Transparency = 0
    End With

    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(200, 200, 200)
        .Weight = 1
    End With

End Sub


Private Sub AddImages(sld As Slide, slideData As Object, diagramsDir As String)
    Dim images As Object
    Set images = GetArrayValue(slideData, "images")
    If images Is Nothing Then Exit Sub

    Dim imgIdx As Long
    For imgIdx = 1 To images.Count
        Dim imgObj As Object
        Set imgObj = images(imgIdx)

        Dim imgPath As String
        imgPath = GetStringValue(imgObj, "path")

        ' Resolve full path
        Dim fullPath As String
        fullPath = diagramsDir & "\" & GetFileName(imgPath)

        ' Try PNG first (pre-converted), then original
        Dim pngPath As String
        pngPath = Replace(fullPath, ".svg", ".png")

        Dim actualPath As String
        If Dir(pngPath) <> "" Then
            actualPath = pngPath
        ElseIf Dir(fullPath) <> "" Then
            actualPath = fullPath
        Else
            ' Try path relative to diagrams parent
            Dim altPath As String
            altPath = GetParentFolder(diagramsDir) & "\" & Replace(imgPath, "/", "\")
            If Dir(altPath) <> "" Then
                actualPath = altPath
            Else
                GoTo NextImage
            End If
        End If

        ' Add picture
        On Error Resume Next
        Dim pic As Shape
        Set pic = sld.Shapes.AddPicture( _
            FileName:=actualPath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=IMG_LEFT, _
            Top:=IMG_TOP, _
            Width:=-1, _
            Height:=-1)

        If Not pic Is Nothing Then
            ' Scale to fit within constraints
            ScaleImage pic, IMG_MAX_WIDTH, IMG_MAX_HEIGHT

            ' Center horizontally
            pic.Left = (SLIDE_WIDTH - pic.Width) / 2
        End If
        On Error GoTo 0

NextImage:
    Next imgIdx
End Sub


Private Sub AddSpeakerNotes(sld As Slide, slideData As Object)
    Dim notesText As String
    notesText = GetStringValue(slideData, "speaker_notes")

    If notesText <> "" Then
        ' Clean markdown formatting from notes
        notesText = CleanMarkdown(notesText)
        sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notesText
    End If
End Sub


Private Sub ScaleImage(shp As Shape, maxW As Single, maxH As Single)
    ' Scale proportionally to fit within max dimensions
    Dim ratio As Single

    If shp.Width > maxW Then
        ratio = maxW / shp.Width
        shp.Width = maxW
        shp.Height = shp.Height * ratio
    End If

    If shp.Height > maxH Then
        ratio = maxH / shp.Height
        shp.Height = maxH
        shp.Width = shp.Width * ratio
    End If
End Sub


' ---------------------------------------------------------------------------
' JSON Parsing (Lightweight — no external dependencies)
' ---------------------------------------------------------------------------
' This is a minimal JSON parser for the specific structure we expect.
' For production use, consider a proper JSON library.

Private Function ParseSlidesJSON(jsonText As String) As Collection
    ' Simple approach: use ScriptControl for JSON parsing
    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"

    ' Parse JSON
    sc.AddCode "function parseJSON(s) { return eval('(' + s + ')'); }"
    sc.AddCode "function getSlides(obj) { return obj.slides; }"
    sc.AddCode "function getCount(arr) { return arr.length; }"
    sc.AddCode "function getItem(arr, i) { return arr[i]; }"
    sc.AddCode "function getProp(obj, prop) { try { var v = obj[prop]; return (v === null || v === undefined) ? '' : v; } catch(e) { return ''; } }"
    sc.AddCode "function isArray(v) { return Object.prototype.toString.call(v) === '[object Array]'; }"
    sc.AddCode "function getArrLen(obj, prop) { try { var v = obj[prop]; if (!v) return 0; return v.length; } catch(e) { return 0; } }"
    sc.AddCode "function getArrItem(obj, prop, i) { return obj[prop][i]; }"

    Dim parsed As Object
    Set parsed = sc.Run("parseJSON", jsonText)

    Dim slidesArr As Object
    Set slidesArr = sc.Run("getSlides", parsed)

    Dim count As Long
    count = CLng(sc.Run("getCount", slidesArr))

    Dim result As New Collection
    Dim i As Long

    For i = 0 To count - 1
        Dim slideObj As Object
        Set slideObj = sc.Run("getItem", slidesArr, i)

        ' Wrap in a helper object that stores the ScriptControl reference
        Dim wrapper As New Collection
        wrapper.Add slideObj, "obj"
        wrapper.Add sc, "sc"

        ' Build a dictionary-like structure
        Dim slideDict As Object
        Set slideDict = CreateObject("Scripting.Dictionary")

        slideDict("title") = CStr(sc.Run("getProp", slideObj, "title"))
        slideDict("subtitle") = CStr(sc.Run("getProp", slideObj, "subtitle"))
        slideDict("is_lead") = CBool(sc.Run("getProp", slideObj, "is_lead"))
        slideDict("speaker_notes") = CStr(sc.Run("getProp", slideObj, "speaker_notes"))
        slideDict("has_table") = CBool(sc.Run("getProp", slideObj, "has_table"))
        slideDict("table_raw") = CStr(sc.Run("getProp", slideObj, "table_raw"))

        ' Body lines array
        Dim bodyLines As New Collection
        Dim blCount As Long
        blCount = CLng(sc.Run("getArrLen", slideObj, "body_lines"))
        Dim j As Long
        For j = 0 To blCount - 1
            bodyLines.Add CStr(sc.Run("getArrItem", slideObj, "body_lines", j))
        Next j
        Set slideDict("body_lines") = bodyLines

        ' Code blocks array
        Dim codeBlocks As New Collection
        Dim cbCount As Long
        cbCount = CLng(sc.Run("getArrLen", slideObj, "code_blocks"))
        For j = 0 To cbCount - 1
            codeBlocks.Add CStr(sc.Run("getArrItem", slideObj, "code_blocks", j))
        Next j
        Set slideDict("code_blocks") = codeBlocks

        ' Images array
        Dim imgArr As New Collection
        Dim imgCount As Long
        imgCount = CLng(sc.Run("getArrLen", slideObj, "images"))
        For j = 0 To imgCount - 1
            Dim imgItem As Object
            Set imgItem = sc.Run("getArrItem", slideObj, "images", j)
            Dim imgDict As Object
            Set imgDict = CreateObject("Scripting.Dictionary")
            imgDict("alt") = CStr(sc.Run("getProp", imgItem, "alt"))
            imgDict("path") = CStr(sc.Run("getProp", imgItem, "path"))
            imgArr.Add imgDict
        Next j
        Set slideDict("images") = imgArr

        result.Add slideDict
    Next i

    Set ParseSlidesJSON = result
End Function


' ---------------------------------------------------------------------------
' Utility Functions
' ---------------------------------------------------------------------------

Private Function SelectJSONFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Slides JSON File"
        .Filters.Clear
        .Filters.Add "JSON Files", "*.json"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectJSONFile = .SelectedItems(1)
        Else
            SelectJSONFile = ""
        End If
    End With
End Function


Private Function ReadFileContent(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        ReadFileContent = ""
        Exit Function
    End If

    Dim stream As Object
    Set stream = fso.OpenTextFile(filePath, 1, False, -1)  ' ForReading, Unicode
    ReadFileContent = stream.ReadAll
    stream.Close
End Function


Private Function GetStringValue(dict As Object, key As String) As String
    On Error Resume Next
    GetStringValue = CStr(dict(key))
    If Err.Number <> 0 Then GetStringValue = ""
    On Error GoTo 0
End Function


Private Function GetArrayValue(dict As Object, key As String) As Object
    On Error Resume Next
    Set GetArrayValue = dict(key)
    If Err.Number <> 0 Then Set GetArrayValue = Nothing
    On Error GoTo 0
End Function


Private Function IsLeadSlide(slideData As Object) As Boolean
    On Error Resume Next
    IsLeadSlide = CBool(slideData("is_lead"))
    If Err.Number <> 0 Then IsLeadSlide = False
    On Error GoTo 0
End Function


Private Function HasSlideImages(slideData As Object) As Boolean
    Dim imgs As Object
    Set imgs = GetArrayValue(slideData, "images")
    If imgs Is Nothing Then
        HasSlideImages = False
    Else
        HasSlideImages = (imgs.Count > 0)
    End If
End Function


Private Function CountBodyLines(slideData As Object) As Long
    Dim lines As Object
    Set lines = GetArrayValue(slideData, "body_lines")
    If lines Is Nothing Then
        CountBodyLines = 0
    Else
        CountBodyLines = lines.Count
    End If
End Function


Private Function JoinBodyLines(slideData As Object) As String
    Dim lines As Object
    Set lines = GetArrayValue(slideData, "body_lines")
    If lines Is Nothing Then
        JoinBodyLines = ""
        Exit Function
    End If

    Dim result As String
    Dim i As Long
    For i = 1 To lines.Count
        If i > 1 Then result = result & vbCrLf
        result = result & CStr(lines(i))
    Next i
    JoinBodyLines = result
End Function


Private Function FindBodyPlaceholder(sld As Slide) As Shape
    Set FindBodyPlaceholder = Nothing

    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Or _
               shp.PlaceholderFormat.Type = ppPlaceholderObject Then
                Set FindBodyPlaceholder = shp
                Exit Function
            End If
        End If
    Next shp
End Function


Private Function GetLowestShapeBottom(sld As Slide) As Single
    Dim lowest As Single
    lowest = 0

    Dim shp As Shape
    For Each shp In sld.Shapes
        Dim bottom As Single
        bottom = shp.Top + shp.Height
        If bottom > lowest Then lowest = bottom
    Next shp

    GetLowestShapeBottom = lowest
End Function


Private Function GetParentFolder(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(filePath)
End Function


Private Function GetFileName(filePath As String) As String
    ' Extract filename from a path that may use / or \
    Dim parts() As String
    If InStr(filePath, "/") > 0 Then
        parts = Split(filePath, "/")
    Else
        parts = Split(filePath, "\")
    End If
    GetFileName = parts(UBound(parts))
End Function


Private Function CleanMarkdown(text As String) As String
    ' Remove common markdown formatting from speaker notes
    Dim result As String
    result = text

    ' Remove ### headers — convert to bold-like text
    result = Replace(result, "### ", "")

    ' Remove bold markers
    result = Replace(result, "**", "")

    ' Remove italic markers (single *)
    ' Be careful not to remove bullet points
    Dim lines() As String
    lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = lines(i)
        ' Remove leading "- " for bullets (PowerPoint has its own bullets)
        If Left(Trim(ln), 2) = "- " Then
            lines(i) = Mid(Trim(ln), 3)
        End If
    Next i
    result = Join(lines, vbCrLf)

    ' Remove backticks
    result = Replace(result, "`", "")

    ' Clean up multiple blank lines
    Do While InStr(result, vbCrLf & vbCrLf & vbCrLf) > 0
        result = Replace(result, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop

    CleanMarkdown = Trim(result)
End Function
