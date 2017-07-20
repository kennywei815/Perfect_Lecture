Attribute VB_Name = "Perfect_Lecturer"
' 2017(c) Perfect Lecture
' Developer: Cheng-Kuan Wei
' URL: https://github.com/kennywei815/Perfect_Lecture
'
' Licensed to the Apache Software Foundation (ASF) under one
' or more contributor license agreements.  See the NOTICE file
' distributed with this work for additional information
' regarding copyright ownership.  The ASF licenses this file
' to you under the Apache License, Version 2.0 (the
' "License"); you may not use this file except in compliance
' with the License.  You may obtain a copy of the License at
'
'   http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing,
' software distributed under the License is distributed on an
' "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
' KIND, either express or implied.  See the License for the
' specific language governing permissions and limitations
' under the License.




'*****************************************************************************
'                        Perfect_Lecturer Module
'              "Compile with Perfect Lecture" main function
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = PowerPoint


Private VBComponents
Private Sep As String
    
Private fileDir As String
Private fileName As String
Private filePath As String
    
Private ProgramDir As String
    


'==============================================================================================================================================
'                                  Common_Path_Vars & Common_Doc_Vars Functions
'                                        Set up commonly used variables
'==============================================================================================================================================
Private Sub Common_Path_Vars()
    
    'Sep = Application.PathSeparator
    Sep = "\"

#If PLATFORM = PowerPoint Then
    fileDir = ActivePresentation.path & Sep
    fileName = ActivePresentation.name
    filePath = ActivePresentation.FullName
    
    Set VBComponents = ActivePresentation.VBProject.VBComponents
            
#ElseIf PLATFORM = Word Then
    fileDir = ActiveDocument.path & Sep
    fileName = ActiveDocument.name
    filePath = ActiveDocument.FullName
    
    Set VBComponents = ActiveDocument.VBProject.VBComponents
            
#ElseIf PLATFORM = Excel Then
    fileDir = ActiveWorkbook.path & Sep
    fileName = ActiveWorkbook.name
    filePath = ActiveWorkbook.FullName
    
    Set VBComponents = ActiveWorkbook.VBProject.VBComponents
    
#End If

    ProgramDir = Environ("APPDATA") & "\Microsoft\AddIns\Perfect_Lecture\"
    
    ChDir fileDir

End Sub

    
Private Sub Common_Doc_Vars()

#If PLATFORM = PowerPoint Then
    Set sld = ActiveWindow.View.Slide
    Set osld = ActiveWindow.Selection.SlideRange(1)
            
#ElseIf PLATFORM = Word Then
    Set sld = ActiveDocument
    Set osld = ActiveDocument
            
#ElseIf PLATFORM = Excel Then
    Set sld = ActiveSheet
    Set osld = ActiveSheet
            
#End If

    Set AllShapes = sld.Shapes

End Sub

'==============================================================================================================================================
'                                            Main Function
'                              "Compile with Perfect Lecture" main function
' Procedure:
'     Step1: 存成要輸出 PPT
'     Step2: 取出備忘稿，存成Precision Lecture Script
'     Step3: 根據Precision Lecture Script，產生post_process VBA檔做out.ppt的後處理：如方程式演變的連續頁面
'     Step4: 存成要輸出 PDF檔
'     Step5: 根據Precision Lecture Script，產生語音合成檔、字幕檔，並將PDF轉成影片、含語音合成口白的輸出PPT
'==============================================================================================================================================
Sub SaveAsMP4()

#If PLATFORM = PowerPoint Then
    
    '==============================================================================================================================================
    ' Step1: 存成要輸出 PPT
    '==============================================================================================================================================
    
    If ActivePresentation.path = "" Then ' therefore new unsaved document
        MsgBox "Please save this document before compiling!"
        Exit Sub
    End If

    With ActivePresentation
        .Save
    End With

    '==============================================================================================================================================
    ' Step2: 設定路徑變數 & 備份原始檔
    '==============================================================================================================================================
    
    Common_Path_Vars
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim tmpSourceFileName As String
    tmpSourceFileName = filePath & ".source"
    
    Dim OutputFileName As String
    OutputFileName = "out_" & fileName
    
    Dim OutputFilePath As String
    OutputFilePath = fileDir & OutputFileName
    
    Dim FilePrefix As String
    FilePrefix = "out"
    'FilePrefix = fileName 'TODO: 用目前檔名當FilePrefix
    
    Dim TempDir As String
    TempDir = "C:\Temp\"
    
    ' If TempDir doesn't exist, create one
    If Dir(TempDir, vbDirectory) = "" Then
       MkDir TempDir
    End If
    

    With Application.ActivePresentation
        fs.copyFile filePath, tmpSourceFileName
    End With
    
    ' Clean up
    CleanUp TempDir, FilePrefix
    
    ' Restore working directory
    ChDir fileDir
    
    
    '==============================================================================================================================================
    ' Step3.1: 取出備忘稿，存成Precision Lecture Script
    '==============================================================================================================================================
    Dim scriptFile As String
    scriptFile = TempDir & FilePrefix & ".script.xml"
    
    SaveNotesToScript (scriptFile)
    
    
    '==============================================================================================================================================
    ' Step3.2: 根據Precision Lecture Script，產生post_process VBA檔做out.ppt的後處理：如方程式演變的連續頁面
    '==============================================================================================================================================
    Dim Command As String
    ''RELEASE
    'Command = "cmd /C python " & ProgramDir & "genPostProcessVBA.py " & TempDir & FilePrefix
    'RunCmd Command, True, vbNormalNoFocus
    
    'DEBUG
    'Command = "cmd /K python " & ProgramDir & "genPostProcessVBA.py " & TempDir & FilePrefix
    Command = "cmd /C python " & ProgramDir & "genPostProcessVBA.py " & TempDir & FilePrefix
    RunCmd Command, True, vbMaximizedFocus
    
   
    Dim Script As String
    Script = TempDir & "post_process.iscript"
    
    ExecScript Script


    '==============================================================================================================================================
    ' Step4.1: 取出備忘稿，存成Precision Lecture Script
    '==============================================================================================================================================
    scriptFile = TempDir & FilePrefix & ".script.xml"
    
    SaveNotesToScript (scriptFile)

    
    '==============================================================================================================================================
    ' Step4.2: 存成要輸出 PDF檔
    '==============================================================================================================================================
    With Application.ActivePresentation
        .SaveAs TempDir & FilePrefix, ppSaveAsPDF
    End With
    
    
    '==============================================================================================================================================
    ' Step4.3: 根據Precision Lecture Script，產生語音合成檔、字幕檔，並將PDF轉成影片、含語音合成口白的輸出PPT
    '==============================================================================================================================================
    
    With ActivePresentation.PageSetup
        Dim SizeSpec As String
        SizeSpec = (.SlideWidth * 2) & "x" & (.SlideHeight * 2)
    End With
    
    ''RELEASE
    'Command = "cmd /C python " & ProgramDir & "pdf2mp4.py " & TempDir & FilePrefix
    ''Shell command, vbNormalNoFocus
    'RunCmd Command, True, vbNormalNoFocus
    
    'DEBUG
    'Command = "cmd /K python " & ProgramDir & "pdf2mp4.py " & TempDir & FilePrefix
    'Command = "cmd /C python " & ProgramDir & "pdf2mp4.py " & TempDir & FilePrefix & " " & SizeSpec
    Command = "cmd /C python " & ProgramDir & "pdf2mp4_size_spec.py " & TempDir & FilePrefix & " " & SizeSpec
    RunCmd Command, True, vbMaximizedFocus
    
    
    
    ExecScript Script
    
    ' Restore working directory
    ChDir fileDir
    With Application.ActivePresentation
        .SaveAs OutputFilePath  'Save to an output file
        
        'Save as videos
        .SaveAs OutputFilePath, ppSaveAsWMV
        .SaveAs OutputFilePath, ppSaveAsMP4
        
        
        If fs.FileExists(filePath) Then
            fs.DeleteFile filePath
        End If
        fs.moveFile tmpSourceFileName, filePath
        'fs.copyFile TempDir & FilePrefix & ".mp4", fileDir & fileName & ".mp4"  '給ffmpeg轉檔搭配使用
    End With
    
    ' Clean up
    CleanUp TempDir, FilePrefix
    
    
    Presentations.Open fileName:=filePath
    
    'Application.Presentations(OutputFileName).Close
    
    '==============================================================================================================================================
    ' Step5: Restore working directory
    '==============================================================================================================================================
    ChDir fileDir
    
#End If
    
    
End Sub


'==============================================================================================================================================
'                                            Helper Functions
'==============================================================================================================================================


Sub CleanUp(TempDir As String, FilePrefix As String)
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    tmpFilePath = FilePrefix & "*.*"
    
    If Dir(TempDir & tmpFilePath) <> Empty Then
        fs.DeleteFile TempDir & tmpFilePath
    End If
End Sub


Sub SaveNotesToScript(scriptFile As String)

    Dim Script As String
    Script = "<?xml version=""1.1"" encoding=""UTF-8""?>" & vbCrLf _
           & "<plscript>" & vbCrLf
           
    For i = 1 To ActivePresentation.Slides.Count
        txt = ActivePresentation.Slides.Range(i).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text
        
        txt = Replace(txt, "&", "&amp;")
        txt = Replace(txt, "<", "&lt;")
        txt = Replace(txt, ">", "&gt;")
        txt = Replace(txt, "'", "&apos;")
        txt = Replace(txt, """", "&quot;")
        
        '[TODO]: vulnerable
        txt = Replace(txt, "&lt;script&gt;", "<script>")
        txt = Replace(txt, "&lt;/script&gt;", "</script>")
        txt = Replace(txt, "&lt;!--", "<!--")
        txt = Replace(txt, "--&gt;", "-->")
        
        'Script = Script & "<page index=""" & i & """>" & vbCrLf _
        '                & txt & vbCrLf _
        '                & "</page>" & vbCrLf

        Script = Script & "<page index=""" & i & """>" _
                        & txt & vbCrLf _
                        & "</page>" & vbCrLf
    Next
    Script = Script & "</plscript>"
    
    WriteToFile_UTF8 Script, scriptFile
End Sub


Public Sub ExecScript(ScriptPath As String)
    Dim code As String
    ReadFromFile_UTF8 code, ScriptPath
    
    Dim table As Collection
    ReadCsv code, table
    
    ''DEBUG
    'For i = 1 To table.Count
    '
    '    For j = 1 To table(i).Count
    '        MsgBox "(" & i & "," & j & ")=" & table(i)(j)
    '    Next j
    '
    'Next i
    
    For i = 1 To table.Count
        'Dim cmd As Collection
        'Set cmd = table(i)
        
        ''DEBUG
        'For j = 1 To table(i).Count
        '    MsgBox "(" & i & "," & j & ")=" & table(i)(j)
        'Next j
        
        If table(i).Count >= 1 Then
            'MsgBox "table(i)(1)=" & QuotedStr(table(i).Item(1)) 'DEBUG
            
            Dim pageIdx As Integer
            Dim shapeName As String
            Dim text As String
            Dim Find As String, Replacement As String
            Dim PointerType As String, R As Single, G As Single, B As Single, PosX As Single, PosY As Single, Width As Single, Height As Single, Rotation As Single
            Dim Track As String
            
            If (table(i).Item(1) = "edit_equation") Then
                pageIdx = CInt(table(i).Item(2))
                shapeName = table(i).Item(3)
                Find = table(i).Item(4)
                Replacement = table(i).Item(5)
                edit_equation pageIdx, shapeName, Find, Replacement
                
            ElseIf (table(i).Item(1) = "duplicate_page") Then
                pageIdx = CInt(table(i).Item(2))
                duplicate_page pageIdx
                
            ElseIf (table(i).Item(1) = "writeNotePage") Then
                pageIdx = CInt(table(i).Item(2))
                text = table(i).Item(3)
                writeNotePage pageIdx, text
                
            ElseIf (table(i).Item(1) = "addNewLineToNotePage") Then
                pageIdx = CInt(table(i).Item(2))
                text = table(i).Item(3)
                addNewLineToNotePage pageIdx, text
            
            ElseIf (table(i).Item(1) = "addPointer") Then
                pageIdx = CInt(table(i).Item(2))
                PointerType = table(i).Item(3)
                R = CInt(table(i).Item(4))
                G = CInt(table(i).Item(5))
                B = CInt(table(i).Item(6))
                PosX = CDbl(table(i).Item(7))
                PosY = CDbl(table(i).Item(8))
                Width = CDbl(table(i).Item(9))
                Height = CDbl(table(i).Item(10))
                Rotation = CDbl(table(i).Item(11))
                
                addPointer pageIdx, PointerType, R, G, B, PosX, PosY, Width, Height, Rotation
                
            ElseIf (table(i).Item(1) = "InsertAudio") Then
                pageIdx = CInt(table(i).Item(2))
                Track = table(i).Item(3)
                InsertAudio pageIdx, Track
            
            End If
        End If
    Next i
End Sub


Sub addPointer(pageIdx As Integer, PointerType As String, R As Single, G As Single, B As Single, PosX As Single, PosY As Single, Width As Single, Height As Single, Rotation As Single)

    ActiveWindow.View.GotoSlide pageIdx

    Common_Doc_Vars

    Const PI = 3.14159265358979

    With ActivePresentation.PageSetup
        PosX = PosX * .SlideWidth
        PosY = PosY * .SlideHeight
    End With

    
    Dim newShape As Shape
    
    If PointerType = "arrow" Then
        
        PosX = PosX - 0.5 * Width + Cos(Rotation * PI / 180) * 0.5 * Width
        PosY = PosY - 0.5 * Height + Sin(Rotation * PI / 180) * 0.5 * Width
    
        Set newShape = AllShapes.AddShape(msoShapeLeftArrow, PosX, PosY, Width, Height)
        
    ElseIf PointerType = "circle" Then
        
        PosX = PosX - 0.5 * Width
        PosY = PosY - 0.5 * Height
        
        Set newShape = AllShapes.AddShape(msoShapeOval, PosX, PosY, Width, Height)
        
    End If
    
    
    With newShape
        'TODO: 設定顏色
        .Fill.ForeColor.RGB = RGB(R, G, B)
        
        'TODO: 旋轉45度或任意角度
        '正值代表順時針旋轉；負值代表逆時針旋轉。
        '若要設定立體圖案繞 x 軸或 y 軸旋轉，請使用RotationX屬性或RotationY屬性ThreeDFormat物件。
        .Rotation = Rotation
        
        '設定框線
        .Line.Weight = 1.5
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
End Sub


Sub edit_equation(pageIdx As Integer, shapeName As String, Find As String, Replacement As String)
    ActiveWindow.View.GotoSlide pageIdx
    ReplaceLaTeX_Batch shapeName, Find, Replacement
End Sub


Sub duplicate_page(pageIdx As Integer)
    ActiveWindow.View.GotoSlide pageIdx
    ActiveWindow.View.Slide.Duplicate
End Sub


Sub writeNotePage(pageIdx As Integer, text As String)
    ActivePresentation.Slides.Range(pageIdx).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = text
End Sub


Sub addNewLineToNotePage(pageIdx As Integer, text As String)
    ActivePresentation.Slides.Range(pageIdx).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text = ActivePresentation.Slides.Range(pageIdx).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text & vbCrLf & text
End Sub


Sub InsertAudio(pageIdx As Integer, Track As String)
    ActiveWindow.View.GotoSlide pageIdx
    
    Common_Doc_Vars

    Dim oShp As Shape
    Dim oEffect As Effect
    'Add the audio shape, and embed it into output document
    Set oShp = sld.Shapes.AddMediaObject2(Track, False, True, 10, 10)
    'Set audio to play automatically
    Set oEffect = sld.TimeLine.MainSequence.AddEffect(oShp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)
    oEffect.MoveTo 1
    'Hide during slide show
    With oEffect
        .EffectInformation.PlaySettings.HideWhileNotPlaying = True
    End With
End Sub

