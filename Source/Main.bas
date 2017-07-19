Attribute VB_Name = "Main"
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
' This file incorporates work from Jonathan Le Roux and Zvika
' Ben-Haim's IguanaTeX project which is originally released
' under the Creative Commons Attribution 3.0 License; you may
' not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'   https://creativecommons.org/licenses/by/3.0/
'
' Unless required by applicable law or agreed to in writing,
' software distributed under the License is distributed on an
' "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
' KIND, either express or implied.  See the License for the
' specific language governing permissions and limitations
' under the License.




'*****************************************************************************
'                               Main Module
'       "New/Edit LaTeX Display" & "Insert/Change Image" main functions
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = PowerPoint


Const TeX4Office As Integer = 0
Const ImportImage As Integer = 1


Private VBComponents
Private Sep As String
    
Private fileDir As String
Private fileName As String
Private fileFullName As String


' Get current slide or document
Public AllShapes As Shapes

#If PLATFORM = PowerPoint Then
    Public SlideIndex As Long
    Public sld As Slide
    Public osld As Slide
            
#ElseIf PLATFORM = Word Then
    Public sld As Document
    Public osld As Document
            
#ElseIf PLATFORM = Excel Then
    Public sld As Worksheet
    Public osld As Worksheet
            
#End If
    


'==============================================================================================================================================
'                                            Common_Vars Function
'                                        Set up commonly used variables
'==============================================================================================================================================
Private Sub Common_Vars()
    
    'Sep = Application.PathSeparator
    Sep = "\"

#If PLATFORM = PowerPoint Then
    fileDir = ActivePresentation.path & Sep
    fileName = ActivePresentation.name
    fileFullName = ActivePresentation.FullName
    
    Set VBComponents = ActivePresentation.VBProject.VBComponents
            
#ElseIf PLATFORM = Word Then
    fileDir = ActiveDocument.path & Sep
    fileName = ActiveDocument.name
    fileFullName = ActiveDocument.FullName
    
    Set VBComponents = ActiveDocument.VBProject.VBComponents
            
#ElseIf PLATFORM = Excel Then
    fileDir = ActiveWorkbook.path & Sep
    fileName = ActiveWorkbook.name
    fileFullName = ActiveWorkbook.FullName
    
    Set VBComponents = ActiveWorkbook.VBProject.VBComponents
    
#End If
    
    
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
    
    ' Restore working directory
    ChDir fileDir

End Sub


'==============================================================================================================================================
'                                Wrappers for NewEditPicture Main Function
'                       "New/Edit LaTeX Display" & "Insert/Change Image" main functions
'==============================================================================================================================================

Public Sub EditLaTeX()
    Common_Vars
    
    NewEditPicture TeX4Office
    
    ' Restore working directory
    ChDir fileDir
    
End Sub


Public Sub InsertPicture()
    Common_Vars

    NewEditPicture ImportImage
    
    ' Restore working directory
    ChDir fileDir
    
End Sub


Public Sub ReplaceLaTeX_Batch(shapeName As String, Find As String, Replacement As String)
    '[TODO]: write the content of TeXFile into this shape.AlternativeText
    Common_Vars


    '==============================================================================================================================================
    ' Step 1: Select the shape called ShapeName
    '==============================================================================================================================================
    
    '[TODO]: check ShapeName exists???
    
    Dim oldShape As Shape
    Dim code As String
    
    
    If IsInShapes(AllShapes, shapeName) Then
        Set oldShape = AllShapes(shapeName)
    Else
        '[TODO]: if not exist, report error!
        
        Dim PosX As Single
        Dim PosY As Single
        Dim defaultPictureFile As String
        Dim defaultTeXFile As String
        
        PosX = 200
        PosY = 200
        defaultPictureFile = Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\default.png" 'TODO_NOW
        defaultTeXFile = Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\default.tex"
        
        ReadFromFile_UTF8 code, defaultTeXFile
    
        Call AddTagsToShape(oldShape, code, shapeName)
        
        
        Set oldShape = AddDisplayShape(TeX4Office, defaultPictureFile, PosX, PosY)
        
        With oldShape
            ' Original size
            .ScaleHeight 1#, msoTrue
            .ScaleWidth 1#, msoTrue
            
            .LockAspectRatio = msoFalse
            
            ' Apply scaling factors
            default_screen_dpi = 96
            OutputDpi = 600    '[TODO]: read OutputDpi from TeX4Office Editor's config.json
            OldDpi = OutputDpi '[TODO]: read OldDpi    from shape.AlternativeText ???
            
            MagicScalingFactorPNG = default_screen_dpi / OutputDpi
            
            '[TODO]: read DPI       from TeX4Office Editor's config.json
            '[TODO]: read PointSize from TeX4Office Editor's config.json
        
            PointSize = 10
            tScale = PointSize / 10 * MagicScalingFactorPNG  ' 1/10 is for the default LaTeX point size (10 pt)
            
            .ScaleHeight tScale, msoTrue
            .ScaleWidth tScale, msoTrue
            
            .LockAspectRatio = msoTrue
        End With
    End If

#If PLATFORM = PowerPoint Then
    ActiveWindow.View.GotoSlide sld.SlideIndex  'DEBUG
#End If

    oldShape.Select
    
    
    '==============================================================================================================================================
    ' Step 2: Read code from TeX file, and save code & other setting to newShape.AlternativeText
    '==============================================================================================================================================
    code = oldShape.AlternativeText
    
    code = Replace(code, Find, Replacement)
        
    Call AddTagsToShape(oldShape, code, shapeName)
    
    ' Restore working directory
    ChDir fileDir
    
    '==============================================================================================================================================
    ' Step 3: Update the LaTeX display
    '==============================================================================================================================================
    NewEditPicture TeX4Office, True
    
    ' Restore working directory
    ChDir fileDir
    
End Sub


Public Sub EditLaTeX_Batch(shapeName As String, TeXFilePath As String)
    '[TODO]: write the content of TeXFile into this shape.AlternativeText
    Common_Vars


    '==============================================================================================================================================
    ' Step 1: Select the shape called ShapeName
    '==============================================================================================================================================
    
    '[TODO]: check ShapeName exists???
    
    Dim oldShape As Shape
    
    
    If IsInShapes(AllShapes, shapeName) Then
        Set oldShape = AllShapes(shapeName)
    Else
        '[TODO]: if not exist, report error!
        
        Dim PosX As Single
        Dim PosY As Single
        Dim defaultPictureFile As String
        Dim defaultTeXFile As String
        Dim code As String
        
        PosX = 200
        PosY = 200
        defaultPictureFile = Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\default.png" 'TODO_NOW
        defaultTeXFile = Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\default.tex"
        
        ReadFromFile_UTF8 code, defaultTeXFile
    
        Call AddTagsToShape(oldShape, code, shapeName)
        
        
        Set oldShape = AddDisplayShape(TeX4Office, defaultPictureFile, PosX, PosY)
        
        With oldShape
            ' Original size
            .ScaleHeight 1#, msoTrue
            .ScaleWidth 1#, msoTrue
            
            .LockAspectRatio = msoFalse
            
            ' Apply scaling factors
            default_screen_dpi = 96
            OutputDpi = 600    '[TODO]: read OutputDpi from TeX4Office Editor's config.json
            OldDpi = OutputDpi '[TODO]: read OldDpi    from shape.AlternativeText ???
            
            MagicScalingFactorPNG = default_screen_dpi / OutputDpi
            
            '[TODO]: read DPI       from TeX4Office Editor's config.json
            '[TODO]: read PointSize from TeX4Office Editor's config.json
        
            PointSize = 10
            tScale = PointSize / 10 * MagicScalingFactorPNG  ' 1/10 is for the default LaTeX point size (10 pt)
            
            .ScaleHeight tScale, msoTrue
            .ScaleWidth tScale, msoTrue
            
            .LockAspectRatio = msoTrue
        End With
    End If

#If PLATFORM = PowerPoint Then
    ActiveWindow.View.GotoSlide sld.SlideIndex  'DEBUG
#End If

    oldShape.Select
    
    
    '==============================================================================================================================================
    ' Step 2: Read code from TeX file, and save code & other setting to newShape.AlternativeText
    '==============================================================================================================================================
    ReadFromFile_UTF8 code, TeXFilePath
    
    Call AddTagsToShape(oldShape, code, shapeName)
    
    ' Restore working directory
    ChDir fileDir
    
    '==============================================================================================================================================
    ' Step 3: Update the LaTeX display
    '==============================================================================================================================================
    NewEditPicture TeX4Office, True
    
    ' Restore working directory
    ChDir fileDir
    
End Sub


'==============================================================================================================================================
'                                   NewEditPicture Main Function
'                       "New/Edit LaTeX Display" & "Insert/Change Image" main functions
' Procedure:
'   Step 1: Get information from old shape (if existed)
'   Step 2: Run LaTeX Editor or other PNG generators
'   Step 3: Get scaling factors from oldShape or according to current DPI and the font size
'   Step 4: Insert PNG image and rescale it
'   Step 5: Save code & other setting to newShape.AlternativeText, and set newShape.Name = FilePrefix
'   Step 6: Copy animation settings, grouping information, and formatting from old image, then delete it
'==============================================================================================================================================
Public Sub NewEditPicture(PROGRAM As Integer, Optional BatchMode As Boolean = False)
    '[DONE]: implement Batch_Mode
    Common_Vars

    
    '[TODO]: implement ßÔ¿…¶W
    '[TODO]: [Konwn Issue] In Excel, we don't know how to get OldGroup & OldGroup.GroupItems. Currently we can't restore group information of oldShape in Excel.
    
    
        
    ' If we are in Edit mode, store parameters of old image
    Dim PosX As Single
    Dim PosY As Single
    Set Sel = ActiveWindow.Selection
    
    Dim oldShape As Shape
    
    Dim s As Shape
    IsInGroup = False
    
    
    '==============================================================================================================================================
    ' Step 1: Get information from old shape (if existed)
    '==============================================================================================================================================

    If selectionIsLaTeXShape() Then
        
        If selectionIsGroup() Then   'Old image is part of a group
            
#If PLATFORM = Excel Then
            Set oldShape = Sel.ShapeRange(1)      'For Excel
#Else
            Set oldShape = Sel.ChildShapeRange(1) 'For PowerPoint & Word
            
#End If
            
            IsInGroup = True
            Dim ShapeNames As New Collection ' gather all shapes to be regrouped later on
            
            ' Store the group's animation and Zorder info in a dummy object tmpGroup
            Dim oldGroup As Shape
            Dim tmpGroup As Shape
            
            Set oldGroup = Sel.ShapeRange(1)
            Set tmpGroup = AllShapes.AddShape(msoShapeDiamond, 1, 1, 1, 1)
            
            MoveAnimation oldGroup, tmpGroup
            MatchZOrder oldGroup, tmpGroup
        
            ' Tag all elements in the group with their hierarchy level and their name or group name
            Dim MaxGroupLevel As Long
            Dim Layers As New Collection
            Dim oldShapeSelectionName As String
            Dim GroupNames As New Collection
            Dim SelectionNames As New Collection
            
            oldShapeSelectionName = InvalidName

#If PLATFORM = PowerPoint Then
            For Each s In Sel.ShapeRange.GroupItems
                If s.name <> oldShape.name Then
                    ShapeNames.Add s.name
                End If
            Next
            MaxGroupLevel = RecordGroupHierarchy_and_Ungroup_Fast(ShapeNames, oldShape.name, Layers, SelectionNames)
            
#ElseIf PLATFORM = Word Then
            MaxGroupLevel = RecordGroupHierarchy_and_Ungroup(oldGroup, oldShape.name, oldShapeSelectionName, ShapeNames, Layers, GroupNames)
#End If
            
            oldShape.Select
            
        Else
            Set oldShape = Sel.ShapeRange(1)
        End If
        PosX = oldShape.Left
        PosY = oldShape.Top
    Else
        PosX = 200
        PosY = 200
    End If
    
    
    '==============================================================================================================================================
    ' Step 2: Run LaTeX Editor or other PNG generators
    '==============================================================================================================================================
    Dim TempDir As String
    Dim FilePrefix As String
    Dim code As String
    
If PROGRAM = TeX4Office Then
    LaTeX2PNG_Func oldShape, TempDir, FilePrefix, code, BatchMode
    
ElseIf PROGRAM = ImportImage Then
    Image2PNG_Func oldShape, TempDir, FilePrefix, code

End If
    
    sourceFileName = FilePrefix & ".tex"
    pictureFileName = FilePrefix & ".png"
    tmpFilePath = FilePrefix & ".*"
    
    sourceFilePath = TempDir & FilePrefix & ".tex"
    pictureFilePath = TempDir & FilePrefix & ".png"
    tmpFilePath = TempDir & FilePrefix & ".*"
    
    
    If Dir(pictureFilePath) = Empty Then
        Exit Sub
    End If
    
    
    '==============================================================================================================================================
    ' Step 3: Get scaling factors from oldShape or according to current DPI and the font size
    '==============================================================================================================================================
    Dim tScaleWidth As Single, tScaleHeight As Single
    
    default_screen_dpi = 96
    OutputDpi = 600    '[TODO]: read OutputDpi from TeX4Office Editor's config.json
    OldDpi = OutputDpi '[TODO]: read OldDpi    from shape.AlternativeText ???
    
    MagicScalingFactorPNG = default_screen_dpi / OutputDpi
    
    If selectionIsLaTeXShape() Then
    
        With oldShape
            ' Save current size
            oldShapeHeight = .Height
            oldShapeWidth = .Width
        
            ' Original size
            .ScaleHeight 1#, msoTrue
            .ScaleWidth 1#, msoTrue
            
            ' Calculate relative size
            tScaleHeight = oldShapeHeight / .Height
            tScaleWidth = oldShapeWidth / .Width
                        
            ' Restore sizing
            .ScaleHeight tScaleHeight, msoTrue
            .ScaleWidth tScaleWidth, msoTrue
        End With
    End If
    
    '==============================================================================================================================================
    ' Step 4: Insert PNG image and rescale it
    '==============================================================================================================================================
    Dim newShape As Shape
    
    Set newShape = AddDisplayShape(PROGRAM, TempDir + pictureFileName, PosX, PosY)
    
    'Delete temporary files
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Dir(tmpFilePath) <> Empty Then
        'fs.DeleteFile tmpFilePath
    End If
    
    ' Resize to the true size of the png file and adjust using the manual scaling factors set in Main Settings
    With newShape
        ' Original size
        .ScaleHeight 1#, msoTrue
        .ScaleWidth 1#, msoTrue
        
        .LockAspectRatio = msoFalse
        
        ' Apply scaling factors
        If selectionIsLaTeXShape() Then
            '[TODO]: read current & old DPI
            
            .ScaleHeight tScaleHeight * OldDpi / OutputDpi, msoTrue
            .ScaleWidth tScaleWidth * OldDpi / OutputDpi, msoTrue
        Else
            '[TODO]: read DPI       from TeX4Office Editor's config.json
            '[TODO]: read PointSize from TeX4Office Editor's config.json
            
If PROGRAM = TeX4Office Then
            PointSize = 10
            tScale = PointSize / 10 * MagicScalingFactorPNG  ' 1/10 is for the default LaTeX point size (10 pt)
            
            .ScaleHeight tScale, msoTrue
            .ScaleWidth tScale, msoTrue
End If
            
        End If
        
        .LockAspectRatio = msoTrue
    End With
        
    ' We are editing+resetting size of an old display, we keep rotation
    If selectionIsLaTeXShape() Then
        newShape.Rotation = oldShape.Rotation
        newShape.LockAspectRatio = oldShape.LockAspectRatio ' Unlock aspect ratio if old display had it unlocked
    End If
    
    
    '==============================================================================================================================================
    ' Step 5: Save code & other setting to newShape.AlternativeText, and set newShape.Name = FilePrefix
    '==============================================================================================================================================
    Call AddTagsToShape(newShape, code, FilePrefix)
    
    '==============================================================================================================================================
    ' Step 6: Copy animation settings, grouping information, and formatting from old image, then delete it
    '==============================================================================================================================================
    If selectionIsLaTeXShape() Then

        If IsInGroup Then

            ' Transfer format to new shape
            Dim GroupName As String
            
            GroupName = oldShapeSelectionName
            
            MatchZOrder oldShape, newShape
            
            oldShape.PickUp
            newShape.Apply
            oldShape.Delete
            
            Dim newGroup As Shape

            ' Group all non-modified elements from old group, plus modified element
            ShapeNames.Add newShape.name
            
            Layers.Add 1, newShape.name
            GroupNames.Add GroupName, newShape.name
            SelectionNames.Add newShape.name, newShape.name
                        
            ' Hierarchically re-group elements
            For Level = 1 To MaxGroupLevel
            
'                '====================================================================================== BEGIN DEBUG ======================================================================================
'                MsgBox "Begin regrouping"
'                MsgBox "MaxGroupLevel=" & MaxGroupLevel
'                MsgBox "ShapeNames.Count=" & ShapeNames.Count
'                MsgBox "Layers.Count=" & Layers.Count
'                MsgBox "SelectionNames.Count=" & SelectionNames.Count
'                MsgBox "GroupNames.Count=" & GroupNames.Count
'                For Each n In ShapeNames
'#If PLATFORM <> Word Then
'                    MsgBox "ShapeNames: (Name, Layer, SelectionName) = " & n & ", " & Layers(n) & ", " & SelectionNames(n)
'#Else
'                    MsgBox "ShapeNames: (Name, Layer, GroupName) = " & n & ", " & Layers(n) & ", " & GroupNames(n)
'#End If
'                Next
'                '====================================================================================== END DEBUG ======================================================================================
            
                Dim CurrentLevelShapeNames As New Collection
                ClearCollection CurrentLevelShapeNames
                
                For Each n In ShapeNames
                    Dim ShapeLevel As Integer
                    ShapeLevel = 0
                    
                    On Error Resume Next
                    ShapeLevel = Layers.Item(n)
                    
                    Dim n_ToBeGroup As String
                    
#If PLATFORM <> Word Then
                    n_ToBeGroup = SelectionNames(n)
#Else
                    n_ToBeGroup = n
#End If
                    
                    If ShapeLevel = Level Then
                        If CurrentLevelShapeNames.Count > 0 Then
                            If Not IsInCollection(CurrentLevelShapeNames, n_ToBeGroup) Then
                                CurrentLevelShapeNames.Add n_ToBeGroup
                            End If
                        Else
                            CurrentLevelShapeNames.Add n_ToBeGroup
                        End If
                    End If
                Next
                
'                '====================================================================================== BEGIN DEBUG ======================================================================================
'                MsgBox "CurrentLevelShapeNames.Count=" & CurrentLevelShapeNames.Count
'                For Each n In CurrentLevelShapeNames
'#If PLATFORM <> Word Then
'                    MsgBox "CurrentLevelShapeNames: ShapeNames: (Name, Layer, SelectionName) = " & n & ", " & Layers(n) & ", " & SelectionNames(n)
'#Else
'                    MsgBox "CurrentLevelShapeNames: ShapeNames: (Name, Layer, GroupName) = " & n & ", " & Layers(n) & ", " & GroupNames(n)
'#End If
'                Next
'                '====================================================================================== END DEBUG ======================================================================================
                
                If CurrentLevelShapeNames.Count > 1 Then
                    Set newGroup = sld.Shapes.Range(ToArray(CurrentLevelShapeNames)).Group()
#If PLATFORM = Word Then
                    newGroup.name = GroupNames(CurrentLevelShapeNames(1))
#End If
                    ShapeNames.Add newGroup.name
                    
                    Layers.Add Level + 1, newGroup.name
                    GroupNames.Add newGroup.name, newGroup.name
                    SelectionNames.Add newGroup.name, newGroup.name
                    
                    For Each n_CurrentLevel In CurrentLevelShapeNames
                        ShapeNames.Remove n_CurrentLevel
                    Next
                End If
                
            Next
            
            ' Use temporary group to retrieve the group's original animation and Zorder
            MoveAnimation tmpGroup, newGroup
            MatchZOrder tmpGroup, newGroup
            tmpGroup.Delete
        Else
            MoveAnimation oldShape, newShape
            MatchZOrder oldShape, newShape
            
            oldShape.PickUp
            newShape.Apply
            oldShape.Delete
        End If
    End If
    
    
    '==============================================================================================================================================
    ' Step 7: Restore working directory
    '==============================================================================================================================================
    ChDir fileDir
    
End Sub

