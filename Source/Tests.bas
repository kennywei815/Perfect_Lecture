Attribute VB_Name = "Tests"
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
'                             Tests Module
'                test functions for Perfect Lecture project
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


'==============================================================================================================================================
'                                            Test Functions
'==============================================================================================================================================
Public Sub test_batch()

    '==============================================================================================================================================
    ' Step1: copy page & select new page
    '==============================================================================================================================================
    Dim sld As Slide
    Dim osld As Slide
    Set sld = ActiveWindow.View.Slide
    Set osld = ActiveWindow.Selection.SlideRange(1)
    
    Set newSlide = sld.Duplicate
    newSlide.Select
    
    '==============================================================================================================================================
    ' Step2: update LaTeX display
    '==============================================================================================================================================
    Dim shapeName As String
    
    shapeName = "tex4office_obj23119"
    
    'Dim path As String
    'path = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\\Desktop\\test.tex"
    'EditLaTeX_Batch shapeName, path

    ActiveWindow.View.GotoSlide 2

    ReplaceLaTeX_Batch shapeName, "$y = x^2 + 2x + 1$", "$y = (x + 1)^2$"
    
End Sub


Sub test_pointer()
    Dim pageIdx As Integer, PointerType As String, R As Single, G As Single, B As Single, PosX As Single, PosY As Single, Width As Single, Height As Single, Rotation As Single
    
    pageIdx = 1
    
    'PointerType = "arrow"
    PointerType = "circle"
    
    'red
    R = 255
    G = 0
    B = 0
    
    'orange
    R = 255
    G = 140
    B = 0
    
    'darkorange
    R = 255
    G = 140
    B = 0
    
    'yellow
    R = 255
    G = 255
    B = 0
    
    'green
    R = 0
    G = 255
    B = 0
    
    'cyan
    R = 0
    G = 255
    B = 255
    
    'blue
    R = 0
    G = 0
    B = 255
    
    'purple
    R = 255
    G = 0
    B = 255
    
    'black
    R = 0
    G = 0
    B = 0
    
    'black
    R = 255
    G = 255
    B = 255
    
    PosX = 0.5
    PosY = 0.5
    'PosX = 200
    'PosY = 200
    
    'Width = 50
    'Height = 50
    Width = 5
    Height = 5
    
    Rotation = 45

    addPointer pageIdx, PointerType, R, G, B, PosX, PosY, Width, Height, Rotation
End Sub


Sub test_csv()
    Dim csv As String
    csv = "123"
    csv = "123,456"
    csv = "123,456" & vbNewLine & """789,1001,1002"""
    Dim table As Collection
    
    ReadCsv csv, table
    
    For i = 1 To table.Count
        
        For j = 1 To table(i).Count
            MsgBox "(" & i & "," & j & ")=" & table(i)(j)
        Next j
        
    Next i
    
    WriteCsv csv, table
    MsgBox csv
End Sub


Sub test_script()
    Dim TempDir As String
    TempDir = "C:\Temp\"

    Dim Script As String
    Script = TempDir & "post_process.iscript"
    ExecScript Script
End Sub



Public Sub reload_and_exec_post_process()
    Dim TempDir As String
    TempDir = "C:\Temp\"

    With ActivePresentation.VBProject.VBComponents.Item("post_process").CodeModule
        '.CountOfLines
        '.DeleteLines
        '.AddFromFile
        '.InsertLines
        '.ReplaceLine
        '.Lines
        
        .DeleteLines 1, .CountOfLines
        
        .AddFromFile TempDir & "post_process.bas"
        
    End With
    
    'Call post_processing
End Sub


Sub SampleTest()
    Call InsertAudio(1, "C:\Temp\out-0.wav")
End Sub
