Attribute VB_Name = "Reload_Build_Module"
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
'                         Reload_Build_Module Module
'               help developers reload Build module from Build.bas
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = PowerPoint
Private Const PLATFORM_DEF As String = "#Const PLATFORM = PowerPoint"


Private VBComponents
Private Sep As String
Private fileDir As String


'==============================================================================================================================================
'                                            Common_Vars Function
'                                        Set up commonly used variables
'==============================================================================================================================================
Private Sub Common_Vars()
    
    'Sep = Application.PathSeparator
    Sep = "\"

#If PLATFORM = PowerPoint Then
    fileDir = ActivePresentation.path & Sep
    Set VBComponents = ActivePresentation.VBProject.VBComponents
            
#ElseIf PLATFORM = Word Then
    fileDir = ActiveDocument.path & Sep
    Set VBComponents = ActiveDocument.VBProject.VBComponents
            
#ElseIf PLATFORM = Excel Then
    fileDir = ActiveWorkbook.path & Sep
    Set VBComponents = ActiveWorkbook.VBProject.VBComponents
    
#End If
    
    ' Restore working directory
    ChDir fileDir

End Sub


'==============================================================================================================================================
'                                            Reload Modules from .bas
'==============================================================================================================================================
Sub Reload_Build_Module()

    Common_Vars
    
    Set fs = CreateObject("Scripting.FileSystemObject")

    With VBComponents
    
        .Remove .Item("Build")
        .Import fileDir & "Build.bas"
        
        
        .Item("Build").CodeModule.ReplaceLine 40, PLATFORM_DEF
        .Item("Build").CodeModule.ReplaceLine 41, "Private Const PLATFORM_DEF As String = """ & PLATFORM_DEF & """"
        
    End With

End Sub
