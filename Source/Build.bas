Attribute VB_Name = "Build"
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
'                              Build Module
'  help developers build Office add-ins and export and reload modules(*.bas)
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = PowerPoint
Private Const PLATFORM_DEF As String = "#Const PLATFORM = PowerPoint"


Public WorkDir As String
Public VBComponents
Public Sep As String
    
Public fileDir As String
Public fileName As String
Public fileFullName As String
    


'==============================================================================================================================================
'                                            Common_Vars Function
'                                        Set up commonly used variables
'==============================================================================================================================================
Sub Common_Vars()
    
    'Sep = Application.PathSeparator
    Sep = "\"

#If PLATFORM = PowerPoint Then
    WorkDir = ActivePresentation.path
    
    fileDir = ActivePresentation.path & Sep
    fileName = ActivePresentation.name
    fileFullName = ActivePresentation.FullName
    
    Set VBComponents = ActivePresentation.VBProject.VBComponents
            
#ElseIf PLATFORM = Word Then
    WorkDir = ActiveDocument.path
    
    fileDir = ActiveDocument.path & Sep
    fileName = ActiveDocument.name
    fileFullName = ActiveDocument.FullName
    
    Set VBComponents = ActiveDocument.VBProject.VBComponents
            
#ElseIf PLATFORM = Excel Then
    WorkDir = ActiveWorkbook.path
    
    fileDir = ActiveWorkbook.path & Sep
    fileName = ActiveWorkbook.name
    fileFullName = ActiveWorkbook.FullName
    
    Set VBComponents = ActiveWorkbook.VBProject.VBComponents
    
#End If
    
    ChDir WorkDir

End Sub


'==============================================================================================================================================
'                                            Build Function
'                                    help developers build Office add-ins
'==============================================================================================================================================
Sub Build()

    Common_Vars
    
    Set fs = CreateObject("Scripting.FileSystemObject")

    fn = Left(fileName, Len(fileName) - 5)  'Check if suffix contains 4 characters
    fn_addin = "Perfect_Lecture"
    
    f_path = WorkDir & Sep & fn
    f_path_addin = WorkDir & Sep & fn_addin
    
    'TODO_NOW
    addinDir = WorkDir & Sep & "..\Install\AppData\Roaming\Microsoft\AddIns" & Sep
    wordStartDir = WorkDir & Sep & "..\Install\AppData\Roaming\Microsoft\Word\STARTUP" & Sep
    excelStartDir = WorkDir & Sep & "..\Install\AppData\Roaming\Microsoft\Excel\XLSTART" & Sep
    
    
    'Copy Python files to addinDir\Perfect_Lecture
    pyFilesDir = addinDir & Sep & "Perfect_Lecture" & Sep
    pyFiles_out = pyFilesDir & Sep & "*.py"
    
    If fs.FileExists(pyFiles_out) Then
        fs.DeleteFile pyFiles_out
    End If
    fs.copyFile WorkDir & Sep & "*.py", pyFilesDir
    

#If PLATFORM = PowerPoint Then
    With ActivePresentation
        .Save
        .SaveAs f_path, ppSaveAsOpenXMLAddin
        '.SaveAs f_path, ppSaveAsAddin
        .SaveAs f_path, ppSaveAsOpenXMLPresentationMacroEnabled
    End With
    
    If fs.FileExists(f_path_addin & ".ppam") Then
        fs.DeleteFile f_path_addin & ".ppam"
    End If
    If fs.FileExists(addinDir & fn_addin & ".ppam") Then
        fs.DeleteFile addinDir & fn_addin & ".ppam"
    End If
    
    fs.moveFile f_path & ".ppam", f_path_addin & ".ppam"
    fs.moveFile f_path_addin & ".ppam", addinDir
            
#ElseIf PLATFORM = Word Then
    With ActiveDocument
        .Save
        .SaveAs f_path, wdFormatXMLTemplateMacroEnabled
        '.SaveAs f_path, wdFormatTemplate
        .SaveAs f_path, wdFormatXMLDocumentMacroEnabled
    End With
    
    If fs.FileExists(f_path_addin & ".dotm") Then
        fs.DeleteFile f_path_addin & ".dotm"
    End If
    If fs.FileExists(addinDir & fn_addin & ".dotm") Then
        fs.DeleteFile addinDir & fn_addin & ".dotm"
    End If
    If fs.FileExists(wordStartDir & fn_addin & ".dotm") Then
        fs.DeleteFile wordStartDir & fn_addin & ".dotm"
    End If
    
    fs.moveFile f_path & ".dotm", f_path_addin & ".dotm"
    fs.copyFile f_path_addin & ".dotm", wordStartDir
    fs.moveFile f_path_addin & ".dotm", addinDir
            
#ElseIf PLATFORM = Excel Then
    With ActiveWorkbook
        .Save
        .SaveAs f_path, xlOpenXMLAddIn
        '.SaveAs f_path, xlAddIn
        .SaveAs f_path, xlOpenXMLWorkbookMacroEnabled
    End With
    
    If fs.FileExists(f_path_addin & ".xlam") Then
        fs.DeleteFile f_path_addin & ".xlam"
    End If
    If fs.FileExists(addinDir & fn_addin & ".xlam") Then
        fs.DeleteFile addinDir & fn_addin & ".xlam"
    End If
    If fs.FileExists(excelStartDir & fn_addin & ".xlam") Then
        fs.DeleteFile excelStartDir & fn_addin & ".xlam"
    End If
    
    fs.moveFile f_path & ".xlam", f_path_addin & ".xlam"
    fs.copyFile f_path_addin & ".xlam", excelStartDir
    fs.moveFile f_path_addin & ".xlam", addinDir
            
#End If


End Sub



'==============================================================================================================================================
'                                            Reload_Modules function
'                                     help developers reload modules from *.bas
'==============================================================================================================================================
Sub Reload_Modules()

    Common_Vars
    
    With VBComponents
    
        .Remove .Item("AutoRun")
        .Remove .Item("Common_Utilities")
        
        .Remove .Item("LaTeX2PNG")
        .Remove .Item("Image2PNG")
        
        .Remove .Item("Main")
        .Remove .Item("Main_Helpers")
        
        .Remove .Item("Perfect_Lecturer")
        .Remove .Item("Tests")
        
        
        .Import WorkDir & Sep & "AutoRun.bas"
        .Import WorkDir & Sep & "Common_Utilities.bas"

        .Import WorkDir & Sep & "LaTeX2PNG.bas"
        .Import WorkDir & Sep & "Image2PNG.bas"
        
        .Import WorkDir & Sep & "Main.bas"
        .Import WorkDir & Sep & "Main_Helpers.bas"
        
        .Import WorkDir & Sep & "Perfect_Lecturer.bas"
        .Import WorkDir & Sep & "Tests.bas"
        
        
        .Item("AutoRun").CodeModule.ReplaceLine 48, PLATFORM_DEF
        
        .Item("Main").CodeModule.ReplaceLine 48, PLATFORM_DEF
        
        .Item("Main_Helpers").CodeModule.ReplaceLine 48, PLATFORM_DEF
        
        .Item("Perfect_Lecturer").CodeModule.ReplaceLine 40, PLATFORM_DEF
        
        .Item("Tests").CodeModule.ReplaceLine 40, PLATFORM_DEF
        
    End With

End Sub



'==============================================================================================================================================
'                                            Export_Modules function
'                                     help developers export modules to *.bas
'==============================================================================================================================================
Sub Export_Modules()

    Common_Vars

    With VBComponents

        .Item("AutoRun").Export WorkDir & Sep & "AutoRun.bas"
        .Item("Common_Utilities").Export WorkDir & Sep & "Common_Utilities.bas"

        .Item("LaTeX2PNG").Export WorkDir & Sep & "LaTeX2PNG.bas"
        .Item("Image2PNG").Export WorkDir & Sep & "Image2PNG.bas"
        
        .Item("Main").Export WorkDir & Sep & "Main.bas"
        .Item("Main_Helpers").Export WorkDir & Sep & "Main_Helpers.bas"
        
        .Item("Build").Export WorkDir & Sep & "Build.bas"
        .Item("Reload_Build_Module").Export WorkDir & Sep & "Reload_Build_Module.bas"
        
        .Item("Perfect_Lecturer").Export WorkDir & Sep & "Perfect_Lecturer.bas"
        .Item("Tests").Export WorkDir & Sep & "Tests.bas"
        
    End With
    
    
End Sub

