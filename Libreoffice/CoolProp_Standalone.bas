REM  *****  BASIC  *****
REM CoolProp Standalone Module for LibreOffice Calc
REM 
REM INSTALLATION:
REM 1. Open LibreOffice Calc
REM 2. Tools → Macros → Edit Macros
REM 3. In the Macro editor, File → New → Module
REM 4. Delete all content and paste this entire file
REM 5. Save and close
REM 6. Functions will be available in spreadsheet formulas

Option Explicit

' Python executable path - adjust if needed
Const PYTHON_PATH As String = "C:\LibreOfficePortable\App\libreoffice\program\python.exe"
Const SITE_PACKAGES As String = "C:\LibreOfficePortable\App\libreoffice\program\python-core-3.10.18\lib\site-packages"

Private Function RunPythonScript(scriptContent As String) As String
    Dim oFileAccess As Object
    Dim oSimpleFileAccess As Object
    Dim tempDir As String
    Dim scriptPath As String
    Dim outputPath As String
    Dim command As String
    Dim result As String
    Dim oShell As Object
    
    On Error GoTo ErrorHandler
    
    ' Get temp directory
    tempDir = Environ("TEMP")
    scriptPath = tempDir & "\coolprop_calc.py"
    outputPath = tempDir & "\coolprop_result.txt"
    
    ' Write script to file
    Open scriptPath For Output As #1
    Print #1, scriptContent
    Close #1
    
    ' Run Python and capture output
    command = Chr(34) & PYTHON_PATH & Chr(34) & " " & Chr(34) & scriptPath & Chr(34) & " > " & Chr(34) & outputPath & Chr(34) & " 2>&1"
    
    Shell(command, 0, True)
    
    ' Wait for completion
    Wait 200
    
    ' Read result
    On Error Resume Next
    Open outputPath For Input As #1
    result = ""
    Do While Not EOF(1)
        Dim line As String
        Line Input #1, line
        If result <> "" Then result = result & vbLf
        result = result & line
    Loop
    Close #1
    
    RunPythonScript = Trim(result)
    Exit Function
    
ErrorHandler:
    RunPythonScript = "ERROR: " & Err.Description & " (" & Err.Number & ")"
End Function

' ============================================================================
' Main CoolProp Functions
' ============================================================================

Function CPROP(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, fluid As String) As Variant
    Dim script As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    script = "import sys" & vbLf & _
             "sys.path.insert(0, r'" & SITE_PACKAGES & "')" & vbLf & _
             "from CoolProp.CoolProp import PropsSI" & vbLf & _
             "try:" & vbLf & _
             "    val = PropsSI('" & output & "', '" & name1 & "', " & Str(value1) & ", '" & name2 & "', " & Str(value2) & ", '" & fluid & "')" & vbLf & _
             "    print(val)" & vbLf & _
             "except Exception as e:" & vbLf & _
             "    print('ERROR:', str(e))"
    
    result = RunPythonScript(script)
    
    If Left(result, 6) = "ERROR:" Then
        CPROP = result
    ElseIf result = "" Then
        CPROP = "ERROR: No output from Python"
    Else
        CPROP = Val(result)
    End If
    Exit Function
    
ErrorHandler:
    CPROP = "ERROR: " & Err.Description
End Function

Function CPROP_E(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, fluid As String) As Variant
    ' Engineering units version (same as CPROP for now)
    CPROP_E = CPROP(output, name1, value1, name2, value2, fluid)
End Function

Function CPROP_SI(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, fluid As String) As Variant
    ' SI units version
    CPROP_SI = CPROP(output, name1, value1, name2, value2, fluid)
End Function

Function CPROPHA(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, name3 As String, value3 As Double) As Variant
    Dim script As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    script = "import sys" & vbLf & _
             "sys.path.insert(0, r'" & SITE_PACKAGES & "')" & vbLf & _
             "from CoolProp.CoolProp import HAPropsSI" & vbLf & _
             "try:" & vbLf & _
             "    val = HAPropsSI('" & output & "', '" & name1 & "', " & Str(value1) & ", '" & name2 & "', " & Str(value2) & ", '" & name3 & "', " & Str(value3) & ")" & vbLf & _
             "    print(val)" & vbLf & _
             "except Exception as e:" & vbLf & _
             "    print('ERROR:', str(e))"
    
    result = RunPythonScript(script)
    
    If Left(result, 6) = "ERROR:" Then
        CPROPHA = result
    ElseIf result = "" Then
        CPROPHA = "ERROR: No output from Python"
    Else
        CPROPHA = Val(result)
    End If
    Exit Function
    
ErrorHandler:
    CPROPHA = "ERROR: " & Err.Description
End Function

Function CPROPHA_E(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, name3 As String, value3 As Double) As Variant
    ' Engineering units version
    CPROPHA_E = CPROPHA(output, name1, value1, name2, value2, name3, value3)
End Function

Function CPROPHA_SI(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, name3 As String, value3 As Double) As Variant
    ' SI units version
    CPROPHA_SI = CPROPHA(output, name1, value1, name2, value2, name3, value3)
End Function
