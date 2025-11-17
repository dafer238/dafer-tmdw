REM  *****  BASIC  *****
REM CoolProp Wrapper Functions for LibreOffice Calc
REM These functions call the Python backend

Option Explicit

Function CPROP(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, fluid As String) As Variant
    ' Calculate thermodynamic properties using engineering units
    Dim oScript As Object
    Dim oProvider As Object
    Dim oScriptContext As Object
    
    On Error GoTo ErrorHandler
    
    ' Get the script provider
    oScriptContext = ThisComponent
    oProvider = oScriptContext.getScriptProvider()
    
    ' Get the Python script
    oScript = oProvider.getScript("vnd.sun.star.script:CoolPropWrapper_Simple.py$CPROP?language=Python&location=user:uno_packages")
    
    ' Call the Python function
    CPROP = oScript.invoke(Array(output, name1, value1, name2, value2, fluid), Array(), Array())
    Exit Function
    
ErrorHandler:
    CPROP = "ERROR: " & Error$
End Function

Function CPROP_SI(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, fluid As String) As Variant
    ' Calculate thermodynamic properties using SI units
    Dim oScript As Object
    Dim oProvider As Object
    Dim oScriptContext As Object
    
    On Error GoTo ErrorHandler
    
    oScriptContext = ThisComponent
    oProvider = oScriptContext.getScriptProvider()
    oScript = oProvider.getScript("vnd.sun.star.script:CoolPropWrapper_Simple.py$CPROP_SI?language=Python&location=user:uno_packages")
    CPROP_SI = oScript.invoke(Array(output, name1, value1, name2, value2, fluid), Array(), Array())
    Exit Function
    
ErrorHandler:
    CPROP_SI = "ERROR: " & Error$
End Function

Function CPROP_E(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, fluid As String) As Variant
    ' Calculate thermodynamic properties using engineering units (explicit)
    CPROP_E = CPROP(output, name1, value1, name2, value2, fluid)
End Function

Function CPROPHA(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, name3 As String, value3 As Double) As Variant
    ' Calculate humid air properties using engineering units
    Dim oScript As Object
    Dim oProvider As Object
    Dim oScriptContext As Object
    
    On Error GoTo ErrorHandler
    
    oScriptContext = ThisComponent
    oProvider = oScriptContext.getScriptProvider()
    oScript = oProvider.getScript("vnd.sun.star.script:CoolPropWrapper_Simple.py$CPROPHA?language=Python&location=user:uno_packages")
    CPROPHA = oScript.invoke(Array(output, name1, value1, name2, value2, name3, value3), Array(), Array())
    Exit Function
    
ErrorHandler:
    CPROPHA = "ERROR: " & Error$
End Function

Function CPROPHA_SI(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, name3 As String, value3 As Double) As Variant
    ' Calculate humid air properties using SI units
    Dim oScript As Object
    Dim oProvider As Object
    Dim oScriptContext As Object
    
    On Error GoTo ErrorHandler
    
    oScriptContext = ThisComponent
    oProvider = oScriptContext.getScriptProvider()
    oScript = oProvider.getScript("vnd.sun.star.script:CoolPropWrapper_Simple.py$CPROPHA_SI?language=Python&location=user:uno_packages")
    CPROPHA_SI = oScript.invoke(Array(output, name1, value1, name2, value2, name3, value3), Array(), Array())
    Exit Function
    
ErrorHandler:
    CPROPHA_SI = "ERROR: " & Error$
End Function

Function CPROPHA_E(output As String, name1 As String, value1 As Double, name2 As String, value2 As Double, name3 As String, value3 As Double) As Variant
    ' Calculate humid air properties using engineering units (explicit)
    CPROPHA_E = CPROPHA(output, name1, value1, name2, value2, name3, value3)
End Function
