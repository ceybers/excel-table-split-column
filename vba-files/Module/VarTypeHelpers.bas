Attribute VB_Name = "VarTypeHelpers"
'@Folder("Helpers")
Option Explicit

Public Function VarTypeToString(ByVal v As Long)
    Dim Result As String
    Dim IsArray As Boolean
    Dim ConstantNames As Variant
    ConstantNames = Array("vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble", "vbCurrency", "vbDate", _
        "vbString", "vbObject", "vbError", "vbBoolean", "vbVariant", "vbDataObject", "vbDecvimal", "vbByte", _
        "undefined", "undefined", "vbLongLong")
    
    If v > 8192 Then
        IsArray = True
        v = v - 8192
    End If
    
    If v = 36 Then
        Result = "vbUserDefinedType"
    ElseIf v >= 0 And v <= 20 Then
        Result = ConstantNames(v)
    Else
        Result = "undefined"
    End If
    
    If IsArray Then
        Result = Result & " (Array)"
    End If
    
    VarTypeToString = Result
End Function
