Option Strict Off
Option Explicit On 
Imports Microsoft.Win32

Module modGeneral
    'Function SetValue(ByRef aValName As String, ByRef aValData As String) As Boolean
    '    Const subKeyName = "Software\TypeTest"
    '    Dim objSubKey As RegistryKey
    '    Dim objParentKey As RegistryKey

    '    objParentKey = Registry.LocalMachine
    '    SetValue = True
    '    Try
    '        objSubKey = objParentKey.OpenSubKey(subKeyName, True)
    '        If objSubKey Is Nothing Then
    '            objSubKey = objParentKey.CreateSubKey(subKeyName)
    '        End If
    '        objSubKey.SetValue(aValName, aValData)
    '    Catch ex As Exception
    '        SetValue = False
    '        MsgBox("SetValue: Writing to registry: " & ex.Message)
    '    End Try
    'End Function

    'Function CreateKey() As Boolean
    '    Const aKey = "TypeTest"
    '    Dim regKey As RegistryKey

    '    CreateKey = True 'Returns true = ok or false = problem
    '    Try
    '        regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
    '        regKey.CreateSubKey(aKey)
    '        regKey.Close()
    '    Catch ex As Exception
    '        MsgBox("CreateKey: Write to registry: " & ex.Message)
    '        CreateKey = False 'Returns true = ok or false = problem
    '    End Try
    'End Function

    Sub CentreForm(ByRef aForm As System.Windows.Forms.Form)
        With aForm
            '.Width = Screen.Width * 0.75 ' Set width of form.
            '.Height = Screen.Height * 0.75   ' Set height of form.
            .Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(.Width)) / 2) ' Center form horizontally.
            .Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(.Height)) / 2) ' Center form vertically.
        End With
    End Sub

    Public Function GetValue(ByRef aValName As String) As String
        'Returns value string or "" if empty or non-existant
        Const subKeyName = "Software\TypeTest"
        Dim objSubKey As RegistryKey
        Dim objParentKey As RegistryKey

        GetValue = ""
        objParentKey = Registry.LocalMachine
        Try
            objSubKey = objParentKey.OpenSubKey(subKeyName, True)
            If Not objSubKey Is Nothing Then
                GetValue = objSubKey.GetValue(aValName)
            End If
        Catch ex As Exception
            MsgBox("GetValue: Reading registry: " & ex.Message)
        End Try
    End Function
End Module