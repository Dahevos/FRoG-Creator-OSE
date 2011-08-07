Attribute VB_Name = "modCodeEdit"
Public Type WMcolors
  bgClr As Long
  frClr As Long
  fntProp As Long
End Type

Public ClrData(19) As WMcolors
Public AFileName As String

Public Sub ResetAllEditVals()

'Save the Default values to the registry
SaveSetting App.EXEName, "EditOptions", "c0a", "0"
SaveSetting App.EXEName, "EditOptions", "c0b", "65535"
SaveSetting App.EXEName, "EditOptions", "c0c", "0"

SaveSetting App.EXEName, "EditOptions", "c1a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c1b", "32768"
SaveSetting App.EXEName, "EditOptions", "c1c", "2"

SaveSetting App.EXEName, "EditOptions", "c2a", "0"
SaveSetting App.EXEName, "EditOptions", "c2b", "16777215"
SaveSetting App.EXEName, "EditOptions", "c2c", "0"

SaveSetting App.EXEName, "EditOptions", "c3a", "0"
SaveSetting App.EXEName, "EditOptions", "c3b", "16777215"
SaveSetting App.EXEName, "EditOptions", "c3c", "0"

SaveSetting App.EXEName, "EditOptions", "c4a", "0"
SaveSetting App.EXEName, "EditOptions", "c4b", "16777152"
SaveSetting App.EXEName, "EditOptions", "c4c", "0"

SaveSetting App.EXEName, "EditOptions", "c5a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c5b", "16711680"
SaveSetting App.EXEName, "EditOptions", "c5c", "1"

SaveSetting App.EXEName, "EditOptions", "c6a", "0"
SaveSetting App.EXEName, "EditOptions", "c6b", "8421504"
SaveSetting App.EXEName, "EditOptions", "c6c", "0"

SaveSetting App.EXEName, "EditOptions", "c7a", "8421504"
SaveSetting App.EXEName, "EditOptions", "c7b", "16777215"
SaveSetting App.EXEName, "EditOptions", "c7c", "0"

SaveSetting App.EXEName, "EditOptions", "c8a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c8b", "0"
SaveSetting App.EXEName, "EditOptions", "c8c", "0"

SaveSetting App.EXEName, "EditOptions", "c9a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c9b", "255"
SaveSetting App.EXEName, "EditOptions", "c9c", "0"

SaveSetting App.EXEName, "EditOptions", "c10a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c10b", "16711680"
SaveSetting App.EXEName, "EditOptions", "c10c", "0"

SaveSetting App.EXEName, "EditOptions", "c11a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c11b", "12583104"
SaveSetting App.EXEName, "EditOptions", "c11c", "0"

SaveSetting App.EXEName, "EditOptions", "c12a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c12b", "128"
SaveSetting App.EXEName, "EditOptions", "c12c", "1"

SaveSetting App.EXEName, "EditOptions", "c13a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c13b", "255"
SaveSetting App.EXEName, "EditOptions", "c13c", "0"

SaveSetting App.EXEName, "EditOptions", "c14a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c14b", "16711680"
SaveSetting App.EXEName, "EditOptions", "c14c", "0"

SaveSetting App.EXEName, "EditOptions", "c15a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c15b", "0"
SaveSetting App.EXEName, "EditOptions", "c15c", "1"

SaveSetting App.EXEName, "EditOptions", "c16a", "16777215"
SaveSetting App.EXEName, "EditOptions", "c16b", "0"
SaveSetting App.EXEName, "EditOptions", "c16c", "0"

SaveSetting App.EXEName, "EditOptions", "c17a", "0"
SaveSetting App.EXEName, "EditOptions", "c17b", "16777215"
SaveSetting App.EXEName, "EditOptions", "c17c", "0"

SaveSetting App.EXEName, "EditOptions", "c18a", "0"
SaveSetting App.EXEName, "EditOptions", "c18b", "8388736"
SaveSetting App.EXEName, "EditOptions", "c18c", "0"

SaveSetting App.EXEName, "EditOptions", "c19a", "0"
SaveSetting App.EXEName, "EditOptions", "c19b", "8388736"
SaveSetting App.EXEName, "EditOptions", "c19c", "0"

SaveSetting App.EXEName, "EditOptions", "Saved", "1"

End Sub
Public Sub GetEditColors()
On Error GoTo EH
Dim I As Integer
'Get the color Values
For I = 0 To 19
ClrData(I).bgClr = CLng(GetSetting(App.EXEName, "EditOptions", "c" & I & "a", "0"))
ClrData(I).frClr = CLng(GetSetting(App.EXEName, "EditOptions", "c" & I & "b", "0"))
ClrData(I).fntProp = CLng(GetSetting(App.EXEName, "EditOptions", "c" & I & "c", "0"))
Next I
Exit Sub
EH:
End Sub
