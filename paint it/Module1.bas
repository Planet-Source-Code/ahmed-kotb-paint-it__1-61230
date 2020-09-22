Attribute VB_Name = "Module1"
Public BruchWidth As Integer
Public Status As Boolean
Public msg As String
Public percentage As Single
Public Scrstat As Integer
Public Type score
easy As String
easyscr As Single
med As String
medscr As Single
hard As String
hardscr As Single
End Type
Public per As score
Sub Savescore()
Open App.Path & "\data.dat" For Random As #1
Put #1, 1, per
Close #1
End Sub
Sub Getscore()
Open App.Path & "\data.dat" For Random As #1
Get #1, 1, per
Close #1
End Sub

