with open("cells_to_link.txt", "r") as f:
   lines = f.readlines()

def formTuples(x):
   x = x.split()
   return zip(x[0::2], x[1::2])
   
tuples_by_line = map(formTuples, lines)


vba = """Public loc_map As Collection

Private Sub Workbook_Open()

   Set loc_map = New Collection"""


for index, tuples in enumerate(tuples_by_line):

   vba = vba + """
   ReDim loc_set(0 To """+ str(len(tuples) - 1) + """) As String 

"""
   for i, tup in enumerate(tuples):
      vba = vba + """   loc_set(""" + str(i) + """) = \"""" + " ".join(tup) + """\"
"""

   vba = vba + ("""
   Dim loc As Variant
""" if index == 0 else "")
   
   vba = vba + """
   For i = 0 To UBound(loc_set)
      loc_map.Add Minus(loc_set, loc_set(i)), loc_set(i)
   Next
   
   Erase loc_set
"""

vba = vba + """
End Sub

Private Function Minus(ByRef old_list() As String, loc As Variant) As String()

   Dim split_loc() As String
   split_loc = Split(loc)

   Dim new_list() As String
   ReDim new_list(0 To Application.CountA(old_list) - 1) As String

   Dim i As Integer
   i = 0

   Dim Val As Variant
   For Each Val In old_list
      Dim split_val() As String
      split_val = Split(Val)
      If StrComp(split_val(0) = split_loc(0), vbTextCompare) <> 0 Or StrComp(split_val(1) = split_loc(1), vbTextCompare) <> 0 Then
         new_list(i) = Val
         i = i + 1
      End If
   Next
      
   Minus = new_list
   
End Function

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
   Application.EnableEvents = False
   Dim cur_loc As String
   cur_loc = Target.address & " " & Target.Parent.name
   If Contains(loc_map, cur_loc) = True Then
      Dim loc As Variant
      For Each loc In loc_map.Item(cur_loc)
         Dim split_loc() As String
         split_loc = Split(loc)
         Sheets(split_loc(1)).Range(split_loc(0)).Value = Target.Value
      Next
   End If
   Application.EnableEvents = True
End Sub

Public Function Contains(col As Collection, key As Variant) As Boolean
    Dim obj As Variant
    On Error GoTo err
        Contains = True
        obj = col(key)
        Exit Function
err:
        Contains = False
End Function"""

with open("vba_code.txt", "w") as f:
   f.write(vba)
