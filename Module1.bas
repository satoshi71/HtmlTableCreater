Attribute VB_Name = "Module1"

Sub showForm()
   UserForm1.Show
End Sub


Sub createTable(padding, bcolor)
   Dim LoopArea As Range

   startR = Selection.Row
   startC = Selection.Column
   Set LoopArea = Selection
   endR = LoopArea.Cells(LoopArea.Count).Row
   endC = LoopArea.Cells(LoopArea.Count).Column

   'table_body = "<table border=""0"" cellpadding=""" & padding & """ cellspacing=""1"" bgcolor=""" & bcolor & """ width=""{total_width}"">" & vbNewLine
   table_body = "<table border=""0"" cellspacing=""1"" style=""padding:" & padding & "px; background:" & bcolor & "; width:{total_width}px;"">" & vbNewLine
   td = "<td style=""width:{width}px;height:{height}px;color:{fcoor};background:{bcolor};font-size:{fsize}pt;" & _
         "font-family:{fname};font-weight:{fweight};text-align:{align};padding:" & padding & "px;"">" & vbNewLine

   For r = startR To endR
      table_body = table_body & "<tr>"
      total_width = 0
      For c = startC To endC

         w = Int(Cells(r, c).ColumnWidth / 0.11797753 + 0.5) 'Width
         h = Int(Cells(r, c).RowHeight / 0.75) 'Height
         fcolor = getColorCode(Cells(r, c).Font.Color) 'ForeColor
         bcolor = getColorCode(Cells(r, c).Interior.Color) 'BackColor
         fsize = Cells(r, c).Font.Size
         fname = Cells(r, c).Font.Name
         isBold = Cells(r, c).Font.Bold
         falign = getAlignCode(Cells(r, c).HorizontalAlignment, Cells(r, c))

         total_width = total_width + w

         td_ = Replace(td, "{width}", w)
         td_ = Replace(td_, "{height}", h)
         td_ = Replace(td_, "{fcoor}", fcolor)
         td_ = Replace(td_, "{bcolor}", bcolor)
         td_ = Replace(td_, "{fsize}", fsize)
         td_ = Replace(td_, "{fname}", fname)
         td_ = Replace(td_, "{align}", falign)
         If isBold = True Then
            td_ = Replace(td_, "{fweight}", "bold")
         Else
            td_ = Replace(td_, "{fweight}", "nomal")
         End If

         'Format
         cell_value = Cells(r, c)
         value_format = Cells(r, c).NumberFormatLocal
         If Mid(value_format, 1, 2) <> "G/" And Mid(value_format, 1, 2) <> "Ge" Then
            cell_value = Format(cell_value, value_format)
         End If
   
         table_body = table_body & td_ & cell_value & "</td>"
               
      Next c
      table_body = table_body & "</tr>" & vbNewLine
   Next r

   table_body = Replace(table_body, "{total_width}", total_width)

   table_body = table_body & "</table>"
   table_html = "<html><body>" & vbNewLine & table_body & vbNewLine & "</body></html>"

   Open ActiveWorkbook.Path & "\index.html" For Output As #1
   Print #1, table_html
   Close #1


   UserForm2.TextBox1.Text = table_body
   UserForm2.Show


End Sub

Function getAlignCode(c, s)
   If c = xlCenter Then
      getAlignCode = "center"
   ElseIf c = xlLeft Then
      getAlignCode = "left"
   ElseIf c = xlRight Then
      getAlignCode = "right"
   Else
      If IsNumeric(s) = True Then
         getAlignCode = "right"
      Else
         getAlignCode = "left"
      End If
   End If

End Function


Function getColorCode(c)
   On Error GoTo ErrHandler

   red = Hex(c Mod 256)
   green = Hex(Int(c / 256) Mod 256)
   blue = Hex(Int(c / 256 / 256))

   getColorCode = "#" & Format(red, "00") & Format(green, "00") & Format(blue, "00")
   Exit Function

ErrHandler:
   getColorCode = "#dddddd"
End Function


