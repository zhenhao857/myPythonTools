# 复制sheet

    Sub Macro1()
        sum_num = 3
        sum_num_min = sum_num - 2
        Sheets("0").Move before:=Sheets(1)
        For i = 1 To sum_num_min
        Sheets("0").Copy before:=Sheets(2)
        Sheets(2).Name = sum_num - i
        Next
      
    End Sub
    
# 查询sheet数量
    在某个单元格中设置公式 =INFO("numfile") 可以计算sheet数量
