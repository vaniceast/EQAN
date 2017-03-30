Imports System.IO
Imports System.Math
Public Class Form1
    Public No_of_Saves As Boolean
    Public save_file As String
    Private Sub MyDGV_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim grid As DataGridView = CType(sender, DataGridView)
        Dim rowIdx As String = (e.RowIndex + 1).ToString()
        Dim rowFont As New System.Drawing.Font("Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Dim centerFormat = New StringFormat()
        centerFormat.Alignment = StringAlignment.Center
        centerFormat.LineAlignment = StringAlignment.Center
        Dim headerBounds As Rectangle = New Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(rowIdx, rowFont, SystemBrushes.ControlText, headerBounds, centerFormat)
    End Sub
    Private Sub dataGridView1_CellParsing(ByVal sender As Object, ByVal e As DataGridViewCellParsingEventArgs) Handles DataGridView1.CellParsing
        Dim v() As String, split_eqn() As String, chk1 As Boolean, chk2 As Boolean, temp() As String, temp1() As String, temp2() As String, temp_two_dim_input(,) As String, temp_two_dim_var(,) As String
        Dim k As Integer, row As Integer, i As Integer, p As Integer
        If e IsNot Nothing Then
            If e.Value IsNot Nothing Then
                If DataGridView1.CurrentCell.ColumnIndex = 0 Then
                    Try
                        split_eqn = split_str(e.Value)
                    Catch ex As Exception
                        MsgBox("check entered equation", MsgBoxStyle.OkOnly, "Error")
                        Exit Sub
                    End Try
                    row = DataGridView1.CurrentCell.RowIndex
                    If Len(Join(Equation_Array)) > 0 Then
                        If row <= UBound(Equation_Array) Then
                            Equation_Array(row) = e.Value
                        Else
                            chk1 = True
                        End If
                    Else
                        chk1 = True
                    End If
                    If chk1 = True Then
                        ReDim Preserve Equation_Array(No_of_Eqns) : ReDim Preserve Equation_Description(No_of_Eqns)
                        Equation_Array(No_of_Eqns) = e.Value
                        No_of_Eqns = No_of_Eqns + 1
                        Label2.Text = "Equations : " & No_of_Eqns
                    End If
                    v = get_operands2(split_eqn)
                    If chk1 = False Then
                        Call remove_row_and_element_variable_array(row, v)
                    End If
                    For i = 0 To UBound(v)
                        If Mid$(v(i), 1, 1) <> Chr(231) And v(i) <> "PI" Then
                            ReDim Preserve temp1(k) : ReDim Preserve temp2(k)
                            temp1(k) = v(i)
                            If Len(Join(Variable_array)) > 0 Then
                                If Variable_array.Contains(v(i)) = False Then
                                    chk2 = True
                                End If
                            Else
                                chk2 = True
                            End If
                            If chk2 = True Then
                                No_of_Variables = No_of_Variables + 1
                                ReDim Preserve Variable_array(No_of_Variables) : ReDim Preserve Variable_Input(No_of_Variables)
                                ReDim Preserve Variable_Description(No_of_Variables) : ReDim Preserve Unit(No_of_Variables)
                                Variable_array(No_of_Variables) = v(i)
                                Variable_Input(No_of_Variables) = "x"
                                DataGridView2.Rows.Add(New String() {"", v(i), "", ""})
                                chk2 = False
                                temp2(k) = "x"
                            Else
                                p = Array.IndexOf(Variable_array, v(i))
                                temp2(k) = Variable_Input(p)
                            End If
                            k = k + 1
                        End If
                    Next i
                    If chk1 = False Then
                        For i = 0 To UBound(Two_Dim_Variables, 1)
                            If i <> row Then
                                temp = get_one_dim_array(Two_Dim_Variables, i)
                                temp_two_dim_var = add_two_dim_array(temp_two_dim_var, temp)
                                temp = get_one_dim_array(Two_Dim_Input, i)
                                temp_two_dim_input = add_two_dim_array(temp_two_dim_input, temp)
                            Else
                                temp_two_dim_var = add_two_dim_array(temp_two_dim_var, temp1)
                                temp_two_dim_input = add_two_dim_array(temp_two_dim_input, temp2)
                            End If
                        Next i
                        Two_Dim_Variables = temp_two_dim_var
                        Two_Dim_Input = temp_two_dim_input
                    Else
                        Two_Dim_Variables = add_two_dim_array(Two_Dim_Variables, temp1)
                        Two_Dim_Input = add_two_dim_array(Two_Dim_Input, temp2)
                    End If
                    Label3.Text = "Variables : " & No_of_Variables + 1
                    Label6.Text = "Unknown : " & No_of_Variables - Known + 1
                ElseIf DataGridView1.CurrentCell.ColumnIndex = 1 Then
                    Equation_Description(DataGridView1.CurrentCell.RowIndex) = e.Value
                End If
            End If
        End If
    End Sub
    Function remove_row_and_element_variable_array(row As Integer, v() As String)
        Dim temp1() As String, temp2() As String
        Dim chk As Boolean
        Dim p As Integer
        temp1 = get_one_dim_array(Two_Dim_Variables, row)
        For i = 0 To UBound(temp1)
            chk = True
            For j = 0 To UBound(Two_Dim_Variables, 1)
                If j <> row Then
                    temp2 = get_one_dim_array(Two_Dim_Variables, j)
                    If temp2.Contains(temp1(i)) = True Then
                        chk = False
                    End If
                End If
            Next j
            If Len(Join(v)) > 0 Then
                If v.Contains(temp1(i)) = True Then
                    chk = False
                End If
            End If
            If chk = True Then
                p = Array.IndexOf(Variable_array, temp1(i))
                DataGridView2.Rows.Remove(DataGridView2.Rows(p))
                Variable_array = reduce_arr(Variable_array, p)
                Variable_Input = reduce_arr(Variable_Input, p)
                Variable_Description = reduce_arr(Variable_Description, p)
                Unit = reduce_arr(Unit, p)
                No_of_Variables = No_of_Variables - 1
            End If
        Next i
    End Function
    Private Sub dataGridView2_CellParsing(ByVal sender As Object, ByVal e As DataGridViewCellParsingEventArgs) Handles DataGridView2.CellParsing
        If e IsNot Nothing Then
            Dim row As Integer, column As Integer, p As Integer
            Dim var1 As String, var2 As String
            Dim cell_empty As Boolean
            row = DataGridView2.CurrentCell.RowIndex
            column = DataGridView2.CurrentCell.ColumnIndex
            var1 = DataGridView2.Rows(row).Cells(1).Value
            p = Array.IndexOf(Variable_array, var1)
            If column = 2 Then
                If Variable_Input(p) = "x" Then
                    cell_empty = True
                End If
                If IsNumeric(e.Value) = True Then
                    Variable_Input(p) = e.Value
                    var2 = e.Value
                    Known = Known + 1
                ElseIf e.Value = "" Then
                    Variable_Input(p) = "x"
                    var2 = "x"
                    If Known >= 1 Then
                        Known = Known - 1
                        cell_empty = True
                    End If
                End If
                Two_Dim_Input = update_two_dimensionsl_array(Two_Dim_Variables, Two_Dim_Input, var1, var2)
                If cell_empty = True Then
                    Label5.Text = "Known : " & Known
                    Label6.Text = "Unknown : " & No_of_Variables - Known + 1
                    cell_empty = False
                End If
            ElseIf column = 0 Then
                Variable_Description(p) = e.Value
            ElseIf column = 3 Then
                Unit(p) = e.Value
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim kneq As String
        Dim i As Integer, j As Integer, k As Integer
        Dim chk As Boolean
        Dim seqns() As String, n() As Integer, disp() As String, result() As String, v(,) As String
        Dim str As String, uknvar As String
        Unknown = Interaction.InputBox("Enter symbol of variable to be calculated", "Unknown Varaible")
        If Len(Unknown) > 0 Then
            If Len(Join(Variable_array)) > 0 Then
                If Variable_array.Contains(Unknown) = False Then
                    MsgBox("unknown variable entered", MsgBoxStyle.OkOnly, "Error")
                    Exit Sub
                End If
            Else
                MsgBox("equation database empty", MsgBoxStyle.OkOnly, "Error")
                Exit Sub
            End If
        Else
            MsgBox("varible not entered", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If
        k = 0
        For i = 0 To UBound(Two_Dim_Variables, 1)
            For j = 0 To UBound(Two_Dim_Variables, 2)
                If Two_Dim_Variables(i, j) = Unknown Then
                    ReDim Preserve Unknown_Eqn_Arr(k)
                    Unknown_Eqn_Arr(k) = Equation_Array(i)
                    k = k + 1
                    Exit For
                End If
            Next j
        Next i
        For i = 0 To UBound(Unknown_Eqn_Arr)
            'among set of equations find and select the equation which contains only one copy of unknown variable
            '("a=c+c","d=m+n","e=i+j") : c unkwn -> returns "a=c+c"
            chk = chk_eqn_solvable(Unknown_Eqn_Arr(i))
            If chk = True Then
                kneq = Unknown_Eqn_Arr(i)
                Exit For
            End If
        Next i
        ReDim Updated_Input_Array(UBound(Variable_Input))
        Variable_Input.CopyTo(Updated_Input_Array, 0)
        ReDim Updated_Two_Dim_Input(UBound(Two_Dim_Input, 1), UBound(Two_Dim_Input, 2))
        For i = 0 To UBound(Two_Dim_Input, 1)
            For j = 0 To UBound(Two_Dim_Input, 2)
                Updated_Two_Dim_Input(i, j) = Two_Dim_Input(i, j)
            Next j
        Next i
        If chk = True Then
            disp = calculate(kneq, Variable_Input)
            If disp(0) = "Using bisection method " Then
                Updated_Input_Array = update_list1(Updated_Input_Array, Unknown, disp(2))
                Updated_Two_Dim_Input = update_list2(Updated_Two_Dim_Input, Unknown, disp(2))
            Else
                Updated_Input_Array = update_list1(Updated_Input_Array, Unknown, disp(4))
                Updated_Two_Dim_Input = update_list2(Updated_Two_Dim_Input, Unknown, disp(4))
            End If
            Updated_Known = Known + 1
        Else
            Try
                seqns = solve_multi_unkwn()
            Catch ex As Exception
                MsgBox("solution not found", MsgBoxStyle.OkOnly, "Error")
                Exit Sub
            End Try
            n = get_eqn_numbers(seqns)
            Updated_Known = Known - 1
            For i = 0 To UBound(seqns)
                v = get_variables_of_selected_eqns(n)
                str = get_eqn_least_number_of_unknowns(seqns, Updated_Two_Dim_Input, v, n)
                result = calculate(str, Updated_Input_Array)
                disp = add_arr1_to_arr(disp, result)
                uknvar = get_unknown_of_eqn(Updated_Two_Dim_Input, str)
                If result(0) = "Using bisection method " Then
                    Updated_Input_Array = update_list1(Updated_Input_Array, uknvar, result(2))
                    Updated_Two_Dim_Input = update_list2(Updated_Two_Dim_Input, uknvar, result(2))
                Else
                    Updated_Input_Array = update_list1(Updated_Input_Array, uknvar, result(4))
                    Updated_Two_Dim_Input = update_list2(Updated_Two_Dim_Input, uknvar, result(4))
                End If
                Updated_Known = Updated_Known + 1
            Next i
            Updated_Known = Updated_Known + 1
        End If
        Call Form2.Create_and_Display_Form(disp)
    End Sub
    Function get_variables_of_selected_eqns(n() As Integer)
        Dim i As Integer, j As Integer
        Dim ret(,) As String
        ReDim ret(UBound(n), UBound(Two_Dim_Variables, 2))
        For i = 0 To UBound(n)
            For j = 0 To UBound(Two_Dim_Variables, 2)
                ret(i, j) = Two_Dim_Variables(n(i), j)
            Next j
        Next i
        get_variables_of_selected_eqns = ret
    End Function
    Public Shared Function update_input_variables()
        Dim i As Integer, p As Integer
        Variable_Input = Updated_Input_Array
        Two_Dim_Input = Updated_Two_Dim_Input
        For i = 0 To UBound(Variable_Input)
            p = Array.IndexOf(Variable_array, Form1.DataGridView2.Rows(i).Cells(1).Value)
            If Variable_Input(p) <> "x" Then
                Form1.DataGridView2.Rows(i).Cells(2).Value = Variable_Input(p)
            End If
        Next i
        Known = Updated_Known
        Form1.Label5.Text = "Known : " & Known
        Form1.Label6.Text = "Unknown : " & No_of_Variables - Known + 1
    End Function
    Function get_unknown_of_eqn(inputv(,) As String, eq As String)
        Dim n As Integer, ret As String
        Dim i As Integer
        n = Array.IndexOf(Equation_Array, eq)
        For i = 0 To UBound(inputv, 2)
            If inputv(n, i) = "x" Then
                ret = Two_Dim_Variables(n, i)
            End If
        Next i
        get_unknown_of_eqn = ret
    End Function
    Function update_list1(inputv() As String, sv As String, num As String)
        Dim i As Integer
        For i = 0 To UBound(Variable_array)
            If Variable_array(i) = sv Then
                inputv(i) = num
            End If
        Next i
        update_list1 = inputv
    End Function
    Function update_list2(inputv(,) As String, sv As String, num As String)
        Dim i As Integer
        For i = 0 To UBound(inputv, 1)
            For j = 0 To UBound(inputv, 2)
                If Two_Dim_Variables(i, j) = sv Then
                    inputv(i, j) = num
                End If
            Next j
        Next i
        update_list2 = inputv
    End Function
    Function get_eqn_least_number_of_unknowns(eqns() As String, inputv(,) As String, v(,) As String, n() As Integer)
        Dim temp() As String, temp1() As String, temp2() As String
        Dim i As Integer, j As Integer, k As Integer, p As Integer
        Dim chk As Boolean
        ReDim temp(UBound(n))
        For i = 0 To UBound(n)
            k = 0 : temp1 = get_one_dim_array(v, i)
            For j = 0 To UBound(inputv, 2)
                If inputv(n(i), j) = "x" Then
                    chk = True
                    If Len(Join(temp2)) > 0 Then
                        If temp2.Contains(temp1(j)) = True Then
                            chk = False
                        End If
                    End If
                    If chk = True Then
                        ReDim Preserve temp2(k)
                        temp2(k) = v(i, j)
                        k = k + 1
                    End If
                End If
                temp(i) = k
            Next j
            Erase temp2
        Next i
        p = Array.IndexOf(temp, "1")
        get_eqn_least_number_of_unknowns = eqns(p)
    End Function
    Function chk_eqn_solvable(eq As String)
        Dim chk As Boolean, arr() As String
        Dim i As Integer, k As Integer, n As Integer, l As Integer
        'get the number of unknowns of the equation : if the number of unkwns = 1 the equation is solvable
        k = 0 : chk = True : l = 0
        n = Array.IndexOf(Equation_Array, eq)
        For i = 0 To UBound(Two_Dim_Input, 2)
            'a=b+c : inv=(x,x,3) -> chk=False | a=b+b : inv=(10,x) -> chk=True
            If Two_Dim_Input(n, i) = "x" Then
                If Len(Join(arr)) = 0 Then
                    ReDim arr(l)
                    arr(l) = Two_Dim_Variables(n, i)
                    l = l + 1
                    k = k + 1
                    'a=b+b : inv=(10,x) -> chk=True 
                ElseIf arr.Contains(Two_Dim_Variables(n, i)) = False Then
                    ReDim Preserve arr(l)
                    arr(l) = Two_Dim_Variables(n, i)
                    l = l + 1
                    k = k + 1
                End If
            End If
        Next i
        If k > 1 Then
            chk = False
        End If
        chk_eqn_solvable = chk
    End Function
    Function solve_multi_unkwn()
        Dim uknv() As String, tmpv() As String, temp() As String, eqset() As String
        Dim ret(,) As String, uknvar() As String, path() As String, soln(,) As String, tmpsoln() As String, ignore() As String
        Dim i As Integer, eq() As String, chk_arr_empty As Boolean
        ReDim uknv(0) : uknv(0) = Unknown
        'get set of equations with main unknown and checks if it is sovable
        'a=b+c contains main unknown "a" & subunkwn "b" -> "b" should be present in some other equation for equation "a=b+c" to be solvable
        eq = return_solvable(Unknown, Unknown_Eqn_Arr)
        For i = 0 To UBound(eq)
            'from equation containg main unknown get sub unknown
            'a=b+c+d : inputv=("?","x","5","x") -> unkvar = (b,d)
            uknvar = get_unknown_var(eq(i))
            ReDim uknv(UBound(uknvar))
            uknvar.CopyTo(uknv, 0)
            Do
                'return array equations which conatin sub unknown ignoring equations already considered
                eqset = get_eqns(uknv, ret, eq(i), uknvar)
                'return equations whose variables are known or have the potential to solved
                eqset = return_solvable(uknv(0), eqset)
                If Len(Join(eqset)) > 0 Then
                    'add equations to two dimensional array
                    'arr=(("b=d1+d2","d1=e1+e2"),("b=d1+d2","d1=f1+f2"),("b=d1+d2","d1=f1+g1")) : unknown = f1 : eqset=("f1=SIN(h)","f1=COS(i)")
                    ' -> arr=(("b=d1+d2","d1=e1+e2"),("b=d1+d2","d1=f1+f2","f1=SIN(h)"),("b=d1+d2","d1=f1+f2","f1=COS(i)"),("b=d1+d2","d1=f1+g1","f1=SIN(h)"),("b=d1+d2","d1=f1+g1","f1=COS(i)"))
                    ret = add_eqns_two_dim_array(ret, eqset, uknv(0), uknvar)
                    'get unknowns of the newly added equations
                    tmpv = get_unknown_vars(uknv(0), eqset, ignore)
                    'uknv=(b,c) -> uknv=(c)
                    uknv = reduce_arr(uknv, 0)
                    If Len(Join(tmpv)) > 0 Then
                        'uknv=(c) : tmpv=(d1,d2) -> unkv=(c,d1,d2)
                        uknv = add_arr1_to_arr(uknv, tmpv)
                        ignore = add_arr1_to_arr(ignore, tmpv)
                    End If
                Else
                    chk_arr_empty = IsArrayEmpty(ret)
                    If chk_arr_empty = False Then
                        If UBound(ret, 1) <> 0 Then
                            'arr=((EQ1,EQ2),(EQ3,EQ4),(EQ5,EQ6,EQ7),(EQ8,EQ9,EQ10)) : EQ3 & EQ8 contains unkvar -> arr=((EQ1,EQ2),(EQ5,EQ6,EQ7))
                            ret = remove_eqns_two_dim_array(ret, uknv(0))
                            'uknv=(b,c) -> uknv=(c)
                            uknv = reduce_arr(uknv, 0)
                        Else
                            Erase ret, uknv
                            Exit For
                        End If
                    Else
                        Erase uknv
                    End If
                End If
            Loop Until Len(Join(uknv)) = 0
            Erase ignore
            chk_arr_empty = IsArrayEmpty(ret)
            If chk_arr_empty = False Then
                path = get_shortest_path(ret)
                tmpsoln = reverse_eqns(path)
                tmpsoln = add_element_to_array(tmpsoln, UBound(tmpsoln) + 1, eq(i))
                soln = add_two_dim_array(soln, tmpsoln)
            End If
        Next i
        temp = get_shortest_path(soln)
        solve_multi_unkwn = temp
    End Function
    Function reverse_eqns(eqns() As String)
        Dim ret() As String
        Dim k As Integer
        k = UBound(eqns)
        ReDim ret(k)
        For i = 0 To UBound(eqns)
            ret(i) = eqns(k)
            k = k - 1
        Next i
        reverse_eqns = ret
    End Function
    Function get_shortest_path(arr(,) As String)
        Dim ret() As String, num() As Integer
        Dim i As Integer, min As Integer, p As Integer
        ReDim num(UBound(arr, 1))
        For i = 0 To UBound(arr, 1)
            num(i) = UBound(get_one_dim_array(arr, i))
        Next i
        min = num.Min
        p = Array.IndexOf(num, min)
        ret = (get_one_dim_array(arr, p))
        get_shortest_path = ret
    End Function
    Function remove_eqns_two_dim_array(arr(,) As String, unkvar As String)
        Dim ret(,) As String, temp() As String, temp1() As String, temp2() As String, n() As Integer
        Dim i As Integer, j As Integer, m As Integer, k As Integer, l As Integer, p As Integer
        'check which equations in the main equation array(arr) contain the unknown ; if the equation contains unknown gwt the equation number
        'arr=((EQ1,EQ2),(EQ3,EQ4),(EQ5,EQ6,EQ7),(EQ8,EQ9,EQ10)) : EQ3 & EQ8 contains unkvar -> n=(2,4)
        For i = 0 To UBound(arr, 1)
            temp1 = get_one_dim_array(arr, i)
            For j = 0 To UBound(temp1)
                p = Array.IndexOf(Equation_Array, temp1(j))
                temp2 = get_one_dim_array(Two_Dim_Variables, p)
                If temp2.Contains(unkvar) = True Then
                    ReDim Preserve n(k)
                    n(k) = i
                    k = k + 1
                    Exit For
                End If
            Next j
        Next i
        'create new array removing rows whose equation contain unsolvable variable
        'arr=((EQ1,EQ2),(EQ3,EQ4),(EQ5,EQ6,EQ7),(EQ8,EQ9,EQ10))  -> ret=((EQ1,EQ2),(EQ5,EQ6,EQ7))
        k = UBound(arr, 1) - k
        l = UBound(arr, 2)
        ReDim ret(k, l)
        m = 0
        For i = 0 To UBound(arr, 1)
            If n.Contains(i) = False Then
                temp = get_one_dim_array(arr, i)
                For l = 0 To UBound(temp)
                    ret(m, l) = temp(l)
                Next l
                m = m + 1
            End If
        Next i
        remove_eqns_two_dim_array = ret
    End Function
    Function get_unknown_vars(var As String, eq() As String, ignore() As String)
        Dim n() As Integer, ret() As String
        Dim i As Integer, j As Integer, k As Integer
        Dim chk As Boolean
        'get the index number of equation in equation set
        'ukneq=("a=b*c","a=d+e") : eqns=("i=m+n","a=b*c","j=h+j","a=d+e") -> n=(2,4)
        n = get_eqn_numbers(eq) : k = 0
        For i = 0 To UBound(n)
            For j = 0 To UBound(Two_Dim_Variables, 2)
                'return unknown varaibles in equation ignoring main unknown
                'eq="a=b+c" : a mani unkwn : inputv=("x","2","x") -> ret=("c")
                If Two_Dim_Variables(n(i), j) <> var And Two_Dim_Input(n(i), j) = "x" Then
                    chk = False
                    'make sure the unknown does not repeat twice(a=b+b, b is unknwon) & not equal to the main unknown(var)
                    If Len(Join(ret)) > 0 Then
                        If ret.Contains(Two_Dim_Variables(n(i), j)) = False Then
                            chk = True
                        End If
                    Else
                        chk = True
                    End If
                    If Len(Join(ignore)) > 0 Then
                        If ignore.Contains(Two_Dim_Variables(n(i), j)) = True Then
                            chk = False
                        End If
                    End If
                    If chk = True Then
                        ReDim Preserve ret(k)
                        ret(k) = Two_Dim_Variables(n(i), j)
                        k = k + 1
                    End If
                End If
            Next j
        Next i
        get_unknown_vars = ret
    End Function
    Function add_eqns_two_dim_array(arr(,) As String, eqset() As String, unkwn As String, main_unknown() As String)
        Dim ret(,) As String, temp() As String
        Dim i As Integer, k As Integer, chk As Boolean, chk_arr_empty As Boolean
        chk_arr_empty = IsArrayEmpty(arr)
        If chk_arr_empty = True Then
            If UBound(eqset) > 0 Then
                For i = 0 To UBound(eqset)
                    Dim temp1() As String = {eqset(i)}
                    ret = add_two_dim_array(ret, temp1)
                Next i
            Else
                ret = add_two_dim_array(ret, eqset)
            End If
        Else
            For i = 0 To UBound(arr, 1)
                'ensure only those rows in two dimensional array which conatain unknown will have eqset added to the ends
                'arr=(("b=d1+d2","d1=e1+e2"),("b=d1+d2","d1=f1+f2"),("b=d1+d2","d1=f1+g1")) : unknown = f1 : eqset=("f1=SIN(h)","f1=COS(i)")
                ' -> arr=(("b=d1+d2","d1=e1+e2"),("b=d1+d2","d1=f1+f2","f1=SIN(h)"),("b=d1+d2","d1=f1+f2","f1=COS(i)"),("b=d1+d2","d1=f1+g1","f1=SIN(h)"),("b=d1+d2","d1=f1+g1","f1=COS(i)"))
                temp = get_one_dim_array(arr, i)
                If main_unknown.Contains(unkwn) = False Then
                    chk = check_unknwon_in_eqset(temp, unkwn)
                End If
                If chk = True Or main_unknown.Contains(unkwn) = True Then
                    '(b=d1+d2) : unknown = d1 : (d1=SIN(i),d1=SIN(j)) -> ((b=d1+d2,d1=SIN(i)),(b=d1+d2,d1=SIN(j)))
                    ReDim Preserve temp(UBound(temp) + 1)
                    For k = 0 To UBound(eqset)
                        temp(UBound(temp)) = eqset(k)
                        ret = add_two_dim_array(ret, temp)
                    Next k
                Else
                    ret = add_two_dim_array(ret, temp)
                End If
            Next i
        End If
        add_eqns_two_dim_array = ret
    End Function
    Function check_unknwon_in_eqset(eqset() As String, unkwn As String)
        Dim temp() As String, n() As Integer
        Dim chk As Boolean
        'get index number of equations in eqset
        n = get_eqn_numbers(eqset)
        For i = 0 To UBound(n)
            temp = get_one_dim_array(Two_Dim_Variables, n(i))
            'if unknown in equation chk = True
            If temp.Contains(unkwn) = True Then
                chk = True
            End If
        Next i
        check_unknwon_in_eqset = chk
    End Function
    Function get_eqns(unkvars() As String, arr(,) As String, eq As String, main_unknowns() As String)
        Dim temp1() As String, temp2() As String, ret() As String
        Dim i As Integer, j As Integer, l As Integer, k As Integer
        Dim chk As Boolean, chk_arr_empty As Boolean
        'returns an array of equations which contain the unknown variable ignoring equations already considered and equation under consideration
        chk_arr_empty = IsArrayEmpty(arr)
        l = 0
        For i = 0 To UBound(Two_Dim_Variables, 1)
            temp1 = get_one_dim_array(Two_Dim_Variables, i)
            If temp1.Contains(unkvars(0)) = True Then
                If chk_arr_empty = False Then
                    chk = False
                    'check if equation already considered in equation set(arr)
                    For j = 0 To UBound(arr, 1)
                        temp2 = get_one_dim_array(arr, j)
                        If temp2.Contains(Equation_Array(i)) = True Then
                            chk = True
                        End If
                    Next j
                End If
                'check that unknowns of main equation are not considered while selecting equation
                If chk = False And main_unknowns.Contains(unkvars(0)) = False Then
                    For k = 0 To UBound(main_unknowns)
                        If temp1.Contains(main_unknowns(k)) = True Then
                            chk = True
                            Exit For
                        End If
                    Next k
                End If
                If chk = False Or chk_arr_empty = True Then
                    'check that equation not equal to equation which contain main unknown
                    If Equation_Array(i) <> eq Then
                        ReDim Preserve ret(l)
                        ret(l) = Equation_Array(i)
                        l = l + 1
                    End If
                End If
            End If
        Next i
        get_eqns = ret
    End Function
    Function get_unknown_var(eq As String)
        Dim n As Integer, ret() As String
        Dim i As Integer, k As Integer
        n = Array.IndexOf(Equation_Array, eq) : k = 0
        For i = 0 To UBound(Two_Dim_Variables, 2)
            If Two_Dim_Variables(n, i) <> Unknown And Two_Dim_Input(n, i) = "x" Then
                If Len(Join(ret)) = 0 Then
                    ReDim Preserve ret(k)
                    ret(k) = Two_Dim_Variables(n, i)
                    k = k + 1
                Else
                    If ret.Contains(Two_Dim_Variables(n, i)) = False Then
                        ReDim Preserve ret(k)
                        ret(k) = Two_Dim_Variables(n, i)
                        k = k + 1
                    End If
                End If
            End If
        Next i
        get_unknown_var = ret
    End Function
    Function convert_to_string(n() As Integer)
        Dim i As Integer
        Dim ret() As String
        ReDim ret(UBound(n))
        For i = 0 To UBound(n)
            ret(i) = CStr(n(i))
        Next i
        convert_to_string = ret
    End Function
    Function return_solvable(var As String, ukneq() As String)
        Dim i As Integer, j As Integer, k As Integer
        Dim temp() As String, chk(,) As String, ret() As String, n() As Integer
        Dim check_arr_not_empty As Boolean
        'get the index number of equation in equation set which contain unknown variable
        'ukneq=("a=b*c","a=d+e") : a is unkwn : eqns=("i=m+n","a=b*c","j=h+j","a=d+e") -> n=(2,4)
        n = get_eqn_numbers(ukneq)
        ReDim chk(UBound(n), UBound(Two_Dim_Variables, 2))
        For i = 0 To UBound(n)
            k = 0
            For j = 0 To UBound(Two_Dim_Variables, 2)
                If Two_Dim_Variables(n(i), j) <> var And Two_Dim_Input(n(i), j) = "x" Then
                    'checks if unknown in equation has the potential to be solved
                    'a=b+c : a is the main unknown : b is sub-unkown -> seach if b is present in another equation : ignore already considered equations with respect to their index numbers "n"
                    chk(i, k) = chk_var_present(Two_Dim_Variables(n(i), j), n)
                    check_arr_not_empty = True
                    k = k + 1
                Else
                    chk(i, k) = "1"
                    k = k + 1
                End If
            Next j
        Next i
        If check_arr_not_empty = True Then
            k = 0
            For i = 0 To UBound(chk, 1)
                temp = get_one_dim_array(chk, i)
                'if the equation has pontential to be solved i.e. no "0" present then add to list of return
                If temp.Contains("0") = False Then
                    ReDim Preserve ret(k)
                    ret(k) = ukneq(i)
                    k = k + 1
                End If
            Next i
        Else
            ReDim ret(0)
            ret(0) = ukneq(0)
        End If
        return_solvable = ret
    End Function
    Function chk_var_present(chk As String, n() As Integer)
        Dim i As Integer
        Dim temp() As String
        'check if unknown variable has the potential to be solved by looking for the same variable in other equations
        'a=b+c : a is the main unknown : b is sub-unkown -> seach if b is present in another equation & ignore equation already considered
        For i = 0 To UBound(Two_Dim_Variables, 1) 'Variable_array.Contains(v(i)) = True
            If n.Contains(i) = False Then
                temp = get_one_dim_array(Two_Dim_Variables, i)
                If temp.Contains(chk) = True Then
                    chk_var_present = 1
                    Exit Function
                End If
            End If
        Next i
        chk_var_present = 0
    End Function
    Function get_eqn_numbers(eq() As String)
        Dim i As Integer, p As Integer, k As Integer
        Dim n() As Integer
        'get the index number of equation in equation set
        'ukneq=("a=b*c","a=d+e") : eqns=("i=m+n","a=b*c","j=h+j","a=d+e") -> n=(2,4)
        k = 0
        For i = 0 To UBound(eq)
            p = Array.IndexOf(Equation_Array, eq(i))
            ReDim Preserve n(k)
            n(k) = p
            k = k + 1
        Next i
        get_eqn_numbers = n
    End Function
    Function update_two_dimensionsl_array(arr1(,) As String, arr2(,) As String, var1 As String, var2 As String)
        Dim i As Integer, j As Integer
        For i = 0 To UBound(arr1, 1)
            For j = 0 To UBound(arr1, 2)
                If arr1(i, j) = var1 Then
                    arr2(i, j) = var2
                End If
            Next j
        Next i
        update_two_dimensionsl_array = arr2
    End Function
    Function calculate(eqn As String, inputv() As String)
        Dim split_eqn() As String, v() As String, input_v() As String, result() As String
        Dim pos As Integer, count As Integer, unkvar As String
        split_eqn = split_str(eqn)
        input_v = get_eqn_input(eqn, inputv)
        v = get_operands2(split_eqn)
        unkvar = get_unknown_of_equation(eqn)
        count = number_of_instances(v, unkvar)
        If count = 1 Then
            pos = Array.IndexOf(split_eqn, unkvar)
            If pos = 0 Then
                result = disp(split_eqn, input_v)
            Else
                result = eqn_typ_1(split_eqn, input_v)
            End If
        Else
            result = eqn_typ_2(split_eqn, input_v, v, unkvar)
        End If
        calculate = result
    End Function
    Function get_unknown_of_equation(eq As String)
        Dim n As Integer, ret As String
        Dim i As Integer
        n = Array.IndexOf(Equation_Array, eq)
        For i = 0 To UBound(Two_Dim_Input, 2)
            If Updated_Two_Dim_Input(n, i) = "x" Then
                ret = Two_Dim_Variables(n, i)
                Exit For
            End If
        Next i
        get_unknown_of_equation = ret
    End Function
    Function get_eqn_input(eqn As String, inputv() As String)
        Dim split_eqn() As String, v() As String, input_v() As String
        Dim i As Integer
        split_eqn = split_str_with_numbers(eqn)
        v = get_operands2(split_eqn)
        ReDim input_v(UBound(v))
        For i = 0 To UBound(v)
            If Variable_array.Contains(v(i)) = True Then
                input_v(i) = inputv(Array.IndexOf(Variable_array, v(i)))
            Else
                input_v(i) = v(i)
            End If
        Next i
        get_eqn_input = input_v
    End Function
    Function number_of_instances(v() As String, unkvar As String)
        Dim i As Integer, k As Integer
        For i = 0 To UBound(v)
            If unkvar = v(i) Then
                k = k + 1
            End If
        Next i
        number_of_instances = k
    End Function

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If Len(Join(Variable_Input)) > 0 Then
            Dim i As Integer
            For i = 0 To UBound(Variable_Input)
                DataGridView2.Rows(i).Cells(2).Value = ""
            Next i
            Known = 0
            Label5.Text = "Known : " & Known
            Label6.Text = "Unknown : " & No_of_Variables + 1
        End If
    End Sub

    Private Sub DataGridView1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        If e.Button = MouseButtons.Right Then
            Dim x As Integer = MousePosition().X
            Dim y As Integer = MousePosition().Y
            If DataGridView1.Rows.Count > 0 Then
                Dim ht As DataGridView.HitTestInfo = DataGridView1.HitTest(e.X, e.Y)
                If ht.Type = DataGridViewHitTestType.Cell Then
                    ContextMenuStrip1.Show(x, y)
                    Equation_Delete_Row_Index = ht.RowIndex
                End If
            End If
        End If
    End Sub

    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked
        Select Case e.ClickedItem.ToString()
            Case "delete"
                If Len(DataGridView1.Item(0, Equation_Delete_Row_Index).Value) > 0 Then
                    Dim temp() As String, temp_two_dim_var(,) As String, temp_two_dim_input(,) As String
                    Dim row As Integer, i As Integer
                    row = Equation_Delete_Row_Index
                    Call remove_row_and_element_variable_array(row, temp)
                    For i = 0 To UBound(Two_Dim_Variables, 1)
                        If i <> row Then
                            temp = get_one_dim_array(Two_Dim_Variables, i)
                            temp_two_dim_var = add_two_dim_array(temp_two_dim_var, temp)
                            temp = get_one_dim_array(Two_Dim_Input, i)
                            temp_two_dim_input = add_two_dim_array(temp_two_dim_input, temp)
                        End If
                    Next i
                    Two_Dim_Variables = temp_two_dim_var
                    Two_Dim_Input = temp_two_dim_input
                    Equation_Array = reduce_arr(Equation_Array, row)
                    No_of_Eqns = No_of_Eqns - 1
                    Label2.Text = "Equations = " & No_of_Eqns
                    Label3.Text = "Variables : " & No_of_Variables + 1
                    Known = 0
                    If Len(Join(Variable_Input)) > 0 Then
                        For i = 0 To UBound(Variable_Input)
                            If Variable_Input(i) <> "x" Then
                                Known = Known + 1
                            End If
                        Next i
                    End If
                    Label5.Text = "Known : " & Known
                    Label6.Text = "Unknown : " & No_of_Variables - Known + 1
                End If
                DataGridView1.Rows.RemoveAt(Equation_Delete_Row_Index)
        End Select
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If No_of_Saves = False Then
            Dim c As Char
            Dim k As Integer
            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Title = "Save"
            saveFileDialog1.OverwritePrompt = True
            saveFileDialog1.DefaultExt = ".eqn"
            saveFileDialog1.AddExtension = True
            Dim DR As DialogResult = saveFileDialog1.ShowDialog
            If saveFileDialog1.FileName <> "" Then
                If DR = Windows.Forms.DialogResult.OK Then
                    save_file = saveFileDialog1.FileName
                    c = Mid(save_file, Len(save_file))
                    k = Len(save_file) - 1
                    Do
                        User_File_Name = User_File_Name & c
                        c = Mid(save_file, k)
                        k = k - 1
                    Loop Until c = "\"
                    User_File_Name = StrReverse(User_File_Name)
                    User_File_Name = Replace(User_File_Name, ".eqn", "")
                    Me.Text = "EQAN - " & User_File_Name
                    Call save_to_file()
                    No_of_Saves = True
                End If
            End If
        Else
            Call save_to_file()
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Title = "Save"
        saveFileDialog1.OverwritePrompt = True
        saveFileDialog1.DefaultExt = ".eqn"
        saveFileDialog1.AddExtension = True
        Dim DR As DialogResult = saveFileDialog1.ShowDialog
        If saveFileDialog1.FileName <> "" Then
            If DR = Windows.Forms.DialogResult.OK Then
                save_file = saveFileDialog1.FileName
                Call save_to_file()
                No_of_Saves = True
            End If
        End If
    End Sub
    Function save_to_file()
        Dim i As Integer, j As Integer
        Dim new_file As System.IO.StreamWriter
        If System.IO.File.Exists(save_file) = True Then
            File.Delete(save_file)
        End If
        new_file = My.Computer.FileSystem.OpenTextFileWriter(save_file, True)
        If Len(Join(Equation_Array)) > 0 Then
            new_file.WriteLine(No_of_Eqns)
            For i = 0 To UBound(Equation_Array)
                new_file.WriteLine(Equation_Array(i))
            Next i
            For i = 0 To UBound(Equation_Description)
                new_file.WriteLine(Equation_Description(i))
            Next i
            new_file.WriteLine(UBound(Variable_array))
            For i = 0 To UBound(Variable_array)
                new_file.WriteLine(Variable_array(i))
            Next i
            For i = 0 To UBound(Variable_Input)
                new_file.WriteLine(Variable_Input(i))
            Next i
            For i = 0 To UBound(Variable_Description)
                new_file.WriteLine(Variable_Description(i))
            Next i
            For i = 0 To UBound(Unit)
                new_file.WriteLine(Unit(i))
            Next i
            new_file.WriteLine(UBound(Two_Dim_Variables, 1))
            new_file.WriteLine(UBound(Two_Dim_Variables, 2))
            For i = 0 To UBound(Two_Dim_Variables, 1)
                For j = 0 To UBound(Two_Dim_Variables, 2)
                    new_file.WriteLine(Two_Dim_Variables(i, j))
                Next j
            Next i
            For i = 0 To UBound(Two_Dim_Input, 1)
                For j = 0 To UBound(Two_Dim_Input, 2)
                    new_file.WriteLine(Two_Dim_Input(i, j))
                Next j
            Next i
            new_file.WriteLine(Known)
            new_file.WriteLine(save_file)
            new_file.WriteLine(User_File_Name)
            new_file.WriteLine(Decimal_Pt)
            new_file.Close()
        End If
    End Function
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim i As Integer, j As Integer, k As Integer, n As Integer
            If Len(Join(Equation_Array)) > 0 Then
                Erase Equation_Array, Variable_array, Variable_Input, Two_Dim_Input, Two_Dim_Variables, Unit, Variable_Description
                DataGridView1.Rows.Clear()
                DataGridView2.Rows.Clear()
            End If
            Dim filename As String
            Dim arr() As String
            filename = openFileDialog1.FileName
            arr = File.ReadAllLines(filename)
            k = 0
            No_of_Eqns = arr(k)
            n = arr(k) - 1
            k = k + 1
            ReDim Equation_Array(n)
            For i = 0 To n
                Equation_Array(i) = arr(k)
                k = k + 1
            Next i
            ReDim Equation_Description(n)
            For i = 0 To n
                Equation_Description(i) = arr(k)
                k = k + 1
            Next i
            For i = 0 To n
                DataGridView1.Rows.Add(Equation_Array(i), Equation_Description(i))
            Next i
            No_of_Variables = arr(k)
            ReDim Variable_array(No_of_Variables) : ReDim Variable_Input(No_of_Variables) : ReDim Variable_Description(No_of_Variables) : ReDim Unit(No_of_Variables)
            k = k + 1
            For i = 0 To No_of_Variables
                Variable_array(i) = arr(k)
                k = k + 1
                n = n + 1
            Next i
            For i = 0 To No_of_Variables
                Variable_Input(i) = arr(k)
                k = k + 1
            Next i
            For i = 0 To No_of_Variables
                Variable_Description(i) = arr(k)
                k = k + 1
            Next i
            For i = 0 To No_of_Variables
                Unit(i) = arr(k)
                k = k + 1
            Next i
            For i = 0 To No_of_Variables
                If Variable_Input(i) = "x" Then
                    DataGridView2.Rows.Add(New String() {Variable_Description(i), Variable_array(i), "", Unit(i)})
                Else
                    DataGridView2.Rows.Add(New String() {Variable_Description(i), Variable_array(i), Variable_Input(i), Unit(i)})
                End If
            Next i
            ReDim Two_Dim_Variables(arr(k), arr(k + 1)) : ReDim Two_Dim_Input(arr(k), arr(k + 1))
            k = k + 2
            For i = 0 To UBound(Two_Dim_Variables, 1)
                For j = 0 To UBound(Two_Dim_Variables, 2)
                    Two_Dim_Variables(i, j) = arr(k)
                    k = k + 1
                Next j
            Next i
            For i = 0 To UBound(Two_Dim_Input, 1)
                For j = 0 To UBound(Two_Dim_Input, 2)
                    Two_Dim_Input(i, j) = arr(k)
                    k = k + 1
                Next j
            Next i
            Known = arr(k)
            save_file = arr(k + 1)
            User_File_Name = arr(k + 2)
            Decimal_Pt = arr(k + 3)
            Me.Text = "EQAN - " & User_File_Name
            Label2.Text = "Equations = " & No_of_Eqns
            Label3.Text = "Variables : " & No_of_Variables + 1
            Label5.Text = "Known : " & Known
            Label6.Text = "Unknown : " & No_of_Variables - Known + 1
            No_of_Saves = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim temp As Integer = Decimal_Pt
            Decimal_Pt = Interaction.InputBox("Enter number of decimal points: ", "Decimal Points")
            If Decimal_Pt < 0 Then
                Decimal_Pt = temp
                MsgBox("check entered number", MsgBoxStyle.OkOnly, "Error")
            ElseIf Decimal_Pt > 9 Then
                Decimal_Pt = temp
                MsgBox("enter number less than or equal to 9", MsgBoxStyle.OkOnly, "Error")
            End If
        Catch ex As Exception
            MsgBox("check entered number", MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub
End Class
Module Module1
    Public Equation_Array() As String, Variable_array() As String, Variable_Input() As String, Two_Dim_Variables(,) As String, Two_Dim_Input(,) As String, Unknown_Eqn_Arr() As String
    Public No_of_Eqns As Integer, row_index As Integer, Known As Integer
    Public Unknown As String, User_File_Name As String
    Public No_of_Variables As Integer = -1
    Public Decimal_Pt As Integer = 4
    Public Equation_Delete_Row_Index As Integer
    Public Variable_Description() As String, Unit() As String, Equation_Description() As String
    Public Updated_Input_Array() As String, Updated_Two_Dim_Input(,) As String, Updated_Known As Integer
    Public Common_Operator As String
    Public app = New Microsoft.Office.Interop.Excel.Application
    Function split_str(eq As String)
        Dim ret() As String, temp() As String
        Dim i As Integer, k As Integer, str As String, c As Integer
        Dim arr = New String() {"+", "-", "*", "/", "^", "(", ")", "=", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT", "PI"}
        k = 0 : c = 1
        For i = 1 To Len(eq)
            'convert and save all number constants of the equation : a=SIN(90+b)+2 -> "90" & "2" are number constants
            If IsNumeric(UCase(Mid$(eq, i, 1))) = True Then
                'a=230+b1 -> "230" saved as "Chr(231) + 1" ; "1" of "b1" saved as "1" as it is not preceded by arr component
                If arr.Contains(temp(k - 1)) = True Then
                    Do
                        i = i + 1
                        If (Mid$(eq, i, 1)) = "." Then
                            i = i + 1
                        End If
                    Loop Until IsNumeric(UCase(Mid$(eq, i, 1))) = False Or i = Len(eq)
                    ReDim Preserve temp(k)
                    temp(k) = Chr(231) & c
                    c = c + 1
                    k = k + 1
                End If
            End If
            If {"LN", "PI"}.Contains(UCase(Mid$(eq, i, 2))) = True And InStr(Mid$(eq, i, 2), "*") = 0 Then
                ReDim Preserve temp(k)
                temp(k) = UCase(Mid$(eq, i, 2))
                k = k + 1
                i = i + 1
            ElseIf {"LOG", "EXP", "SIN", "COS", "TAN"}.Contains(UCase(Mid$(eq, i, 3))) = True And InStr(Mid$(eq, i, 3), "*") = 0 Then
                ReDim Preserve temp(k)
                temp(k) = UCase(Mid$(eq, i, 3))
                k = k + 1
                i = i + 2
                'ElseIf IsError(Application.Match(UCase(Mid$(eq, i, 4)), Array("ASIN", "ACOS", "ATAN"), False)) = False And _
                'InStr(Mid$(eq, i, 4), "*") = 0 Then
            ElseIf {"ASIN", "ACOS", "ATAN", "SQRT"}.Contains(UCase(Mid$(eq, i, 4))) = True And InStr(Mid$(eq, i, 4), "*") = 0 Then
                ReDim Preserve temp(k)
                temp(k) = UCase(Mid$(eq, i, 4))
                k = k + 1
                i = i + 3
            Else
                ReDim Preserve temp(k)
                temp(k) = Mid$(eq, i, 1)
                k = k + 1
            End If
        Next i
        str = "" : k = 0
        For i = 0 To UBound(temp)
            'combines variabe name to form variable
            'a1=b1+b2 -> t(1)="a" & t(2)="1" -> combine and stored as "a1"
            If arr.Contains(temp(i)) = False Then
                Do
                    str = str & temp(i)
                    If i + 1 > UBound(temp) Then
                        i = UBound(temp)
                    Else
                        i = i + 1
                    End If
                Loop Until arr.Contains(temp(i)) = True Or i = UBound(temp)
                ReDim Preserve ret(k)
                ret(k) = str
                k = k + 1
            End If
            'combines last variable name of string
            'a1=b1+b2 -> t(n-1)="b" & t(n)="2" (n=ubound(t)) -> combine and stored as "b2"
            If i = UBound(temp) And temp(i) <> ")" And temp(i) <> "PI" Then
                'a=b+c -> variable "c" should not get stored twice
                If temp(i) <> ret(k - 1) Then
                    str = str & temp(i)
                    ret(k - 1) = str
                End If
            End If
            str = ""
            'save all operators of equation ("+", "-", "LOG", "SIN", etc. )
            If arr.Contains(temp(i)) = True Then
                ReDim Preserve ret(k)
                ret(k) = temp(i)
                k = k + 1
            End If
        Next i
        'a=(b+c) -> a=b+c
        ret = remove_brackets_if_needed(ret)
        'a=(b)+c -> a=b+c
        ret = remove_brackets_single_variable(ret)
        'a=-b+c -> a=(-b)+c
        ret = add_brackets_negative_variables(ret)
        'SINb ->SIN(b) & LOGz -> LOG(z)
        ret = add_brackets_operands(ret)
        'a=SIN(b)+COS(c) -> a=(SIN(b))+(COS(c))
        ret = check_split_str(ret)
        split_str = ret
    End Function
    Function split_str_with_numbers(eq As String)
        Dim ret() As String, temp() As String
        Dim i As Integer, k As Integer, str As String, c As Integer
        Dim num As String
        Dim arr = New String() {"+", "-", "*", "/", "^", "(", ")", "=", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT", "PI"}
        k = 0 : c = 1
        For i = 1 To Len(eq)
            'convert and save all number constants of the equation : a=SIN(90+b)+2 -> "90" & "2" are number constants
            If IsNumeric(UCase(Mid$(eq, i, 1))) = True Then
                'a=230+b1 -> "230" saved as "Chr(231) + 1" ; "1" of "b1" saved as "1" as it is not preceded by arr component
                If arr.Contains(temp(k - 1)) = True Then
                    num = ""
                    Do
                        num = num & (Mid$(eq, i, 1))
                        i = i + 1
                        If (Mid$(eq, i, 1)) = "." Then
                            num = num & (Mid$(eq, i, 1))
                            i = i + 1
                        End If
                    Loop Until IsNumeric(UCase(Mid$(eq, i, 1))) = False Or i = Len(eq)
                    ReDim Preserve temp(k)
                    temp(k) = num
                    c = c + 1
                    k = k + 1
                End If
            End If
            If {"LN", "PI"}.Contains(UCase(Mid$(eq, i, 2))) = True And InStr(Mid$(eq, i, 2), "*") = 0 Then
                ReDim Preserve temp(k)
                temp(k) = UCase(Mid$(eq, i, 2))
                k = k + 1
                i = i + 1
            ElseIf {"LOG", "EXP", "SIN", "COS", "TAN"}.Contains(UCase(Mid$(eq, i, 3))) = True And InStr(Mid$(eq, i, 3), "*") = 0 Then
                ReDim Preserve temp(k)
                temp(k) = UCase(Mid$(eq, i, 3))
                k = k + 1
                i = i + 2
            ElseIf {"ASIN", "ACOS", "ATAN", "SQRT"}.Contains(UCase(Mid$(eq, i, 4))) = True And InStr(Mid$(eq, i, 4), "*") = 0 Then
                ReDim Preserve temp(k)
                temp(k) = UCase(Mid$(eq, i, 4))
                k = k + 1
                i = i + 3
            Else
                ReDim Preserve temp(k)
                temp(k) = Mid$(eq, i, 1)
                k = k + 1
            End If
        Next i
        str = "" : k = 0
        For i = 0 To UBound(temp)
            'combines variabe name to form variable
            'a1=b1+b2 -> t(1)="a" & t(2)="1" -> combine and stored as "a1"
            If arr.Contains(temp(i)) = False Then
                Do
                    str = str & temp(i)
                    If i + 1 > UBound(temp) Then
                        i = UBound(temp)
                    Else
                        i = i + 1
                    End If
                Loop Until arr.Contains(temp(i)) = True Or i = UBound(temp)
                ReDim Preserve ret(k)
                ret(k) = str
                k = k + 1
            End If
            'combines last variable name of string
            'a1=b1+b2 -> t(n-1)="b" & t(n)="2" (n=ubound(t)) -> combine and stored as "b2"
            If i = UBound(temp) And temp(i) <> ")" And temp(i) <> "PI" Then
                'a=b+c -> variable "c" should not get stored twice
                If temp(i) <> ret(k - 1) Then
                    str = str & temp(i)
                    ret(k - 1) = str
                End If
            End If
            str = ""
            'save all operators of equation ("+", "-", "LOG", "SIN", etc. )
            If arr.Contains(temp(i)) = True Then
                ReDim Preserve ret(k)
                ret(k) = temp(i)
                k = k + 1
            End If
        Next i
        split_str_with_numbers = ret
    End Function
    Function remove_brackets_if_needed(eq() As String)
        Dim chk As Boolean
        'if a=((b+c)) then chk=true
        chk = check_brackets_to_be_removed(eq)
        If chk = True Then
            Do
                'a=(b+c))
                eq = reduce_arr(eq, 2)
                'a=(b+c)
                ReDim Preserve eq(UBound(eq) - 1)
                chk = check_brackets_to_be_removed(eq)
                'a=((b+c)) -> a=(b+c) -> a=b+c
            Loop Until chk = False
        End If
        remove_brackets_if_needed = eq
    End Function
    Function check_brackets_to_be_removed(eq() As String)
        Dim i As Integer, n As Integer, chk As Boolean
        If UBound(eq) > 2 Then
            If eq(2) = "(" And eq(1) = "=" Then
                n = 1 : i = 2
                Do
                    i = i + 1
                    If eq(i) = "(" Then
                        n = n + 1
                    ElseIf eq(i) = ")" Then
                        n = n - 1
                    End If
                Loop Until n = 0
            End If
            'if bracket after "=" & at end then chk=true example : a=(b+c)
            If i = UBound(eq) Then
                chk = True
            End If
        End If
        check_brackets_to_be_removed = chk
    End Function
    Function reduce_arr(arr() As String, i As Integer)
        Dim ret() As String, num As Integer
        ReDim ret(UBound(arr))
        arr.CopyTo(ret, 0)
        num = i
        If i <> UBound(arr) Then
            Do
                ret(num) = ret(num + 1)
                num = num + 1
            Loop Until num = UBound(ret)
            ReDim Preserve ret(UBound(ret) - 1)
        ElseIf i = UBound(arr) Then
            ReDim Preserve ret(UBound(ret) - 1)
        End If
        reduce_arr = ret
    End Function
    Function remove_brackets_single_variable(eq() As String)
        Dim chk As Boolean
        Dim i As Integer, k As Integer
        Dim ret() As String, temp() As String, v() As String, o() As String
        Dim arr = New String() {"LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        Do
            chk = False
            For i = 0 To UBound(ret)
                If ret(i) = ")" Then
                    temp = get_eqn_part_before_closing_bracket(ret, i)
                    k = i - UBound(temp)
                    v = get_operands2(temp)
                    o = get_operators2(temp)
                    'check if no. of varibles(nov) in bracket=1 and no. of operators(noo)=0 eq.: (b) nov=1("b") & noo=0
                    If UBound(v) = 0 And Len(Join(o)) = 0 Then
                        If k - 1 >= 0 Then
                            'check if brackets are preceded by math function like "SIN", "LOG", "LN", etc.
                            'if not chk=true
                            If arr.Contains(ret(k - 1)) = False Or ret(k - 1) = "*" Then
                                chk = True
                            End If
                            'if brackets in beginning of equation : (a)=b+c -> chk=true
                        ElseIf k - 1 <= 0 Then
                            chk = True
                        End If
                        If chk = True Then
                            'a=(b)+c -> a=b)+c
                            ret = reduce_arr(ret, k)
                            'a=b)+c -> a=b+c
                            ret = reduce_arr(ret, k + 1)
                            Exit For
                        End If
                    End If
                End If
            Next i
        Loop Until chk = False
        remove_brackets_single_variable = ret
    End Function
    Function add_brackets_negative_variables(eq() As String)
        Dim v() As String
        Dim i As Integer, t1 As Integer, t2 As Integer, n As Integer, m As Integer
        Dim arr = New String() {"+", "-", "*", "/", "^", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        v = get_operands2(eq)
        i = 0
        Do
            If eq(i) = "-" And eq(i) <> "*" And i - 1 > 0 And i + 1 <= UBound(eq) Then
                If arr.Contains(eq(i - 1)) = True Then
                    If eq(i + 1) = "(" Or v.Contains(eq(i + 1)) = True Then
                        If eq(i + 1) = "(" Then
                            n = 1 : m = i + 1
                            Do
                                m = m + 1
                                If eq(m) = "(" Then
                                    n = n + 1
                                ElseIf eq(m) = ")" Then
                                    n = n - 1
                                End If
                            Loop Until n = 0
                        ElseIf v.Contains(eq(i + 1)) = True Then
                            m = i + 1
                        End If
                        t1 = UBound(eq)
                        eq = add_element_to_array(eq, i, "(")
                        eq = add_element_to_array(eq, m + 2, ")")
                        t2 = UBound(eq)
                        i = i - (t1 - t2)
                    End If
                End If
            End If
            i = i + 1
        Loop Until i > UBound(eq)
        add_brackets_negative_variables = eq
    End Function
    Function get_eqn_part_before_closing_bracket(eq() As String, i As Integer)
        Dim k As Integer, m As Integer, n As Integer
        Dim ret() As String
        m = 1 : n = 0 : k = i
        Do
            ReDim Preserve ret(n)
            ret(n) = eq(k)
            n = n + 1
            k = k - 1
            If eq(k) = "(" Then
                m = m - 1
            ElseIf eq(k) = ")" Then
                m = m + 1
            End If
        Loop Until m = 0
        ReDim Preserve ret(n)
        ret(n) = "("
        ret = flip_str(ret)
        get_eqn_part_before_closing_bracket = ret
    End Function
    Function flip_str(str() As String)
        Dim i As Integer, k As Integer
        Dim temp() As String
        ReDim temp(UBound(str))
        k = UBound(str)
        For i = 0 To UBound(str)
            temp(i) = str(k)
            k = k - 1
        Next i
        flip_str = temp
    End Function
    Function get_operands2(t() As String)
        Dim temp() As String
        Dim i As Integer, k As Integer : k = 0
        Dim arr = New String() {"+", "-", "*", "/", "^", "(", ")", Chr(212), Chr(238), Chr(234), Chr(241), Chr(165), "=", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        For i = 0 To UBound(t)
            If arr.Contains(t(i)) = False And t(i) <> "*" Then
                'If arr.Contains(t(k - 1)) = True Then
                ReDim Preserve temp(k)
                temp(k) = t(i)
                k = k + 1
            End If
        Next i
        get_operands2 = temp
    End Function
    Function get_operators2(t() As String)
        Dim temp() As String
        Dim i As Integer, k As Integer : k = 0
        Dim arr = New String() {"+", "-", "*", "/", "^", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        For i = 0 To UBound(t)
            If arr.Contains(t(i)) = True Then
                ReDim Preserve temp(k)
                temp(k) = t(i)
                k = k + 1
            ElseIf Mid$(t(i), 1, 1) = "^" Then
                ReDim Preserve temp(k)
                temp(k) = "^"
                k = k + 1
            End If
        Next i
        get_operators2 = temp
    End Function
    Function add_brackets_operands(eq() As String)
        Dim i As Integer
        Dim arr = New String() {"LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        i = 0
        Do
            If arr.Contains(eq(i)) = True And eq(i) <> "*" Then
                'SINb -> SIN(b)
                If eq(i + 1) <> "(" Then
                    'SINb -> SIN(b
                    eq = add_element_to_array(eq, i + 1, "(")
                    'SIN(b -> SIN(b)
                    eq = add_element_to_array(eq, i + 3, ")")
                End If
            End If
            i = i + 1
        Loop Until i = UBound(eq)
        add_brackets_operands = eq
    End Function
    Function add_element_to_array(arr() As String, k As Integer, str As String)
        Dim i As Integer, n As Integer
        Dim ret() As String
        ReDim ret(UBound(arr))
        arr.CopyTo(ret, 0)
        ReDim Preserve ret(k)
        ret(k) = str
        If k <= UBound(arr) Then
            n = UBound(ret) + 1
            For i = k To UBound(arr)
                ReDim Preserve ret(n)
                ret(n) = arr(i)
                n = n + 1
            Next i
        End If
        add_element_to_array = ret
    End Function
    Function check_split_str(eq() As String)
        Dim teq() As String, neq() As String
        Dim i As Integer, l As Integer, t1 As Integer, t2 As Integer
        i = 0
        Do
            If eq(i) = ")" Then
                teq = get_eqn_part_before_closing_bracket(eq, i)
                l = i - UBound(teq) - 1
                'a*b+c-> (a*b)+c
                neq = check_brackets_needed(teq)
                If join_arr(neq) <> join_arr(teq) Then
                    t1 = UBound(eq)
                    'SIN(b*c+d) -> SIN((b*c)+d)
                    eq = replace_array_part_with_new_arr(eq, neq, l, i + 1)
                    t2 = UBound(eq)
                    i = i - (t1 - t2)
                End If
            End If
            i = i + 1
        Loop Until i > UBound(eq)
        'a=SIN(b)+c -> a=(SIN(b))+c
        eq = check_brackets_needed(eq)
        check_split_str = eq
    End Function
    Function join_arr(arr() As String)
        Dim ret As String
        Dim i As Integer
        For i = 0 To UBound(arr)
            ret = ret & arr(i)
        Next i
        join_arr = ret
    End Function
    Function replace_array_part_with_new_arr(arr() As String, new_arr() As String, p1 As Integer, p2 As Integer)
        Dim i As Integer, k As Integer
        Dim ret() As String
        ReDim ret(UBound(arr)) : arr.CopyTo(ret, 0)
        '(b*3)^2 : p1=2 : p2=7 : new_arr=(b^2)*6 -> b-A1+2
        If p1 = -1 Then
            ReDim ret(0)
        Else
            ReDim Preserve ret(p1)
        End If
        ret = add_arr1_to_arr(ret, new_arr)
        k = UBound(ret) + 1
        For i = p2 To UBound(arr)
            ReDim Preserve ret(k)
            ret(k) = arr(i)
            k = k + 1
        Next i
        replace_array_part_with_new_arr = ret
    End Function
    Function add_arr1_to_arr(arr() As String, arr1() As String)
        Dim i As Integer, k As Integer
        Dim ret() As String
        If Len(Join(arr)) > 0 Then
            ReDim ret(UBound(arr)) : arr.CopyTo(ret, 0)
            ReDim Preserve ret(arr1.Length + ret.Length - 1)
            k = UBound(arr) + 1
            For i = 0 To UBound(arr1)
                ret(k) = arr1(i)
                k = k + 1
            Next i
        Else
            ReDim ret(UBound(arr1)) : arr1.CopyTo(ret, 0)
        End If
        add_arr1_to_arr = ret
    End Function
    Function check_brackets_needed(eqn() As String)
        Dim brkeq() As String, chk As Boolean, teqo() As String, teq() As String, temp() As String
        Dim chkb As Boolean
        'remove brackets at start and end
        ReDim brkeq(UBound(eqn))
        eqn.CopyTo(brkeq, 0)
        chkb = check_brackets_start_end(brkeq)
        If chkb = True Then
            brkeq = reduce_arr(brkeq, 0)
            ReDim Preserve brkeq(UBound(brkeq) - 1)
        End If
        teq = get_compressed_equation(brkeq)
        teqo = get_operators2(teq)
        If Len(Join(teqo)) > 0 Then
            Do
                chk = check_operators(teqo)
                'add brackets if needed : a=b*c+d -> a=(b*c)+d
                temp = identify_operand_and_add_bracket(teq, chk)
                'convert tmpeq to eq : a=z1+d -> a=(b*c)+d
                brkeq = disassemble_tempeq(brkeq, temp)
                teq = get_compressed_equation(brkeq)
                teqo = get_operators2(teq)
            Loop Until chk = True
        End If
        If chkb = True Then
            brkeq = add_brackets_start_end(brkeq)
        End If
        check_brackets_needed = brkeq
    End Function
    Function check_brackets_start_end(eq() As String)
        Dim chk As Boolean
        Dim i As Integer, n As Integer
        If eq(0) = "(" Then
            n = 1 : i = 0
            Do
                i = i + 1
                If eq(i) = "(" Then
                    n = n + 1
                ElseIf eq(i) = ")" Then
                    n = n - 1
                End If
            Loop Until n = 0
            If i = UBound(eq) Then
                chk = True
            End If
        End If
        check_brackets_start_end = chk
    End Function
    Function get_compressed_equation(t() As String)
        Dim temp1() As String
        Dim i As Integer, n As Integer, m As Integer, k As Integer
        k = 0 : m = 0
        For i = 0 To UBound(t)
            If t(i) = "(" Then
                n = 1
                Do
                    i = i + 1
                    If t(i) = "(" Then
                        n = n + 1
                    ElseIf t(i) = ")" Then
                        n = n - 1
                    End If
                Loop Until n = 0
                ReDim Preserve temp1(k)
                temp1(k) = Chr(158) & m
                k = k + 1 : m = m + 1
            Else
                ReDim Preserve temp1(k)
                temp1(k) = t(i)
                k = k + 1
            End If
        Next i
        get_compressed_equation = temp1
    End Function
    Function check_operators(t() As String)
        Dim temp() As String
        Dim k As Integer
        Dim chk As Boolean
        k = 0 : chk = True
        For i = 0 To UBound(t)
            If t(i) <> "=" Then
                ReDim Preserve temp(k)
                If t(i) = "-" Then
                    temp(k) = "+"
                Else
                    temp(k) = t(i)
                End If
                k = k + 1
            End If
        Next i
        For i = 0 To UBound(temp) - 1
            If temp(0) <> temp(i + 1) Then
                chk = False
            End If
        Next i
        If chk = True And UBound(temp) > 0 Then
            If temp(0) = "/" Or temp(0) = "^" Then
                chk = False
            End If
        End If
        check_operators = chk
    End Function
    Function get_operators(t() As String)
        Dim temp() As String
        Dim i As Integer, k As Integer : k = 0
        Dim arr = New String() {"+", "-", "*", "/", "(", ")", Chr(212), Chr(238), Chr(234), Chr(241), Chr(165), "=", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        For i = 0 To UBound(t)
            If arr.Contains(t(i)) = True Then
                ReDim Preserve temp(k)
                temp(k) = t(i)
                k = k + 1
            ElseIf Mid$(t(i), 1, 1) = "^" Then
                ReDim Preserve temp(k)
                temp(k) = "^"
                k = k + 1
            End If
        Next i
        get_operators = temp
    End Function
    Function add_brackets_start_end(arr() As String)
        Dim k As Integer
        ReDim Preserve arr(UBound(arr) + 2)
        k = UBound(arr)
        Do
            arr(k) = arr(k - 1)
            k = k - 1
        Loop Until k = 0
        arr(0) = "("
        arr(UBound(arr)) = ")"
        add_brackets_start_end = arr
    End Function
    Function disassemble_tempeq(eq() As String, tmpeq() As String)
        Dim ret() As String, arr(,) As String, temp() As String
        Dim i As Integer, j As Integer, k As Integer, n As Integer
        arr = get_two_dim_array(eq)
        k = 0
        For i = 0 To UBound(tmpeq)
            If Mid$(tmpeq(i), 1, 1) = Chr(158) Then
                n = Mid$(tmpeq(i), 2)
                temp = get_one_dim_array(arr, n)
                For j = 0 To UBound(temp)
                    ReDim Preserve ret(k)
                    ret(k) = temp(j)
                    k = k + 1
                Next j
            Else
                ReDim Preserve ret(k)
                ret(k) = tmpeq(i)
                k = k + 1
            End If
        Next i
        disassemble_tempeq = ret
    End Function
    Function get_two_dim_array(arr() As String)
        Dim i As Integer
        Dim ret(,) As String, temp() As String
        For i = 0 To UBound(arr)
            If arr(i) = "(" Then
                temp = get_eqn_part_after_opening_bracket(arr, i)
                i = i + UBound(temp) - 1
                ret = add_two_dim_array(ret, temp)
            End If
        Next i
        get_two_dim_array = ret
    End Function
    Function get_eqn_part_after_opening_bracket(eq() As String, i As Integer)
        Dim k As Integer, n As Integer, m As Integer
        Dim ret() As String
        k = i : n = 1 : m = 0
        Do
            ReDim Preserve ret(m)
            ret(m) = eq(k)
            k = k + 1
            m = m + 1
            If eq(k) = "(" Then
                n = n + 1
            ElseIf eq(k) = ")" Then
                n = n - 1
            End If
        Loop Until n = 0
        ReDim Preserve ret(m)
        ret(m) = eq(k)
        get_eqn_part_after_opening_bracket = ret
    End Function
    Function add_two_dim_array(arr(,) As String, eqns() As String)
        Dim ret(,) As String
        Dim i As Integer, j As Integer, m As Integer, n As Integer, k As Integer, chk As Boolean
        chk = IsArrayEmpty(arr)
        If chk = True Then
            m = UBound(eqns) : n = 1
            ReDim ret(0, m)
            For i = 0 To m
                ret(0, i) = eqns(i)
            Next i
        Else
            If UBound(eqns) > UBound(arr, 2) Then
                'if new row dimension > two dimensional array then make column size of new two dim array equal to new row dimesion
                'arr=((EQ1,EQ2)) : new_row=(EQ3,EQ4,EQ5) -> arr=((EQ1,EQ2,""),(EQ3,EQ4,EQ5))
                m = UBound(eqns)
            Else
                'add new row to two dimensional array and add new row
                'arr=((EQ1,EQ2)) : new_row=(EQ3,EQ4) -> arr=((EQ1,EQ2),(EQ3,EQ4))
                m = UBound(arr, 2)
            End If
            n = UBound(arr, 1) + 1
            ReDim ret(n, m)
            For i = 0 To UBound(arr, 1)
                For j = 0 To UBound(arr, 2)
                    ret(i, j) = arr(i, j)
                Next j
            Next i
            For k = 0 To UBound(eqns)
                ret(i, k) = eqns(k)
            Next k
        End If
        add_two_dim_array = ret
    End Function
    Function IsArrayEmpty(arr As Object) As Boolean
        Dim k As Long
        On Error Resume Next
        k = UBound(arr, 1)
        If Err.Number = 0 Then
            IsArrayEmpty = False
        Else
            IsArrayEmpty = True
        End If
    End Function
    Function get_one_dim_array(arr(,) As String, n As Integer)
        Dim ret() As String
        Dim i As Integer, k As Integer
        k = 0
        For i = 0 To UBound(arr, 2)
            If arr(n, i) <> "" Then
                ReDim Preserve ret(k)
                ret(k) = arr(n, i)
                k = k + 1
            End If
        Next i
        get_one_dim_array = ret
    End Function
    Function identify_operand_and_add_bracket(eqn() As String, chk As Boolean)
        Dim tempf As String, chkf As Boolean, p1 As Integer, p2 As Integer
        Dim arr = New String() {"SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT", "LOG", "LN", "EXP"}
        'identify the operand and add brackets accordingly
        If chk = False Then
            For i = 0 To UBound(eqn)
                If arr.Contains(eqn(i)) = True And eqn(i) <> "*" Then
                    tempf = eqn(i)
                    chkf = True
                    Exit For
                End If
            Next i
            If chkf = True Then
                'eq=SIN(a)+b -> eq=(SIN(a))+b
                p1 = Array.IndexOf(eqn, tempf)
                p2 = p1 + 3
            ElseIf eqn.Contains("^") = True Then
                'eq=a^b+c -> eq=(a^b)+c
                p1 = Array.IndexOf(eqn, "^") - 1
                p2 = p1 + 4
            ElseIf eqn.Contains("/") = True Then
                'eq=a/b+c -> eq=(a/b)+c
                p1 = Array.IndexOf(eqn, "/") - 1
                p2 = p1 + 4
            ElseIf eqn.Contains("*") = True Then
                'eq=a*b*c+d -> eq=(a*b*c)+d
                p1 = Array.IndexOf(eqn, "*") - 1
                p2 = p1
                Do
                    p2 = p2 + 2
                    If p2 + 2 > UBound(eqn) Then
                        Exit Do
                    End If
                Loop Until p2 > UBound(eqn) Or eqn(p2 + 1) <> "*"
                p2 = p2 + 2
            End If
            eqn = add_element_to_array(eqn, p1, "(")
            eqn = add_element_to_array(eqn, p2, ")")
        End If
        identify_operand_and_add_bracket = eqn
    End Function
    Function display_two_dim_array(arr(,) As String)
        Dim disp() As String, temp() As String
        Dim i As Integer, k As Integer
        Dim str As String
        k = 0
        For i = 0 To UBound(arr, 1)
            temp = get_one_dim_array(arr, i)
            str = Join(temp, ", ")
            ReDim Preserve disp(k)
            disp(k) = str
            k = k + 1
        Next i
        MsgBox(Join(disp, vbCr))
    End Function
End Module
Module Module2
    Function eqn_typ_1(eqn() As String, inputv() As String)
        Dim feq() As String, n_input() As String, chkf() As String, ret() As String
        feq = eqn : n_input = inputv
        Do
            chkf = Chk_Feq(feq, n_input)
            n_input = rearrange_inputv(feq, chkf, n_input)
            feq = chkf
        Loop Until chkf(1) = "="
        ret = disp(chkf, n_input)
        eqn_typ_1 = ret
    End Function
    Function get_position_unknown_var(eq() As String, inputv() As String)
        'eq "a=(b^c)+d" : inputv=(2,3,x,4) : v=(a,(,b,c,),d)
        Dim temp() As String, convar() As String, v() As String
        Dim k As Integer, n As Integer, pos As Integer
        v = get_operands(eq)
        k = 0 : n = 0
        For i = 0 To UBound(v)
            ReDim Preserve temp(k)
            'temp=(2,(,3,x,),4)
            If v(i) = "(" Or v(i) = ")" Then
                temp(k) = v(i)
            Else
                temp(k) = inputv(n)
                n = n + 1
            End If
            k = k + 1
        Next i
        '(2,(,3,x,),4) -> (y,x,y) : (10,(5,2),(3,x),8) -> (y,y,x,y)
        k = 0
        For i = 0 To UBound(temp)
            ReDim Preserve convar(k)
            If temp(i) = "x" Then
                convar(k) = "x"
            ElseIf temp(i) = "(" Then
                n = 1
                Do
                    i = i + 1
                    If temp(i) = "x" Then
                        convar(k) = "x"
                    ElseIf temp(i) = "(" Then
                        n = n + 1
                    ElseIf temp(i) = ")" Then
                        n = n - 1
                    End If
                Loop Until n = 0
                If convar(k) <> "x" Then
                    convar(k) = "y"
                End If
            Else
                convar(k) = "y"
            End If
            k = k + 1
        Next i
        pos = Array.IndexOf(convar, "x")
        get_position_unknown_var = pos
    End Function
    Function get_operands(t() As String)
        Dim temp() As String
        Dim i As Integer, k As Integer : k = 0
        Dim arr = New String() {"+", "-", "*", "/", "^", Chr(212), Chr(238), Chr(234), Chr(241), Chr(165), "=", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        For i = 0 To UBound(t)
            If arr.Contains(t(i)) = False Then
                ReDim Preserve temp(k)
                temp(k) = t(i)
                k = k + 1
            End If
        Next i
        get_operands = temp
    End Function
    Function Chk_Feq(t1() As String, inputv() As String)
        Dim tempeq() As String, feq() As String, neq() As String
        Dim operand As String
        Dim v() As String, o() As String
        Dim k As Integer, p As Integer
        p = get_position_unknown_var(t1, inputv)
        tempeq = get_compressed_equation(t1)
        If tempeq(1) <> "=" Then
            'a+b-c=z1 -> z1=a+b-c
            tempeq = shuffle(tempeq)
            p = p + 1
        End If
        o = get_operators(tempeq)
        'a=b+c-d -> v=(a,b,c,d) unkwn=b nv=(b,a,c,d) : b+c+d=z1 -> v=(b,c,d,z1) unkwn=c nv=(c,z1,c,d)
        If o(0) <> "=" Then
            operand = o(0)
        Else
            operand = o(1)
        End If
        v = get_operands2(tempeq)
        If operand <> "/" Then
            'a=b+c-d : b is unkwn -> v=(a,b,c,d) p=2 -> ret=(b,a,c,d)
            v = rearrange(v, p)
        End If
        Select Case operand
            Case "+", "-"
                neq = new_eq_add(tempeq, v, o)
            Case "*"
                neq = new_eq_mul(tempeq, v, o)
            Case "/"
                neq = new_eq_div(tempeq, v, o, p)
            Case "^"
                neq = new_eq_pow(tempeq, v, o, p)
            Case "LOG"
                neq = new_eq_log(tempeq, v, o)
            Case "LN"
                neq = new_eq_ln(tempeq, v, o)
            Case "EXP"
                neq = new_eq_exp(v)
            Case "SQRT"
                neq = new_eq_sqrt(v)
            Case "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN"
                neq = new_eq_trig(tempeq, v, o)
        End Select
        'a=(b*c)+d b unkwn -> a=z1+d -> z1=a-d -> (b*c)=a-d -> b*c=a-d -> b*c=(a-d) -> b*c=z1 -> z1=b*c -> b=z1/c -> b=(a-d)/c -> b=((a-d)/c)
        feq = disassemble_tempeq(t1, neq)
        '(a-c)=d*e -> a-c=d*e
        k = Array.IndexOf(feq, "=")
        If feq(0) = "(" And feq(k - 1) = ")" Then
            feq = reduce_arr(feq, 0)
            feq = reduce_arr(feq, k - 2)
            k = k - 2
        End If
        'a-c=d*e -> a-c=(d*e)
        feq = add_element_to_array(feq, k + 1, "(")
        feq = add_element_to_array(feq, UBound(feq) + 1, ")")
        If feq(0) = "-" Then
            '-b=(a/(c*d)) -> b=(-(a/(c*d)))
            k = k - 1
            feq = reduce_arr(feq, 0)
            feq = add_element_to_array(feq, k + 1, "(")
            feq = add_element_to_array(feq, k + 2, "-")
            feq = add_element_to_array(feq, UBound(feq) + 1, ")")
            k = Array.IndexOf(feq, "=")
            If feq(0) = "(" And feq(k - 1) = ")" Then
                feq = reduce_arr(feq, 0)
                feq = reduce_arr(feq, k - 2)
            End If
        End If
        Chk_Feq = feq
    End Function
    Function shuffle(eq() As String)
        Dim ret() As String
        Dim i As Integer, j As Integer, k As Integer
        ReDim ret(UBound(eq))
        'b-c+d=z1 -> z1 -> z1= -> z1=b-c+d
        k = UBound(eq) : i = 0
        Do
            ret(i) = eq(k)
            k = k - 1
            i = i + 1
        Loop Until eq(k) = "="
        ret(i) = eq(k)
        i = i + 1
        j = 0
        Do
            ret(i) = eq(j)
            i = i + 1
            j = j + 1
        Loop Until eq(j) = "="
        shuffle = ret
    End Function
    Function rearrange(v() As String, p As Integer)
        Dim ret() As String
        Dim k As Integer
        ReDim ret(UBound(v))
        'a=b+c-d : b is unkwn -> v=(a,b,c,d) p=1 -> ret=(b,a,c,d)
        ret(0) = v(p) : ret(1) = v(0) : k = 2
        For i = 0 To UBound(v)
            If v(i) <> ret(0) And v(i) <> ret(1) Then
                ret(k) = v(i)
                k = k + 1
            End If
        Next i
        rearrange = ret
    End Function
    Function new_eq_add(eq() As String, v() As String, o() As String)
        Dim no() As String, temp() As String
        Dim i As Integer
        If UBound(v) = 1 And UBound(o) = 1 And o(1) = "-" Then
            'a=-(b+c) -> a=-z1 -> z1=a -> z1=(-a) -> (b+c)=(-a) -> b+c=(-a)
            ReDim no(0)
            no(0) = "="
        Else
            Dim c As New Collection
            For i = 0 To UBound(eq)
                If eq(i) = "=" Then
                    c.Add("=")
                    'a=b+c-d is equavilant to a=+b+c-d thus "+" before b changes to "-"
                    If eq(i + 1) <> "+" And eq(i + 1) <> "-" And eq(i + 1) <> v(0) Then
                        c.Add("-")
                    End If
                ElseIf i <> UBound(eq) Then
                    'a=b-c+d & "c" is unkvar then do not consider the "-" sign before "c"
                    If eq(i + 1) = v(0) Then
                        i = i + 1
                    End If
                End If
                'a=b+c-d : c is unkvar : o=("=","+","-") -> no=("=","-","+")
                If eq(i) = "-" Then
                    c.Add("+")
                ElseIf eq(i) = "+" Then
                    c.Add("-")
                End If
            Next i
            no = convert_collection_to_array(c)
        End If
        'v=(b,a,c) : no=(=,-) -> combine = (b,=,a,-,c)
        temp = combine(v, no)
        'eq="a=b-c+d : c is unkwn  -> c=a-b-d -> c=-(a-b-d)
        If eq(Array.IndexOf(eq, v(0)) - 1) <> "+" And eq(Array.IndexOf(eq, v(0)) - 1) <> "=" Then
            temp = add_element_to_array(temp, 2, "-")
            temp = add_element_to_array(temp, 3, "(")
            temp = add_element_to_array(temp, UBound(temp) + 1, ")")
        End If
        new_eq_add = temp
    End Function
    Function new_eq_mul(eq() As String, nv() As String, no() As String)
        Dim temp() As String
        Dim i As Integer
        'a=b*c : v=(a,b,c) : b unkwn -> nv=(b,a,c)
        '("=","*","*","*") -> ("=","/","*","*")
        no(1) = "/"
        For i = 2 To UBound(no)
            no(i) = "*"
        Next i
        'a=b*c*d : unkvar = c : nv = (c,a,b,d) : no = ("=","/","*") -> c=a/b*d -> c=a/(b*d)
        temp = combine(nv, no)
        If UBound(nv) > 2 Then
            temp = add_element_to_array(temp, 4, "(")
            temp = add_element_to_array(temp, UBound(temp) + 1, ")")
        End If
        new_eq_mul = temp
    End Function
    Function new_eq_div(eq() As String, v() As String, no() As String, p As Integer)
        Dim temp() As String, nv() As String
        ReDim nv(UBound(v))
        Select Case p
            'a=b/c
            Case 1
                'b unkwn -> nv=(b,a,c)
                nv(0) = v(1) : nv(1) = v(0) : nv(2) = v(2)
                'no=("=","*")
                no(0) = "=" : no(1) = "*"
            Case 2
                'c unkwn -> nv=(c,b,a)
                nv(0) = v(2) : nv(1) = v(1) : nv(2) = v(0)
                'no=("=","/")
                no(0) = "=" : no(1) = "/"
        End Select
        'a=b/c : b unkwn -> b=a*c : c unkwn -> c=b/a
        temp = combine(nv, no)
        new_eq_div = temp
    End Function
    Function new_eq_pow(eq() As String, nv() As String, no() As String, p As Integer)
        Dim temp() As String, ieq() As String
        Dim i As Integer, k As Integer
        Select Case p
            Case 1
                'no=("=","i")
                no(0) = "=" : no(1) = Chr(238)
                'a=b^c : b unkwn -> b=aic : i=^(1/
                temp = combine(nv, no)
            Case 2
                'no=("=","/")
                no(0) = "=" : no(1) = "/"
                'a=b^c : c unkwn -> c=a/b -> c=LOG(a)/LOG(b)
                ieq = combine(nv, no)
                k = 0 : ReDim temp(UBound(ieq) + 6)
                For i = 0 To UBound(temp)
                    If i = 2 Or i = 7 Then
                        temp(i) = "LOG"
                        temp(i + 1) = "("
                        temp(i + 2) = ieq(k)
                        temp(i + 3) = ")"
                        i = i + 3
                        k = k + 1
                    Else
                        temp(i) = ieq(k)
                        k = k + 1
                    End If
                Next i
        End Select
        new_eq_pow = temp
    End Function
    Function new_eq_log(eq() As String, nv() As String, no() As String)
        Dim temp() As String
        'no=("=","O")
        no(0) = "="
        no(1) = Chr(212)
        'a=log(b) b is unkwn -> b=O(a) O=10^
        temp = combine_2(nv, no)
        new_eq_log = temp
    End Function
    Function new_eq_ln(eq() As String, nv() As String, no() As String)
        Dim temp() As String
        'no=("=","n")
        no(0) = "="
        no(1) = Chr(241)
        'a=LN(b) b is unkwn -> b=n(a) n=2.71828^
        temp = combine_2(nv, no)
        new_eq_ln = temp
    End Function
    Function new_eq_exp(nv() As String)
        'a=EXP(b) b is unkwn -> b=(LOG(a))/(LOG(e)) e=2.71828....
        Dim ret = New String() {nv(0), "=", "(", "LOG", "(", nv(1), ")", ")", "/", "(", "LOG", "(", Chr(234), ")", ")"}
        new_eq_exp = ret
    End Function
    Function new_eq_sqrt(nv() As String)
        'a=sqrt(b) b is unkwn -> b=(a)Y Y=^2
        Dim ret = New String() {nv(0), "=", "(", nv(1), ")", Chr(165)}
        new_eq_sqrt = ret
    End Function
    Function new_eq_trig(eq() As String, nv() As String, o() As String)
        Dim no() As String, temp() As String
        Dim pos As Integer
        Dim arr1 = New String() {"SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN"}
        Dim arr2 = New String() {"ASIN", "ACOS", "ATAN", "SIN", "COS", "TAN"}
        ReDim no(UBound(o))
        'o(2)=SIN -> no(2)=ASIN : o(2)=TAN -> no(2)=ATAN
        pos = Array.IndexOf(arr1, o(1))
        no(0) = "="
        no(1) = arr2(pos)
        'a=SIN(b) b unkwn -> b=ASIN(a)
        temp = combine_2(nv, no)
        new_eq_trig = temp
    End Function
    Function convert_collection_to_array(c As Collection)
        Dim ret() As String
        Dim i As Integer, k As Integer
        If c.Count > 1 Then
            ReDim ret((c.Count) - 1)
            For i = 1 To c.Count
                ret(k) = c.Item(i)
                k = k + 1
            Next i
        End If
        convert_collection_to_array = ret
    End Function
    Function add_array_to_collection(ByVal c As Collection, arr() As String) As Collection
        Dim i As Integer
        For i = 0 To UBound(arr)
            c.Add(arr(i))
        Next i
        add_array_to_collection = c
    End Function
    Function combine(v() As String, o() As String)
        Dim i As Integer, l As Integer, k As Integer
        Dim ret() As String
        k = v.Length + o.Length - 1
        ReDim ret(k)
        'v=(a,b,c) -> ret=(a, ,b, ,c)
        l = 0
        For i = 0 To UBound(v)
            ret(l) = v(i)
            l = l + 2
        Next i
        'o=(=,+) -> ret=(a,=,b,+,c)
        l = 1
        For i = 0 To UBound(o)
            ret(l) = o(i)
            l = l + 2
        Next i
        combine = ret
    End Function
    Function combine_2(v() As String, o() As String)
        Dim temp() As String
        ReDim temp(5)
        'v=(b,a) : o=("=","ACOS") -> b=ACOS(a)
        temp(0) = v(0)
        temp(1) = o(0)
        temp(2) = o(1)
        temp(3) = "("
        temp(4) = v(1)
        temp(5) = ")"
        combine_2 = temp
    End Function
    Function rearrange_inputv(eq() As String, neq() As String, inputv() As String)
        Dim v() As String, nv() As String, ret() As String
        Dim p As Integer, n As Integer
        v = get_operands2(eq)
        nv = get_operands2(neq)
        n = 0
        For i = 0 To UBound(v)
            p = Array.IndexOf(v, nv(i))
            ReDim Preserve ret(n)
            ret(n) = inputv(p)
            n = n + 1
        Next i
        rearrange_inputv = ret
    End Function
    Function disp(eq() As String, rvin() As String)
        Dim v() As String, e() As String, deq() As String, de() As String, teq() As String, temp() As String, ret() As String
        Dim disp1 As String, disp2 As String
        Dim i As Integer, k As Integer
        Dim coll As New Collection
        'ret = disp(chkf, avo, inputv, nov, pass, des, unit)
        v = get_operands2(eq)
        teq = eq
        For i = 0 To UBound(teq)
            'a=b^c : b unkwn -> b=aic : i=^(1/ -> b=a^(1/c)
            If teq(i) = Chr(238) Then
                coll.Add("^") : coll.Add("(") : coll.Add("1") : coll.Add("/")
                temp = get_eqn_part_after_opening_bracket(teq, i + 1)
                coll = add_array_to_collection(coll, temp)
                coll.Add(")")
                i = i + temp.Length
                'a=log(b) b is unkwn -> b=O(a) O=10^ -> b=10^(a)
            ElseIf teq(i) = Chr(212) Then
                coll.Add("10 ^")
                'a=LN(b) b is unkwn -> b=n(a) n=2.71828^ -> b=2.71828^(a)
            ElseIf teq(i) = Chr(241) Then
                coll.Add("2.71828182845905 ^")
                'a=EXP(b) b is unkwn -> b=(LOG(a))/(LOG(e)) e=2.71828 -> b=(LOG(a))/(LOG(2.71828))
            ElseIf teq(i) = Chr(234) Then
                coll.Add("2.71828182845905")
            ElseIf teq(i) = Chr(165) Then
                coll.Add("^ 2")
            ElseIf teq(i) = "PI" Then
                coll.Add("3.14159265358979")
            Else
                coll.Add(teq(i))
            End If
        Next i
        eq = convert_collection_to_array(coll)
        'convert variable of constants to constant numbers a=b+C1 C1=3 -> a=b+3
        For i = 0 To UBound(eq)
            If Mid$(eq(i), 1, 1) = Chr(231) Then
                eq(i) = rvin(Array.IndexOf(v, eq(i)))
            End If
        Next i
        ReDim e(UBound(eq))
        eq.CopyTo(e, 0)
        disp1 = ""
        Dim arr = New String() {"+", "-", "*", "/", "^", "2.71828182845905", "3.14159265358979", "2.71828182845905 ^", "(", ")", "=", "10 ^", "^ 2", "^ (1 /", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        'replace variables with numbers a=b+c : b=2 & c=3 -> a=2+3
        For i = 0 To UBound(e)
            If v.Contains(e(i)) = True And e(i) <> "*" Then
                e(i) = rvin(Array.IndexOf(v, eq(i)))
            End If
        Next i
        ReDim deq(UBound(eq) - 2)
        ReDim de(UBound(eq) - 2)
        k = 0
        For i = 2 To UBound(eq)
            deq(k) = eq(i)
            de(k) = e(i)
            k = k + 1
        Next i
        '(COS(b))+(SIN(c)) -> COS(b) + SIN(c)
        disp1 = disp_neq(deq)
        disp2 = disp_neq(de)
        If InStr(disp1, "3.14159265358979") <> 0 Then
            disp1 = disp1.Replace("3.14159265358979", "PI")
            disp2 = disp2.Replace("3.14159265358979", "3.142")
        End If
        'SIN(a) -> SIN(a*PI()/180) : ASIN(a) -> ((ASIN(a))*180/PI())
        e = adjust_eqn_for_trignometric_functions(e)
        Dim result As String
        For i = 2 To UBound(e)
            result = result & e(i)
        Next i
        result = app.Evaluate(result)
        result = CStr(Round(CDbl(result), Decimal_Pt))
        ReDim ret(5)
        ret(0) = Variable_Description(Array.IndexOf(Variable_array, eq(0)))
        ret(1) = eq(0)
        ret(2) = disp1
        ret(3) = disp2
        ret(4) = result
        ret(5) = Unit(Array.IndexOf(Variable_array, eq(0)))
        disp = ret
    End Function
    Function disp_neq(eq() As String)
        Dim tempe() As String, ret As String
        Dim i As Integer, n As Integer, j As Integer
        Dim chk As Boolean
        Dim arr = New String() {"LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        chk = check_brackets_start_end(eq)
        If chk = True Then
            tempe = get_compressed_equation(eq)
            If UBound(tempe) = 1 Or UBound(tempe) = 0 Then
                eq = remove_brackets_start_end(eq)
            End If
        End If
        'a=(COS(b))+c -> a=COS(b))+c -> a=COS(b)+c
        i = 0
        Do
            If arr.Contains(eq(i)) = True And eq(i) <> "*" And i - 1 > 0 Then
                'a=(COS(b))+c
                If eq(i - 1) = "(" Then
                    j = i - 1 : n = 1
                    Do
                        j = j + 1
                        If eq(j) = "(" Then
                            n = n + 1
                        ElseIf eq(j) = ")" Then
                            n = n - 1
                        End If
                    Loop Until n = 0
                    'a=(COS(b))+c -> a=COS(b))+c
                    eq = reduce_arr(eq, i - 1)
                    'a=(COS(b))+c -> a=COS(b)+c
                    eq = reduce_arr(eq, j - 1)
                End If
            End If
            i = i + 1
        Loop Until i = UBound(eq)
        '10^(a) -> 10^a
        eq = remove_brackets_single_variable(eq)
        chk = check_brackets_start_end(eq)
        If chk = True Then
            eq = remove_brackets_start_end(eq)
        End If
        'add spaces between varaiables before displaying
        'b^c -> b ^ c : COS(b*c) -> COS(b * c)
        For i = 0 To (UBound(eq))
            If arr.Contains(eq(i)) = True And eq(i) <> "*" Then
                ret = ret & eq(i)
                ret = ret & eq(i + 1)
                i = i + 1
            ElseIf eq(i) = "(" Then
                ret = ret & eq(i)
            ElseIf eq(i) = ")" Then
                ret = Trim(ret)
                ret = ret & eq(i) & " "
            Else
                ret = ret & eq(i) & " "
            End If
        Next i
        disp_neq = ret
    End Function
    Function adjust_eqn_for_trignometric_functions(eq() As String)
        Dim ret() As String, temp() As String
        Dim i As Integer
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        i = 0
        Do
            If ret(i) = "SIN" Or ret(i) = "COS" Or ret(i) = "TAN" Then
                'SIN(f+1) -> (f+1) -> (f+1)*PI()/180 -> ((f+1)*PI()/180) -> SIN((f+1)*PI()/180)
                temp = get_eqn_part_after_opening_bracket(ret, i + 1)
                ReDim Preserve temp(UBound(temp) + 1)
                temp(UBound(temp)) = "*PI()/180"
                temp = add_brackets_start_end(temp)
                ret = replace_array_part_with_new_arr(ret, temp, i, i + UBound(temp) - 1)
                i = i + 2
            ElseIf ret(i) = "ASIN" Or ret(i) = "ACOS" Or ret(i) = "ATAN" Then
                'SIN(f+1) -> (f+1) -> (f+1)*PI()/180 -> ((f+1)*PI()/180) -> SIN((f+1)*PI()/180)
                temp = get_eqn_part_after_opening_bracket(ret, i + 1)
                temp = add_element_to_array(temp, 0, ret(i))
                temp = add_brackets_start_end(temp)
                temp = add_element_to_array(temp, UBound(temp) + 1, "*180/PI()")
                temp = add_brackets_start_end(temp)
                ret = replace_array_part_with_new_arr(ret, temp, i - 1, i + UBound(temp) - 4)
                i = i + 2
            End If
            i = i + 1
        Loop Until i > UBound(ret)
        adjust_eqn_for_trignometric_functions = ret
    End Function
End Module
Module Module3
    Public unkvar_rhs_division As Boolean
    Private simplified_equation_part1() As String
    Private simplified_equation_part2() As String
    Private func_start_point As Integer
    Private func_end_point As Integer
    Private Eqn_Part_With_Unknowns() As String
    Function eqn_typ_2(eqn() As String, inputv() As String, v() As String, unkvar As String)
        Dim redeq() As String, result() As String, common_part() As String, tmpeq() As String, neq() As String, nv() As String
        Dim n As Integer
        Dim chk As Boolean
        redeq = reduce_eqn(eqn, inputv, v, unkvar)
        n = Form1.number_of_instances(redeq, unkvar)
        If n > 1 Then
            common_part = get_part_of_eqn_with_multiple_instances(redeq, unkvar)
            'eq=(3*b)+b : str=b -> tmpeq=(3*f1)+f2
            tmpeq = get_temp_simplified_eqn(redeq, common_part)
            'tmpeq=(3*f)+f : str=b -> ret=b*((3*1)+1)
            chk = check_simplified_equation(tmpeq)
            If chk = True Then
                neq = get_final_simplified_eqn(tmpeq, common_part)
            End If
        Else
            chk = True
            neq = redeq
        End If
        'MsgBox(Join(redeq, "") & Environment.NewLine & Join(neq, ""))
        If chk = True Then
            nv = get_operands2(neq)
            inputv = rearrange_input(v, nv, inputv)
            n = Form1.number_of_instances(neq, unkvar)
            If n > 1 Then
                result = eqn_typ_2(neq, inputv, nv, unkvar)
            Else
                result = eqn_typ_1(neq, inputv)
            End If
        Else
            result = bisection_method(redeq, inputv, v, unkvar)
        End If
        eqn_typ_2 = result
    End Function
    Function bisection_method(eq() As String, inputv() As String, v() As String, uknvar As String)
        Dim f() As String
        Dim str As String, soln1 As Double, soln2 As Double, disp As String
        Dim a As Double, b As Double, c As Double, c1 As Double
        Dim n As Integer, p As Integer, chk As Boolean
        ReDim f(UBound(eq))
        eq.CopyTo(f, 0)
        'a=(b^2)+b : a=17 -> (b^2)+b -> ((b^2)+b) -> ((b^2)+b)-17
        f = reduce_arr(f, 0)
        f = reduce_arr(f, 0)
        If eq(0) <> Chr(220) Then
            f = add_brackets_start_end(f)
            ReDim Preserve f(UBound(f) + 2)
            p = Array.IndexOf(v, eq(0))
            If Sign(CDbl(inputv(p))) = -1 Then
                f(UBound(f) - 1) = "+"
                f(UBound(f)) = Mid$(inputv(p), 2)
            Else
                f(UBound(f) - 1) = "-"
                f(UBound(f)) = inputv(p)
            End If
        End If
        disp = join_arr(f)
        f = adjust_eqn_for_trignometric_functions(f)
        str = join_arr(f)
        a = 0 : b = 1
        Do
            If IsError(app.Evaluate(Replace(str, uknvar, a))) = True Then
                a = a + 1
                b = b + 1
            Else
                chk = True
            End If
        Loop Until chk = True
        chk = False
        Do
            soln1 = CDbl(app.Evaluate(Replace(str, uknvar, a)))
            soln2 = CDbl(app.Evaluate(Replace(str, uknvar, b)))
            If Sign(soln1) <> Sign(soln2) Then
                chk = True
            Else
                a = a + 1
                b = b + 1
            End If
        Loop Until chk = True
        n = 1 : c1 = 0
        Do
            c1 = Round(c, 9)
            c = (a + b) / 2
            soln1 = CDbl(app.Evaluate(Replace(str, uknvar, c)))
            soln2 = CDbl(app.Evaluate(Replace(str, uknvar, b)))
            If Sign(soln1) <> Sign(soln2) Then
                a = c
            Else
                b = c
            End If
            n = n + 1
        Loop Until c1 = Round(c, 9) Or soln1 = 0 Or n = 1000
        c = CStr(Round(CDbl(c), Decimal_Pt))
        Dim result = New String() {"Using bisection method ", Variable_Description(Array.IndexOf(Variable_array, uknvar)) & " " & uknvar, c, Unit(Array.IndexOf(Variable_array, uknvar))}
        bisection_method = result
    End Function
    Function rearrange_input(v() As String, nv() As String, inputv() As String)
        Dim ret() As String
        Dim i As Integer, p As Integer
        ReDim ret(UBound(nv))
        For i = 0 To UBound(nv)
            If v.Contains(nv(i)) = True Then
                p = Array.IndexOf(v, nv(i))
                ret(i) = inputv(p)
            ElseIf nv(i) = Chr(220) Then
                ret(i) = "0"
            Else
                ret(i) = nv(i)
            End If
        Next i
        rearrange_input = ret
    End Function
    Function get_final_simplified_eqn(eq() As String, common_part() As String)
        Dim ret() As String
        Dim k As Integer
        Dim c As New Collection
        If Len(Join(Eqn_Part_With_Unknowns)) > 0 Then
            'COS(f+f+3) : (f+f+3) -> part1 = f+f : part2 = +3
            Call get_final_smpeq_part1_part2(Eqn_Part_With_Unknowns)
            ReDim ret(0)
            ret(0) = "("
            k = 1
        Else
            'f-(f*3)+6 -> part1 = f-(f*3) : part2 = +6
            Call get_final_smpeq_part1_part2(eq)
            ReDim ret(2)
            ret(0) = eq(0) : ret(1) = eq(1) : ret(2) = "("
            k = 3
        End If
        If unkvar_rhs_division = True Then
            '(1/b)+(2/b) : common_part = b -> (1/b)*3
            ReDim Preserve ret(k + 2)
            ret(k) = "(" : ret(k + 1) = "1" : ret(k + 2) = "/"
            ret = add_arr1_to_arr(ret, common_part)
            ret = add_element_to_array(ret, UBound(ret) + 1, ")")
            unkvar_rhs_division = False
        Else
            ret = add_arr1_to_arr(ret, common_part)
        End If
        ret = add_element_to_array(ret, UBound(ret) + 1, CStr("*"))
        ret = add_arr1_to_arr(ret, simplified_equation_part1)
        ret = add_element_to_array(ret, UBound(ret) + 1, ")")
        If Len(Join(simplified_equation_part2)) > 0 Then
            'a=b+b+3 : part1 = (b*2) & part2 = +3 -> (b*2)+3
            ret = add_arr1_to_arr(ret, simplified_equation_part2)
        Else
            'a=b+b -> a=(b*2) -> a=b*2
            ret = remove_brackets_if_needed(ret)
        End If
        If Len(Join(Eqn_Part_With_Unknowns)) > 0 Then
            ret = replace_array_part_with_new_arr(eq, ret, func_start_point, func_end_point)
        End If
        ret = remove_unecessry_brackets(ret)
        get_final_simplified_eqn = ret
    End Function
    Function get_final_smpeq_part1_part2(eq() As String)
        Dim tmpeq() As String, v1() As String, v2() As String
        Dim i As Integer
        v1 = get_operands2(eq)
        tmpeq = get_compressed_equation(eq)
        'a=f+z1+3 -> f+z1+3
        If tmpeq(1) = "=" And tmpeq(1) <> "*" Then
            tmpeq = reduce_arr(tmpeq, 0)
            tmpeq = reduce_arr(tmpeq, 0)
        End If
        v2 = get_operands2(tmpeq)
        'f+z1+3 -> simplified_equation_part2 = (+,3)
        simplified_equation_part2 = get_remaining_part_eq(tmpeq, CStr(Chr(131)), v2)
        'f+z1+3 -> simplified_equation_part1 = (f,+,z1)
        tmpeq = modify_tmpeq(tmpeq, v1, Chr(131))
        'f+z1+3 -> (f,+,z1) -> simplified_equation_part1 = (f,+,(,f,*,3,))
        simplified_equation_part1 = disassemble_tempeq(eq, tmpeq)
        '(f,+,(,f,*,3,)) -> (1,+,(,1,*,3,)) -> (1+(1*3)) -> 4
        For i = 0 To UBound(simplified_equation_part1)
            If simplified_equation_part1(i) = Chr(131) Then
                simplified_equation_part1(i) = "1"
            End If
        Next i
        simplified_equation_part1 = add_brackets_start_end(simplified_equation_part1)
        simplified_equation_part1 = simplify_mul_and_div_by_one(simplified_equation_part1, Chr(131))
    End Function
    Function simplify_mul_and_div_by_one(eq() As String, var As String)
        Dim i As Integer, m As Integer, n As Integer, t1 As Integer, t2 As Integer
        Dim chk As Boolean
        Do
            chk = False
            i = 0
            Do
                If eq(i) = "1" And i + 1 <= UBound(eq) Then
                    If eq(i + 1) = "*" Then
                        chk = True
                        t1 = UBound(eq)
                        m = i + 2 : n = i
                        Do
                            eq(n) = eq(m)
                            n = n + 1
                            m = m + 1
                        Loop Until m > UBound(eq)
                        ReDim Preserve eq(UBound(eq) - 2)
                        t2 = UBound(eq)
                        i = t1 - t2
                    End If
                End If
                If eq(i) = "1" And i - 1 >= 0 Then
                    If eq(i - 1) = "*" Then
                        chk = True
                        t1 = UBound(eq)
                        m = i - 2 : n = i
                        Do
                            eq(n) = eq(m)
                            n = n - 1
                            m = m - 1
                        Loop Until m = -1
                        eq = reduce_arr(eq, 0)
                        eq = reduce_arr(eq, 0)
                        t2 = UBound(eq)
                        i = t1 - t2
                    End If
                End If
                i = i + 1
            Loop Until i > UBound(eq)
            If UBound(eq) > 0 Then
                i = 0
                Do
                    If eq(i) = "1" And i - 1 >= 0 Then
                        If eq(i - 1) = "/" Then
                            chk = True
                            If i + 1 <= UBound(eq) Then
                                m = i + 1 : n = i - 1
                            ElseIf i = UBound(eq) Then
                                m = UBound(eq) : n = i - 1
                            End If
                            Do
                                eq(n) = eq(m)
                                n = n + 1
                                m = m + 1
                            Loop Until m > UBound(eq)
                            ReDim Preserve eq(UBound(eq) - 2)
                            i = i - 2
                        End If
                    End If
                    i = i + 1
                Loop Until i > UBound(eq)
            End If
            eq = remove_brackets_single_variable(eq)
            eq = remove_unecessry_brackets(eq)
        Loop Until chk = False
        simplify_mul_and_div_by_one = eq
    End Function
    Function check_simplified_equation(eq() As String)
        Dim temp() As String, tmpeq() As String, arr(,) As String, chkarr() As String, t() As String
        Dim i As Integer, k As Integer, p As Integer, num1 As String, num2 As String
        Dim chk As Boolean, chk1 As Boolean, chk2 As Boolean, chk3 As Boolean, chk4 As Boolean
        Dim oper = New String() {"*", "/"}
        'get number of instance of "f" in eq
        num1 = Form1.number_of_instances(eq, CStr(Chr(131)))
        For i = 0 To UBound(eq)
            If eq(i) = ")" Then
                temp = get_eqn_part_before_closing_bracket(eq, i)
                If temp.Contains(CStr(Chr(131))) = True Then
                    'get number of instance of "f" between closed brackets
                    num2 = Form1.number_of_instances(temp, CStr(Chr(131)))
                    If num2 > 1 And num1 = num2 Then
                        'indiacate all "f" between sub-eq in eq
                        chk1 = True
                        temp = remove_brackets_start_end(temp)
                        '(f+(f*2))^2 -> tmpeq = f+(f*2)
                        tmpeq = temp
                        Exit For
                    End If
                End If
            End If
        Next i
        If chk1 = False Then
            tmpeq = eq
        End If
        k = 0
        For i = 0 To UBound(tmpeq)
            If tmpeq(i) = "(" Then
                temp = get_eqn_part_after_opening_bracket(tmpeq, i)
                i = i + UBound(temp) - 1
                If temp.Contains(CStr(Chr(131))) = True Then
                    arr = add_two_dim_array(arr, temp)
                    ReDim Preserve chkarr(k)
                    chkarr(k) = join_arr(temp)
                    k = k + 1
                End If
            ElseIf tmpeq(i) = CStr(Chr(131)) Then
                ReDim temp(0)
                temp(0) = tmpeq(i)
                arr = add_two_dim_array(arr, temp)
                ReDim Preserve chkarr(k)
                chkarr(k) = tmpeq(i)
                k = k + 1
            End If
        Next i
        For i = 0 To UBound(chkarr)
            '(f,(f*3)) -> chk=false
            If chkarr(0) <> chkarr(i) Then
                chk = False
            End If
        Next i
        If chk = False Then
            chk2 = True
            For i = 0 To UBound(arr, 1)
                temp = get_one_dim_array(arr, i)
                If UBound(temp) > 0 Then
                    chkarr = get_operators2(temp)
                    For j = 0 To UBound(chkarr)
                        If oper.Contains(chkarr(j)) = False And chkarr(j) <> "*" Then
                            'chk2 indicates the presence of operators such as "^","cos","tan" etc : (cos(f+2), f^2, log(f*3)) -> chk2=false
                            chk2 = False
                        End If
                        If chkarr(j) = "/" Then
                            'chk3 indicates presence of "/" : (2/f, f/2) -> chk3=true
                            chk3 = True
                        End If
                    Next j
                End If
            Next i
        End If
        If chk2 = True And chk3 = True Then
            'if division is present in equation all commom parts should be on same side of "/"
            chk = True
            For i = 0 To UBound(arr, 1)
                temp = get_one_dim_array(arr, i)
                If temp.Contains("/") = True Then
                    p = Array.IndexOf(temp, "/")
                    '"3/(f+3)" or "f/3" -> chk=true
                    For j = p + 1 To UBound(temp)
                        If temp(j) = CStr(Chr(131)) Then
                            'checks if unknown is on the RHS of "/" : 3/f -> chk4=True
                            chk4 = True : unkvar_rhs_division = True
                        End If
                    Next j
                End If
            Next i
            If chk4 = True Then
                For i = 0 To UBound(arr, 1)
                    temp = get_one_dim_array(arr, i)
                    If temp.Contains("/") = True Then
                        p = Array.IndexOf(temp, "/")
                        ReDim t(UBound(temp) - p + 1)
                        k = 1
                        'temp=2/(f*3) : p=2 -> t=(f*3)
                        For j = p + 1 To UBound(temp)
                            t(k) = temp(j)
                            k = k + 1
                        Next j
                        'f=(f+3) -> chk=false
                        If t.Contains(CStr(Chr(131))) = False Or t.Contains("+") = True Or t.Contains("-") = True Then
                            chk = False
                        End If
                    Else
                        chk = False
                    End If
                Next i
            End If
        ElseIf chk2 = True Then
            chk = True
        End If
        check_simplified_equation = chk
    End Function
    Function get_temp_simplified_eqn(eq() As String, common_part() As String)
        Dim i As Integer, k As Integer, n As Integer, t1 As Integer, t2 As Integer
        Dim ret() As String, temp() As String
        Dim repstr As String, str As String
        str = join_arr(common_part)
        Dim oper = New String() {"+", "-"}
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        i = 0
        Do
            'eq=(3*(b+3))+(b+3) -> (3*(b+3))
            If ret(i) = ")" Then
                temp = get_eqn_part_before_closing_bracket(ret, i)
                repstr = join_arr(temp)
                If repstr = str Then
                    'eq=(3*(b+3))+(b+3) : repstr=(b+3) : str=(b+3) -> ret=3*f1+f2
                    t1 = UBound(ret)
                    k = i - UBound(temp)
                    ret = replace_array_part_with_str(ret, CStr(Chr(131)), k, i)
                    t2 = UBound(ret)
                    i = i - (t1 - t2)
                End If
            ElseIf eq(i) = str Then
                'eq=(3*b)+b : str=b -> ret=(3*f1)+f2
                ret(i) = CStr(Chr(131))
            End If
            i = i + 1
        Loop Until i > UBound(ret)
        'a=SIN(b+((b+2)*3)) -> a=SIN(f+((f+2)*3)) -> Eqn_Part_With_Unknowns = f+((f+2)*3)
        Eqn_Part_With_Unknowns = chk_uknown_in_math_func(ret)
        If Len(Join(Eqn_Part_With_Unknowns)) > 0 Then
            'f+((f+2)*3) -> f+(f*3)+6
            Eqn_Part_With_Unknowns = modify_temp_simplified_eqn_wrt_multiplication_division(Eqn_Part_With_Unknowns)
            n = UBound(ret)
            'a=SIN(f+((f+2)*3)) -> a=SIN(f+(f*3)+6)
            ret = replace_array_part_with_new_arr(ret, Eqn_Part_With_Unknowns, func_start_point, func_end_point)
            func_end_point = func_end_point - (n - UBound(ret))
        Else
            ret = modify_temp_simplified_eqn_wrt_multiplication_division(ret)
        End If
        get_temp_simplified_eqn = ret
    End Function
    Function modify_temp_simplified_eqn_wrt_multiplication_division(eq() As String)
        Dim temp() As String, tmpeq() As String, o() As String, rep() As String, ret() As String, v() As String
        Dim i As Integer, m As Integer, n As Integer, k As Integer, t1 As Integer, t2 As Integer
        Dim str As String, str1 As String, str2 As String
        Dim oper = New String() {"+", "-", "*", "/"}
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        i = 0
        Do
            If eq(i) = ")" Then
                tmpeq = get_eqn_part_before_closing_bracket(eq, i)
                tmpeq = reduce_arr(tmpeq, 0)
                tmpeq = reduce_arr(tmpeq, UBound(tmpeq))
                temp = get_compressed_equation(tmpeq)
                o = get_operators2(temp)
                If Len(Join(o)) > 0 Then
                    If o(0) = "+" Or o(0) = "-" Then
                        If i + 2 <= UBound(eq) Then
                            '(f+2)*3 -> f+2 -> (f*3)+6
                            If IsNumeric(eq(i + 2)) = True And (eq(i + 1) = "/" Or eq(i + 1) = "*") Then
                                str1 = eq(i + 1)
                                str2 = eq(i + 2)
                                v = get_operands2(temp)
                                k = 0
                                For j = 1 To UBound(temp)
                                    If v.Contains(temp(j)) = True Then
                                        If IsNumeric(temp(j)) = True Then
                                            str = app.evaluate(temp(j) & str1 & str2)
                                            ReDim Preserve rep(k)
                                            rep(k) = str
                                            k = k + 1
                                        Else
                                            ReDim Preserve rep(k + 4)
                                            rep(k) = "(" : rep(k + 1) = temp(j) : rep(k + 2) = str1 : rep(k + 3) = str2 : rep(k + 4) = ")"
                                            k = k + 5
                                        End If
                                    Else
                                        ReDim Preserve rep(k)
                                        rep(k) = temp(j)
                                        k = k + 1
                                    End If
                                Next j
                                rep = disassemble_tempeq(tmpeq, rep)
                                t1 = UBound(eq)
                                m = i - tmpeq.Length - 2
                                n = i + 3
                                'f+((f+2)*3) : rep=(f*3)+6 -> f+((f*3)+6) -> f+(f*3)+6
                                eq = replace_array_part_with_new_arr(eq, rep, m, n)
                                eq = remove_unnecessary_brackets_add_subtract(eq)
                                t2 = UBound(eq)
                                i = i - (t1 - t2)
                            End If
                        End If
                    End If
                End If
            End If
            i = i + 1
        Loop Until i > UBound(eq)
        modify_temp_simplified_eqn_wrt_multiplication_division = eq
    End Function
    Function chk_uknown_in_math_func(eq() As String)
        Dim temp() As String, ret() As String
        Dim i As Integer, num1 As Integer, num2 As Integer
        num1 = Form1.number_of_instances(eq, CStr(Chr(131)))
        For i = 0 To UBound(eq)
            '(2/f1)+(4/f2) -> (2/f1) & (4/f2)
            If eq(i) = ")" Then
                temp = get_eqn_part_before_closing_bracket(eq, i)
                If temp.Contains(CStr(Chr(131))) = True Then
                    num2 = Form1.number_of_instances(temp, CStr(Chr(131)))
                    func_start_point = i - UBound(temp)
                    func_end_point = i
                    If num1 = num2 Then
                        'COS(f+f+3) -> (f+f+2) -> num1=num2=2
                        temp = reduce_arr(temp, 0)
                        temp = reduce_arr(temp, UBound(temp))
                        chk_uknown_in_math_func = temp
                        Exit Function
                    End If
                End If
            End If
        Next i
        chk_uknown_in_math_func = ret
    End Function
    Function get_part_of_eqn_with_multiple_instances(eq() As String, uknvar As String)
        Dim tmparr(,) As String, temp() As String, tmpeq() As String, arr() As String, retarr() As String, chkarr(,) As String, arrlen() As Integer, refarr() As String
        Dim i As Integer, j As Integer, p As Integer, min As Integer, tmplen As Integer
        Dim chk As Boolean, chk1 As Boolean, chk2 As Boolean, chk_arr As Boolean, chkb As Boolean
        ReDim tmpeq(UBound(eq))
        eq.CopyTo(tmpeq, 0)
        Do
            chkb = True
            'two dimensional array of unkvar and functions containing unkvar
            'a=(b*3)+b+3 -> array=(((b*3)),(b))
            For i = 0 To UBound(tmpeq)
                If tmpeq(i) = "(" Then
                    temp = get_eqn_part_after_opening_bracket(tmpeq, i)
                    i = i + UBound(temp) - 1
                    If temp.Contains(uknvar) = True Then
                        chkarr = add_two_dim_array(chkarr, temp)
                    End If
                ElseIf tmpeq(i) = uknvar Then
                    ReDim temp(0)
                    temp(0) = uknvar
                    chkarr = add_two_dim_array(chkarr, temp)
                End If
            Next i
            chk_arr = IsArrayEmpty(chkarr)
            If chk_arr = False Then
                If UBound(chkarr, 1) = 0 Then
                    'all unkvar in single function e.g.: log((b+b+2)^2) -> ((b+b+2)^2) -> b+b+2
                    Erase chkarr
                    For i = 0 To UBound(tmpeq)
                        If tmpeq(i) = "(" Then
                            temp = get_eqn_part_after_opening_bracket(tmpeq, i)
                            i = i + UBound(temp) - 1
                            If temp.Contains(uknvar) = True Then
                                tmpeq = reduce_arr(temp, 0)
                                ReDim Preserve tmpeq(UBound(tmpeq) - 1)
                                Exit For
                            End If
                        End If
                    Next i
                Else
                    chkb = False
                End If
            End If
        Loop Until chk_arr = True Or chkb = False
        If chk_arr = False Then
            Do
                chk = True : chk1 = True
                'chk = get all parts of two dimnsional array and check if same
                For i = 0 To UBound(chkarr, 1)
                    retarr = get_one_dim_array(chkarr, i)
                    ReDim Preserve arr(i)
                    arr(i) = join_arr(retarr)
                Next i
                For i = 0 To UBound(arr)
                    If arr(0) <> arr(i) Then
                        chk = False
                    End If
                Next i
                If UBound(arr) = 0 Then
                    chk = False
                End If
                'chk1 = if all parts not same get array length of the each element and check if they are same or different
                If chk = False Then
                    ReDim arrlen(UBound(arr))
                    For i = 0 To UBound(chkarr, 1)
                        temp = get_one_dim_array(chkarr, i)
                        arrlen(i) = UBound(temp)
                    Next i
                    For i = 0 To UBound(arr)
                        If arrlen(0) <> arrlen(i) Then
                            chk1 = False
                        End If
                    Next i
                    'if part lenghts are the same
                    If chk1 = True Then
                        'reduce all same length variables
                        ' ((b*3),(b*2)) -> (b*3) -> b & (b*2) -> b
                        For i = 0 To UBound(chkarr, 1)
                            temp = get_one_dim_array(chkarr, i)
                            '(b+3)*3 -> (b+3)
                            temp = reduce_array(temp, uknvar)
                            tmparr = add_two_dim_array(tmparr, temp)
                        Next i
                        chkarr = tmparr
                        Erase tmparr
                        'if reduced to one variable exit function
                        For i = 0 To UBound(chkarr, 1)
                            temp = get_one_dim_array(chkarr, i)
                            If UBound(temp) = 1 Then
                                chk2 = True
                                chk = False
                            End If
                        Next i
                    Else
                        'get subarr with the least number of variables and operators
                        '3*(b+5),(b+2) -> (b+2)
                        min = arrlen.Min
                        p = Array.IndexOf(arrlen, min)
                        refarr = get_one_dim_array(chkarr, p)
                        For i = 0 To UBound(arrlen)
                            If arrlen(i) <> min Then
                                temp = get_one_dim_array(chkarr, i)
                                'reduce lenght of smaller array
                                '(b+2) : 2*(b+3) -> ret=b
                                temp = reduce_array_unequal_lenght(temp, refarr, uknvar)
                                tmplen = UBound(temp)
                                If tmplen <= min Then
                                    For j = 0 To UBound(chkarr, 2)
                                        chkarr(i, j) = ""
                                    Next j
                                    For j = 0 To UBound(temp)
                                        chkarr(i, j) = temp(j)
                                    Next j
                                    min = tmplen
                                    refarr = get_one_dim_array(chkarr, i)
                                End If
                            End If
                        Next i
                    End If
                End If
            Loop Until chk = True Or chk2 = True
        Else
            chk2 = True
        End If
        If chk = True Then
            get_part_of_eqn_with_multiple_instances = retarr
        ElseIf chk2 = True Then
            ReDim temp(0)
            temp(0) = uknvar
            get_part_of_eqn_with_multiple_instances = temp
        End If
    End Function
    Function reduce_array_unequal_lenght(arr() As String, refarr() As String, uknvar As String)
        Dim ret() As String, temp() As String, str1 As String, str2 As String, chk As Boolean
        Dim n As Integer
        str1 = join_arr(refarr)
        str2 = join_arr(arr)
        n = InStr(str2, str1)
        If n <> 0 Then
            'str1=(b+3) : str2=2*(b+3) -> ret=(b+3)
            ReDim ret(UBound(refarr)) : refarr.CopyTo(ret, 0)
        Else
            ReDim temp(UBound(refarr)) : refarr.CopyTo(temp, 0)
            Do
                'str1=(b+5) : str2=2*(b+3) -> temp="b"
                temp = reduce_array(temp, uknvar)
                str1 = join_arr(temp)
                'b in 2*(b+3) -> n=3 -> n <> 0
                n = InStr(str2, str1)
                If UBound(temp) = 1 Then
                    chk = True
                End If
            Loop Until n <> 0 Or chk = True
            ret = temp
        End If
        reduce_array_unequal_lenght = ret
    End Function
    Function reduce_array(arr() As String, uknvar As String)
        Dim ret() As String
        Dim chk As Boolean
        Dim oper = New String() {"+", "-", "*", "/", "^"}
        ReDim ret(UBound(arr))
        arr.CopyTo(ret, 0)
        '(b*3) -> b*3
        chk = check_brackets_start_end(ret)
        If chk = True Then
            ret = reduce_arr(ret, 0)
            ret = reduce_arr(ret, UBound(ret))
        End If
        If ret(0) = "(" Then
            '(b+3)*3 -> (b+3)*
            ReDim Preserve ret(UBound(ret) - 1)
        ElseIf UBound(ret) > 0 Then
            '2/b -> /b
            If ret(0) <> uknvar Then
                ret = reduce_arr(ret, 0)
            Else
                'b^2 -> b^
                ReDim Preserve ret(UBound(ret) - 1)
            End If
        End If
        If oper.Contains(ret(0)) = True Then
            '/b -> b
            ret = reduce_arr(ret, 0)
        ElseIf oper.Contains(ret(UBound(ret))) = True Then
            'b^ -> b
            ReDim Preserve ret(UBound(ret) - 1)
        End If
        reduce_array = ret
    End Function
    Function reduce_eqn(eq() As String, inputv() As String, v() As String, unkvar As String)
        Dim arr() As String, strarr As String, ret() As String, tmpret() As String, o() As String
        Dim temp() As String, str1 As String, str2 As String
        Dim i As Integer, k As Integer, l As Integer, n As Integer, p As Integer, t1 As Integer, t2 As Integer
        Dim mfunc = New String() {"LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        Dim trigfunc = New String() {"SIN", "COS", "TAN"}
        Dim atrigfunc = New String() {"ASIN", "ACOS", "ATAN"}
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        'replace the known variables with the input value
        'a=b+b+c : c=4 -> a=b+b+4
        'If arr.Contains(t(k - 1)) = True Then
        For i = 1 To UBound(ret)
            If v.Contains(ret(i)) = True And ret(i) <> unkvar And ret(i) <> "*" Then
                p = Array.IndexOf(v, ret(i))
                ret(i) = inputv(p)
            End If
        Next i
        'if unkwn is on LHS then shift to RHS
        'b=3*b-2 -> 0=3*b-2-b
        If ret(0) = unkvar Then
            temp = get_compressed_equation(eq)
            o = get_operators2(temp)
            If o(0) <> "+" Or o(0) <> "-" Then
                ret = add_element_to_array(ret, 2, "(")
                ret = add_element_to_array(ret, UBound(ret) + 1, ")")
            End If
            ReDim Preserve ret(UBound(ret) + 2)
            ret(UBound(ret) - 1) = "-"
            ret(UBound(ret)) = eq(0)
            ret(0) = "0"
        End If
        i = 0
        Do
            'a=b+(3*2) -> (3*2)
            If ret(i) = ")" Then
                arr = get_eqn_part_before_closing_bracket(ret, i)
                k = i - UBound(arr) - 1
                l = i + 1
                If arr.Contains(unkvar) = False Then
                    strarr = join_arr(arr)
                    If trigfunc.Contains(arr(1)) = True Then
                        'SIN(90) -> SIN(90*PI()/180) -> 1
                        strarr = Left(strarr, Len(strarr) - 2) & "*PI()/180))"
                    ElseIf atrigfunc.Contains(arr(1)) = True Then
                        'ASIN(1) -> ASIN(1)*(180/PI()) -> 90
                        strarr = strarr & "*(180/PI())"
                    End If
                    '(3*2) -> 6
                    strarr = app.evaluate(strarr)
                Else
                    '(3+b+2) -> (b+5)
                    temp = solve_eqn_part_contains_unknown(arr, unkvar)
                End If
                'check if ans is negative i.e -5 or -12 or ans is part of math func such as log or sin
                'if yes enclose in brackets
                If Len(strarr) > 0 Then
                    If Sign(CDbl(strarr)) = -1 Then
                        ReDim temp(1)
                        temp(0) = "-"
                        temp(1) = Mid$(CStr(strarr), 2)
                    Else
                        ReDim temp(0)
                        temp(0) = strarr
                    End If
                    If Sign(CDbl(strarr)) = -1 Or mfunc.Contains(ret(k)) = True And ret(k) <> "*" Then
                        temp = add_brackets_start_end(temp)
                    End If
                    strarr = ""
                End If
                str1 = join_arr(arr)
                str2 = join_arr(temp)
                If str1 <> str2 Then
                    t1 = UBound(ret)
                    'replace the part of equation with the new solved part
                    'a=b+b+(3*2) : (3*2)=6 -> a=b+b+6
                    ret = replace_array_part_with_new_arr(ret, temp, k, l)
                    t2 = UBound(ret)
                    i = i - (t1 - t2)
                End If
            End If
            i = i + 1
        Loop Until i > UBound(ret)
        'get part of equation after "=" sign
        'a=6+b+b-1 -> tmpret = 6+b+b-1
        n = 0
        ReDim tmpret((UBound(ret) - 2))
        For i = 2 To UBound(ret)
            tmpret(n) = ret(i)
            n = n + 1
        Next i
        ReDim Preserve ret(1)
        'solve remaining equation parts
        '6+b+b-1 -> b+b+5
        tmpret = solve_eqn_part_contains_unknown(tmpret, unkvar)
        'join remaining part of equation
        'a=6+b+b-1 : (6+b+b-1)=b+b+5 -> a=b+b+5
        ret = add_arr1_to_arr(ret, tmpret)
        If eq(0) = unkvar Then
            ret(0) = Chr(220)
        End If
        reduce_eqn = ret
    End Function
    Function solve_eqn_part_contains_unknown(eq() As String, unkvar As String)
        Dim tmpeq() As String, o() As String, v() As String, reteq() As String, remeq() As String
        Dim solve As String
        Dim i As Integer, chkb As Boolean, chk As Boolean
        Dim operand = New String() {"+", "-", "*", "/", "^"}
        'remove brackets at beginning & end of input
        '(b+2) -> b+2
        chkb = check_brackets_start_end(eq)
        If chkb = True Then
            eq = remove_brackets_start_end(eq)
        End If
        'b+(b+1) -> b+b+1
        eq = remove_unnecessary_brackets_add_subtract(eq)
        tmpeq = get_compressed_equation(eq)
        v = get_operands2(tmpeq)
        'check if eq contains known variables to be shifted
        '3*b*z1 -> chk=True : b+b+2 -> chk=False as "2" is at end of equation
        For i = 0 To UBound(v)
            If Mid$(v(i), 1, 1) <> Chr(158) And v(i) <> unkvar And i <> UBound(v) Then
                chk = True
            End If
        Next i
        If chk = True Then
            'z1-2+b+5 -> -2+5
            remeq = get_remaining_part_eq(tmpeq, unkvar, v)
            If Len(Join(remeq)) > 0 Then
                v = get_operands2(remeq)
                If remeq(0) = "*" Then
                    remeq = reduce_arr(remeq, 0)
                End If
                solve = join_arr(remeq)
                If UBound(remeq) > 0 Then
                    solve = app.evaluate(solve)
                End If
                o = get_operators(tmpeq)
                'z1-2+b+5 -> z1+b
                tmpeq = modify_tmpeq(tmpeq, v, unkvar)
                'add solve to end of tmpeq
                'z1+b : solve=3 -> z1+b+3
                If o(0) = "+" Or o(0) = "-" Then
                    ReDim Preserve tmpeq(UBound(tmpeq) + 2)
                    If Sign(CDbl(solve)) = -1 Then
                        'z1+b : solve="-3" -> z1+b-3
                        tmpeq(UBound(tmpeq) - 1) = "-"
                        tmpeq(UBound(tmpeq)) = Mid$(CStr(solve), 2)
                    Else
                        tmpeq(UBound(tmpeq) - 1) = "+"
                        tmpeq(UBound(tmpeq)) = CStr(solve)
                    End If
                Else
                    If Sign(CDbl(solve)) = -1 Then
                        'solve = -4 -> z1/(-4)
                        ReDim Preserve tmpeq(UBound(tmpeq) + 5)
                        tmpeq(UBound(tmpeq) - 4) = o(0)
                        tmpeq(UBound(tmpeq) - 3) = "("
                        tmpeq(UBound(tmpeq) - 2) = "-"
                        tmpeq(UBound(tmpeq) - 1) = Mid$(CStr(solve), 2)
                        tmpeq(UBound(tmpeq)) = ")"
                    Else
                        'solve = 4 -> z1/4
                        ReDim Preserve tmpeq(UBound(tmpeq) + 2)
                        tmpeq(UBound(tmpeq) - 1) = o(0)
                        tmpeq(UBound(tmpeq)) = CStr(solve)
                    End If
                End If
                reteq = disassemble_tempeq(eq, tmpeq)
                'replace all "z" with corresponding equation part
                'z1+b+3 : z1=(b*2) -> (b*2)+b+3
            End If
        End If
        If Len(Join(reteq)) = 0 Then
            reteq = eq
        End If
        'b*b -> b^2
        reteq = simplify_mul_to_power(reteq, unkvar)
        'b*(b+1) -> (b^2)+b
        reteq = simplify_wrt_division_multiplication(reteq, unkvar)
        If chkb = True Then
            reteq = add_brackets_start_end(reteq)
        End If
        solve_eqn_part_contains_unknown = reteq
    End Function
    Function simplify_mul_to_power(eq() As String, unkvar As String)
        Dim ret() As String, tmpret() As String, temp() As String, arr(,) As String, eqpart() As String, str As String, o() As String
        Dim i As Integer, k As Integer, l As Integer, chk As Boolean, p As Integer
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        Do
            k = 0 : chk = False
            'get all parts of eq containing unkvar
            '(COS(b))*b*(COS(b))*2 -> arr=((COS(b)),b,(COS(b)))
            For i = 0 To UBound(ret)
                If ret(i) = "(" Then
                    temp = get_eqn_part_after_opening_bracket(ret, i)
                    i = i + UBound(temp) - 1
                    If temp.Contains(unkvar) = True Then
                        ReDim Preserve eqpart(k)
                        eqpart(k) = join_arr(temp)
                        k = k + 1
                    End If
                    arr = add_two_dim_array(arr, temp)
                ElseIf ret(i) = unkvar Then
                    ReDim Preserve eqpart(k)
                    eqpart(k) = ret(i)
                    k = k + 1
                End If
            Next i
            'gets duplicate values of array if any ; arr=((COS(b)),b,(COS(b))) -> str=(COS(b))
            str = get_common_part(eqpart)
            If Len(str) > 0 Then
                temp = get_compressed_equation(ret)
                o = get_operators(temp)
                l = 0 : k = 0
                If o(0) = "*" Then
                    chk = True
                    For i = 0 To UBound(ret)
                        If ret(i) = "(" Then
                            temp = get_eqn_part_after_opening_bracket(ret, i)
                            i = i + UBound(temp)
                            If join_arr(temp) = str And l = 0 Then
                                '(COS(b))*b*(COS(b)) -> ((COS(b))^e)
                                ReDim Preserve tmpret(k)
                                tmpret(k) = "("
                                tmpret = add_arr1_to_arr(tmpret, temp)
                                k = UBound(tmpret) + 1
                                ReDim Preserve tmpret(UBound(tmpret) + 3)
                                tmpret(k) = "^" : tmpret(k + 1) = Chr(233) : tmpret(k + 2) = ")"
                                k = k + 3
                                l = l + 1
                            ElseIf join_arr(temp) = str And l <> 0 Then
                                '(COS(b))*b*(COS(b)) : temp=(COS(b)) ->  ((COS(b))^e)*b* -> ((COS(b))^e)*b & l=2
                                ReDim Preserve tmpret(UBound(tmpret) - 1)
                                k = k - 1
                                l = l + 1
                            Else
                                tmpret = add_arr1_to_arr(tmpret, temp)
                                k = UBound(tmpret) + 1
                            End If
                        ElseIf ret(i) = str And l = 0 Then
                            'b*b*b*2 -> (b^e)*2 : l=3-> (b^3)*2
                            ReDim Preserve tmpret(k + 4)
                            tmpret(k) = "(" : tmpret(k + 1) = ret(i) : tmpret(k + 2) = "^" : tmpret(k + 3) = CStr(Chr(233)) : tmpret(k + 4) = ")"
                            k = k + 5
                            l = l + 1
                        ElseIf ret(i) = str And l <> 0 Then
                            ReDim Preserve tmpret(UBound(tmpret) - 1)
                            k = k - 1
                            l = l + 1
                        Else
                            ReDim Preserve tmpret(k)
                            tmpret(k) = ret(i)
                            k = k + 1
                        End If
                    Next i
                    '(COS(b))*b*(COS(b)) -> ((COS(b))^e)*b : l=2 -> ((COS(b))^2)*b
                    p = Array.IndexOf(tmpret, CStr(Chr(233)))
                    tmpret(p) = l
                    ReDim ret(UBound(tmpret)) : tmpret.CopyTo(ret, 0)
                End If
            End If
        Loop Until chk = False
        simplify_mul_to_power = ret
    End Function
    Function get_common_part(eqpart() As String)
        Dim i As Integer, j As Integer
        Dim ret As String
        For i = 0 To UBound(eqpart)
            For j = 0 To UBound(eqpart)
                If i <> j Then
                    If eqpart(i) = eqpart(j) Then
                        ret = eqpart(i)
                    End If
                End If
            Next j
        Next i
        get_common_part = ret
    End Function
    Function modify_tmpeq(eq() As String, v() As String, unkvar As String)
        Dim i As Integer
        Dim tmpeq() As String, var() As String
        Dim oper = New String() {"+", "-", "*"}
        tmpeq = eq
        'removes all variables except unknown and compressed subeqns
        'b-2+z1+4 -> b-+z1+
        i = 0
        Do
            If v.Contains(tmpeq(i)) = True And tmpeq(i) <> unkvar And tmpeq(i) <> "*" Then
                tmpeq = reduce_arr(tmpeq, i)
                i = i - 1
            End If
            i = i + 1
        Loop Until i > UBound(tmpeq)
        var = get_operands2(tmpeq)
        'removes all unnecessary operators(+,-,*,etc)
        'b-+z1+ -> b+z1
        i = 0
        Do
            If oper.Contains(tmpeq(i)) = True Then
                If i + 1 <= UBound(tmpeq) Then
                    If var.Contains(tmpeq(i + 1)) = False Or tmpeq(i + 1) = "*" Then
                        tmpeq = reduce_arr(tmpeq, i)
                        i = i - 1
                    End If
                ElseIf i = UBound(tmpeq) Then
                    tmpeq = reduce_arr(tmpeq, i)
                End If
            End If
            i = i + 1
        Loop Until i > UBound(tmpeq)
        '*b -> b
        If var.Contains(tmpeq(0)) = False And tmpeq(0) <> "-" Or tmpeq(0) = "*" Then
            tmpeq = reduce_arr(tmpeq, 0)
        End If
        modify_tmpeq = tmpeq
    End Function
    Function get_remaining_part_eq(eq() As String, unkvar As String, v() As String)
        Dim ret() As String
        Dim i As Integer
        Dim c As New Collection
        Dim oper = New String() {"+", "-", "*"}
        '3*b*2 -> *3*2 : b+b+3 -> +3
        For i = 0 To UBound(eq)
            'check for varaible other than unknown and tmpeq variable(z1) i.e. 3*b*z1*2 -> 3 & 2
            If v.Contains(eq(i)) = True And eq(i) <> unkvar And eq(i) <> "*" And Mid$(eq(i), 1, 1) <> Chr(158) Then
                If i = 0 Then
                    '3*b -> *3
                    If oper.Contains(eq(i + 1)) = True Then
                        '3-b -> +3
                        If eq(i + 1) = "-" Then
                            c.Add("+")
                        Else
                            c.Add(eq(i + 1))
                        End If
                        c.Add(eq(i))
                    End If
                Else
                    'b*2 -> *2
                    If oper.Contains(eq(i - 1)) = True Then
                        c.Add(eq(i - 1))
                        c.Add(eq(i))
                    End If
                End If
            End If
        Next i
        ret = convert_collection_to_array(c)
        get_remaining_part_eq = ret
    End Function
    Function remove_unnecessary_brackets_add_subtract(eq() As String)
        Dim i As Integer, t1 As Integer, t2 As Integer, k As Integer
        Dim temp() As String, rep() As String, ret() As String
        i = 0
        Do
            If eq(i) = ")" Then
                temp = get_eqn_part_before_closing_bracket(eq, i)
                k = i - UBound(temp)
                temp = remove_brackets_start_end(temp)
                rep = remove_brackets_add_sub(temp)
                If join_arr(temp) <> join_arr(rep) Then
                    t1 = UBound(eq)
                    eq = replace_array_part_with_new_arr(eq, rep, k, i + 1)
                    t2 = UBound(eq)
                    i = i - (t1 - t2)
                End If
            End If
            i = i + 1
        Loop Until i > UBound(eq)
        ret = remove_brackets_add_sub(eq)
        remove_unnecessary_brackets_add_subtract = ret
    End Function
    Function remove_brackets_add_sub(eq() As String)
        Dim neq() As String, ret() As String, temp() As String, temp1() As String, tmpeq() As String, arr(,) As String, o() As String, o1() As String, o2() As String
        Dim i As Integer, j As Integer, k As Integer, n As Integer, m As Integer, num As Integer, t1 As Integer, t2 As Integer
        Dim str As String
        Dim chk As Boolean
        ReDim neq(UBound(eq))
        eq.CopyTo(neq, 0)
        i = 0 : num = 0
        Do
            If neq(i) = ")" Then
                'b+(b+1) -> temp=(b+1)
                temp = get_eqn_part_before_closing_bracket(neq, i)
                k = i - UBound(temp)
                o = get_operators2(temp)
                'b-((b*3)+(b*2)-2) -> b-(A1+A2-2)
                If Len(Join(o)) > 0 Then
                    chk = check_operators(o)
                    If chk = True And (o(0) = "+" Or o(0) = "-") Then
                    Else
                        temp1 = temp : ReDim temp(0)
                        m = 0
                        For j = 0 To UBound(temp1)
                            'b-((COS(b*3))-2)) -> b-((COSA1)-2)) : A1=(b*3) -> b-((COS(b*3))-2)) -> b-(A2-2)) : A2=(COS(b*3))
                            If Mid$(temp1(j), 1, 1) = Chr(197) And temp1(j) <> "*" Then
                                n = Mid$(temp1(j), 2)
                                tmpeq = get_one_dim_array(arr, n)
                                temp = add_arr1_to_arr(temp, tmpeq)
                                m = UBound(temp) + 1
                            Else
                                ReDim Preserve temp(m)
                                temp(m) = temp1(j)
                                m = m + 1
                            End If
                        Next j
                        arr = add_two_dim_array(arr, temp)
                        t1 = UBound(neq)
                        str = Chr(197) & num : num = num + 1
                        neq = replace_array_part_with_str(neq, str, k, i)
                        t2 = UBound(neq)
                        i = t1 - t2
                    End If
                End If
            End If
            i = i + 1
        Loop Until i > UBound(neq)
        str = ""
        o1 = get_operators(neq)
        o2 = get_operators2(neq)
        If Len(Join(o2)) > 0 Then
            chk = check_operators(o2)
            If chk = True And (o2(0) = "+" Or o2(0) = "-") And o1.Contains("(") = True And UBound(o2) > 0 Then
                i = 0
                Do
                    If neq(i) = "(" Then
                        k = i
                        t1 = UBound(neq)
                        If i - 1 > 0 Then
                            If neq(i - 1) = "-" Then
                                If neq(i + 1) = "-" Or neq(i + 1) = "+" Then
                                    'b-(-b+1) -> b+(+b+1) -> b+(b+1) -> b+b+1 : 'b-(+b+1) -> b-(b+1) -> b-b-1
                                    If neq(i + 1) = "-" Then
                                        neq(i - 1) = "+"
                                    End If
                                    neq = reduce_arr(neq, i + 1)
                                End If
                                'b-(b+b-2) -> b-b-b+2
                                Do
                                    If neq(i) = "-" Then
                                        neq(i) = "+"
                                    ElseIf neq(i) = "+" Then
                                        neq(i) = "-"
                                    End If
                                    i = i + 1
                                Loop Until neq(i) = ")"
                            ElseIf neq(i - 1) = "+" Then
                                If neq(i + 1) = "-" Or neq(i + 1) = "+" Then
                                    ''b+(-b-2) -> b-(-b+2) -> b-(b+2) -> b-b-2
                                    If neq(i + 1) = "-" Then
                                        neq(i - 1) = "-"
                                    End If
                                    neq = reduce_arr(neq, i + 1)
                                End If
                                Do
                                    i = i + 1
                                Loop Until neq(i) = ")"
                            End If
                        Else
                            Do
                                i = i + 1
                            Loop Until neq(i) = ")"
                        End If
                        neq = reduce_arr(neq, k)
                        neq = reduce_arr(neq, i - 1)
                        t2 = UBound(neq)
                        i = t1 - t2
                    End If
                    i = i + 1
                Loop Until i > UBound(neq)
            End If
        End If
        'b-A1+3 : A1=(b*3) -> b-(b*3)+3
        k = 0
        For i = 0 To UBound(neq)
            If Mid$(neq(i), 1, 1) = Chr(197) Then
                n = Mid$(neq(i), 2)
                temp = get_one_dim_array(arr, n)
                For j = 0 To UBound(temp)
                    ReDim Preserve ret(k)
                    ret(k) = temp(j)
                    k = k + 1
                Next j
            Else
                ReDim Preserve ret(k)
                ret(k) = neq(i)
                k = k + 1
            End If
        Next i
        ret = remove_unecessry_brackets(ret)
        remove_brackets_add_sub = ret
    End Function
    Function remove_unecessry_brackets(eq() As String)
        Dim ret() As String, temp() As String, tmpeq() As String, v() As String, o() As String
        Dim i As Integer, n As Integer, m As Integer, l As Integer, k As Integer
        ReDim ret(UBound(eq)) : eq.CopyTo(ret, 0) : i = 0
        Do
            If ret(i) = "(" Then
                k = 0 : n = 1 : m = i : l = i
                Do
                    ReDim Preserve tmpeq(k)
                    tmpeq(k) = ret(m)
                    m = m + 1
                    k = k + 1
                    If ret(m) = "(" Then
                        n = n + 1
                    ElseIf ret(m) = ")" Then
                        n = n - 1
                    End If
                Loop Until n = 0
                ReDim Preserve tmpeq(k)
                tmpeq(k) = ")"
                temp = reduce_arr(tmpeq, 0)
                ReDim Preserve temp(UBound(temp) - 1)
                temp = get_compressed_equation(temp)
                v = get_operands2(temp)
                o = get_operators2(temp)
                If UBound(v) = 0 And v(0) = Chr(158) & 0 And Len(Join(o)) = 0 Then
                    ret = reduce_arr(ret, l)
                    ret = reduce_arr(ret, m - 1)
                    i = 0
                End If
            End If
            i = i + 1
        Loop Until i > UBound(ret)
        remove_unecessry_brackets = ret
    End Function
    Function replace_array_part_with_str(arr() As String, str As String, p1 As Integer, p2 As Integer)
        Dim i As Integer, k As Integer
        Dim ret() As String
        ReDim ret(UBound(arr)) : arr.CopyTo(ret, 0)
        'b-(b*2)+2 : p1=3 : p2=7 : str=A1 -> b-A1+2
        ReDim Preserve ret(p1)
        ret(p1) = str
        k = UBound(ret) + 1
        For i = p2 + 1 To UBound(arr)
            ReDim Preserve ret(k)
            ret(k) = arr(i)
            k = k + 1
        Next i
        replace_array_part_with_str = ret
    End Function
    Function remove_brackets_start_end(arr() As String)
        Dim i As Integer
        For i = 0 To UBound(arr) - 1
            arr(i) = arr(i + 1)
        Next i
        ReDim Preserve arr(UBound(arr) - 2)
        remove_brackets_start_end = arr
    End Function
End Module
Module module4
    Private Common_Operator As String
    Private Structure ModuleType
        Dim Divisor_Multiplier() As String
        Dim Divident_Multiplicand() As String
        Dim Type_of_Operation As String
        Dim Div_Mul_Part() As String
        Dim Modified_Div_Mul() As String
        Dim Check_Division As Boolean
        Dim Check_Multiplication As Boolean
        Dim Base As String
        Dim Power() As String
        Dim String_Power As String
        Dim Base_Array() As String
        Dim Power_Array(,) As String
        Dim Div_Mul_Part_Position As Boolean
    End Structure
    Private Operation As ModuleType
    Private emptytype As ModuleType
    Function simplify_wrt_division_multiplication(eq() As String, unkvar As String)
        Dim ret() As String, neq() As String
        Dim i As Integer, k As Integer, l As Integer, n1 As Integer
        Dim chkb As Boolean, chkp As Boolean
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0) : n1 = 1
        '(b^2)^2 -> b^4
        Common_Operator = ""
        ret = simplify_wrt_powers(ret, unkvar)
        Do
            For i = 0 To UBound(ret)
                If ret(i) = ")" Then
                    Operation.Divisor_Multiplier = get_eqn_part_before_closing_bracket(ret, i)
                    l = i
                    k = i - UBound(Operation.Divisor_Multiplier)
                    'check if part can act as multiplier or divisor
                    'the equation part should only contain operators *,/ ,^ or should have "^" as main operator
                    chkp = check_multiplier_divisor(unkvar)
                ElseIf ret(i) = unkvar Then
                    chkp = True
                    ReDim Operation.Divisor_Multiplier(0)
                    Operation.Divisor_Multiplier(0) = unkvar
                    k = i
                    l = i
                End If
                If chkp = True Then
                    neq = check_for_and_simplify_wrt_division(ret, unkvar, k)
                    If Len(Join(neq)) > 0 Then
                        ReDim ret(UBound(neq)) : neq.CopyTo(ret, 0)
                        Exit For
                    Else
                        neq = check_for_and_simplify_wrt_multiplication(ret, unkvar, k, l)
                        If Len(Join(neq)) > 0 Then
                            ReDim ret(UBound(neq)) : neq.CopyTo(ret, 0)
                            Exit For
                        End If
                    End If
                    chkp = False
                End If
            Next i
            Operation = emptytype
        Loop Until Len(Join(neq)) = 0
        chkb = check_brackets_start_end(ret)
        If chkb = True Then
            ret = remove_brackets_start_end(ret)
        End If
        ret = remove_unnecessary_brackets_add_subtract(ret)
        'If Join(ret, "") <> Join(eq, "") Then
        '    MsgBox(Join(eq, "") & Environment.NewLine & Join(ret, ""))
        'End If
        simplify_wrt_division_multiplication = ret
    End Function
    Function check_for_and_simplify_wrt_multiplication(eq() As String, unkvar As String, k As Integer, l As Integer)
        Dim ret() As String, multiplicand() As String
        Dim i As Integer, n As Integer, multiplicand_start_point As Integer, multiplicand_end_point As Integer
        Dim chk1 As Boolean, chk2 As Boolean
        If l + 1 <= UBound(eq) Then
            If eq(l + 1) = "*" Then
                chk1 = True
            End If
        End If
        If k - 1 > 0 Then
            If eq(k - 1) = "*" Then
                chk2 = True
            End If
        End If
        If chk1 = True Then
            If UBound(Operation.Divisor_Multiplier) = 0 And Operation.Divisor_Multiplier(0) = unkvar And eq(l + 2) = unkvar Then
                ReDim ret(UBound(eq)) : eq.CopyTo(ret, 0)
                ret(l + 1) = "^"
                ret(l + 2) = "2"
                check_for_and_simplify_wrt_multiplication = ret
                Exit Function
            End If
            For i = l + 2 To UBound(eq)
                If eq(i) = "(" Or eq(i) = unkvar Then
                    If eq(i) = "(" Then
                        Operation.Divident_Multiplicand = get_eqn_part_after_opening_bracket_without_start_end_brackets(eq, i)
                        n = i
                        i = i + Operation.Divident_Multiplicand.Length + 1
                    Else
                        ReDim Operation.Divident_Multiplicand(0)
                        Operation.Divident_Multiplicand(0) = unkvar
                        n = i
                    End If
                    If eq(n - 1) = "*" Then
                        Operation.Type_of_Operation = "Multiplication"
                        multiplicand = simplify_division_multiplication(unkvar)
                        If Operation.Check_Multiplication = True Then
                            ReDim ret(UBound(eq)) : eq.CopyTo(ret, 0)
                            ret = remove_array_part(ret, k - 1, l + 2)
                            If UBound(Operation.Divident_Multiplicand) = 0 Then
                                multiplicand_start_point = n - Operation.Divisor_Multiplier.Length - 3
                                multiplicand_end_point = n + Operation.Divident_Multiplicand.Length - Operation.Divisor_Multiplier.Length
                            Else
                                multiplicand_start_point = n - 1 - Operation.Divisor_Multiplier.Length - 1
                                multiplicand_end_point = n + 2 + Operation.Divident_Multiplicand.Length - Operation.Divisor_Multiplier.Length - 1
                            End If
                            multiplicand = add_brackets_start_end(multiplicand)
                            ret = replace_array_part_with_new_arr(ret, multiplicand, multiplicand_start_point, multiplicand_end_point)
                            ret = remove_unecessry_brackets(ret)
                            check_for_and_simplify_wrt_multiplication = ret
                            Exit Function
                        End If
                    Else
                        Exit For
                    End If
                ElseIf eq(i) = ")" Then
                    Exit For
                End If
            Next i
        End If
        If chk2 = True Then
            For i = k - 2 To 0 Step -1
                If eq(i) = ")" Or eq(i) = unkvar Then
                    If eq(i) = ")" Then
                        Operation.Divident_Multiplicand = get_eqn_part_before_closing_bracket(eq, i)
                        Operation.Divident_Multiplicand = remove_brackets_start_end(Operation.Divident_Multiplicand)
                        n = i
                        i = i - Operation.Divident_Multiplicand.Length - 1
                    Else
                        ReDim Operation.Divident_Multiplicand(0)
                        Operation.Divident_Multiplicand(0) = unkvar
                        n = i
                    End If
                    If eq(n + 1) = "*" Then
                        Operation.Type_of_Operation = "Multiplication"
                        multiplicand = simplify_division_multiplication(unkvar)
                        If Operation.Check_Multiplication = True Then
                            ReDim ret(UBound(eq)) : eq.CopyTo(ret, 0)
                            ret = remove_array_part(ret, k - 2, l + 1)
                            If UBound(Operation.Divident_Multiplicand) = 0 Then
                                multiplicand_end_point = n + 1
                                multiplicand_start_point = n - 1
                            Else
                                multiplicand_end_point = n + 1
                                multiplicand_start_point = n - 1 - Operation.Divident_Multiplicand.Length - 1
                            End If
                            multiplicand = add_brackets_start_end(multiplicand)
                            ret = replace_array_part_with_new_arr(ret, multiplicand, multiplicand_start_point, multiplicand_end_point)
                            ret = remove_unecessry_brackets(ret)
                            check_for_and_simplify_wrt_multiplication = ret
                            Exit Function
                        End If
                    Else
                        Exit For
                    End If
                ElseIf eq(i) = "(" Then
                    Exit For
                End If
            Next i
        End If
        check_for_and_simplify_wrt_multiplication = ret
    End Function
    Function remove_array_part(arr() As String, p1 As Integer, p2 As Integer)
        Dim i As Integer, k As Integer
        Dim ret() As String
        ReDim ret(UBound(arr)) : arr.CopyTo(ret, 0)
        If p1 = 0 Then
            k = 0
        Else
            ReDim Preserve ret(p1)
            k = UBound(ret) + 1
        End If
        For i = p2 To UBound(arr)
            ReDim Preserve ret(k)
            ret(k) = arr(i)
            k = k + 1
        Next i
        remove_array_part = ret
    End Function
    Function get_eqn_part_after_opening_bracket_without_start_end_brackets(eq() As String, i As Integer)
        Dim k As Integer, m As Integer, n As Integer
        Dim ret() As String
        k = i : n = 1 : m = 0
        Do
            k = k + 1
            ReDim Preserve ret(m)
            ret(m) = eq(k)
            m = m + 1
            If eq(k) = "(" Then
                n = n + 1
            ElseIf eq(k) = ")" Then
                n = n - 1
            End If
        Loop Until n = 0
        ReDim Preserve ret(UBound(ret) - 1)
        get_eqn_part_after_opening_bracket_without_start_end_brackets = ret
    End Function
    Function check_multiplier_divisor(unkvar As String)
        Dim temp() As String, o() As String, arr(,) As String, tmpeq() As String, teq() As String
        Dim i As Integer, k As Integer
        Dim chk As Boolean
        'checks if the equation contains unkvar and is a combination of operators "*", "/" & "^" only & number of unkvar is equal to one only
        chk = check_operators_multiplier_divisor(Operation.Divisor_Multiplier, unkvar)
        If chk = False Then
            'create a two dimensional array consisitng of eqution
            arr = add_two_dim_array(arr, Operation.Divisor_Multiplier)
            k = 0
            Do
                tmpeq = get_one_dim_array(arr, k)
                For i = 0 To UBound(tmpeq)
                    If tmpeq(i) = "(" Then
                        temp = get_eqn_part_after_opening_bracket(tmpeq, i)
                        i = i + UBound(temp) - 1
                        'checks if the equation contains unkvar and is a combination of operators "*", "/" & "^" only
                        chk = check_operators_multiplier_divisor(temp, unkvar)
                        If chk = True Then
                            Exit Do
                        Else
                            temp = remove_brackets_start_end(temp)
                            teq = get_compressed_equation(temp)
                            o = get_operators2(teq)
                            If o(0) = "^" Then
                                '2^(COS(b)) -> 2^z1 -> chk=True
                                chk = True
                                Exit Do
                            ElseIf o(0) = "/" Or o(0) = "*" Then
                                '((COS(b))*(b^3)) -> (COS(b))*(b^3) : z1*z2 -> add to arr
                                arr = add_two_dim_array(arr, temp)
                            End If
                        End If
                    ElseIf tmpeq(i) = unkvar Then
                        'b*2 -> b : chk = true
                        chk = True
                        Exit Do
                    End If
                Next i
                k = k + 1
            Loop Until k > UBound(arr, 1)
        End If
        check_multiplier_divisor = chk
    End Function
    Function check_operators_multiplier_divisor(eq() As String, unkvar As String)
        Dim o() As String
        Dim i As Integer, n As Integer, m As Integer
        Dim chk As Boolean
        'checks if the equation contains unkvar and is a combination of operators "*", "/" & "^" only 2*((b^2)/3) -> chk = True
        Dim oper = New String() {"*", "/", "^"}
        If eq.Contains(unkvar) = True Then
            chk = True
            o = get_operators2(eq)
            If Len(Join(o)) > 0 Then
                For i = 0 To UBound(o)
                    If oper.Contains(o(i)) = False Then
                        chk = False
                    End If
                Next i
            End If
        End If
        If chk = True And eq.Contains("^") = True Then
            n = 0 : m = 0
            For i = 0 To UBound(eq)
                If eq(i) = "^" And eq(i) <> "*" Then
                    n = n + 1
                ElseIf eq(i) = unkvar And eq(i) <> "*" Then
                    m = m + 1
                End If
            Next i
            'check that the number of unkvar are greater than 1 e.g. (2^b)*(b^2) -> chk = False
            If n > 1 Or m > 1 Then
                chk = False
            End If
        End If
        check_operators_multiplier_divisor = chk
    End Function
    Function check_for_and_simplify_wrt_division(eq() As String, unkvar As String, k As Integer)
        Dim ret() As String, divo() As String
        Dim divident_start_point As Integer, divident_end_point As Integer
        If UBound(eq) > 0 Then
            If k - 2 >= 0 Then
                '(b+2)/(b^2) : divisor=(b^2) -> divident=(b+2)
                If eq(k - 1) = "/" And eq(k - 2) = ")" Then
                    Operation.Divident_Multiplicand = get_eqn_part_before_closing_bracket(eq, k - 2)
                    Operation.Divident_Multiplicand = remove_brackets_start_end(Operation.Divident_Multiplicand)
                    'check if division posiible
                    Operation.Type_of_Operation = "Division"
                    divo = simplify_division_multiplication(unkvar)
                    divident_start_point = k - 5 - UBound(Operation.Divident_Multiplicand)
                ElseIf eq(k - 1) = "/" And eq(k - 2) = unkvar Then
                    ReDim Operation.Divident_Multiplicand(0)
                    Operation.Divident_Multiplicand(0) = unkvar
                    divident_start_point = k - UBound(Operation.Divident_Multiplicand) - 3
                    Operation.Type_of_Operation = "Division"
                    'b/(2^b) -> divo = ""
                    divo = simplify_division_multiplication(unkvar)
                    divident_end_point = k + UBound(Operation.Divisor_Multiplier) + 1
                End If
            End If
        End If
        If Operation.Check_Division = True Then
            divident_end_point = k + UBound(Operation.Divisor_Multiplier) + 1
            '(b/b)+(1/b) -> 1+(1/b)
            If UBound(divo) = 0 Then
                If eq(divident_start_point) = "(" And eq(divident_end_point) = ")" Then
                    divident_start_point = divident_start_point - 1
                    divident_end_point = divident_end_point + 1
                End If
            End If
            ret = replace_array_part_with_new_arr(eq, divo, divident_start_point, divident_end_point)
        End If

        check_for_and_simplify_wrt_division = ret
    End Function
    Function simplify_wrt_powers(eq() As String, unkvar As String)
        Dim ret() As String, rep() As String, temp() As String, tmpeq() As String, o() As String, pwr1() As String
        Dim pwr As String, num_pwr As String
        Dim i As Integer, k As Integer, p1 As Integer, p2 As Integer
        Dim chk As Boolean, chk_pwr As Boolean, chkb As Boolean
        'Dim app = New Microsoft.Office.Interop.Excel.Application
        ReDim ret(UBound(eq))
        eq.CopyTo(ret, 0)
        i = 0
        Do
            If ret(i) = ")" Then
                '((b*3)^2) -> (b*3)
                temp = get_eqn_part_before_closing_bracket(ret, i)
                p1 = i - temp.Length
                p2 = i
                '(b*3) -> b*3
                tmpeq = remove_brackets_start_end(temp)
                '(b*3)/2 -> z1/2
                tmpeq = get_compressed_equation(tmpeq)
                o = get_operators2(tmpeq)
                If Len(Join(o)) > 0 Then
                    If o(0) = "*" Or o(0) = "/" Or o(0) = "^" Or IsNumeric(join_arr(tmpeq)) = True Then
                        chk = True
                    End If
                End If
                If chk = True And i + 2 <= UBound(ret) Then
                    If ret(i + 1) = "^" And IsNumeric(ret(i + 2)) = True Then
                        p2 = p2 + 3
                        pwr = ret(i + 2)
                        pwr1 = get_power(tmpeq, unkvar)
                        If UBound(pwr1) = 0 And IsNumeric(pwr1(0)) = True Then
                            chk_pwr = True
                            num_pwr = pwr1(0)
                        End If
                        k = 0
                        If chk_pwr = True Then
                            If IsNumeric(join_arr(tmpeq)) = True Then
                                '(-3)^2 -> 9
                                ReDim rep(0)
                                rep(0) = app.Evaluate(join_arr(tmpeq) & "^" & pwr)
                                p1 = p1 - 1
                                p2 = p2 + 1
                            Else
                                For j = 0 To UBound(tmpeq)
                                    If IsNumeric(tmpeq(j)) = True Then
                                        ReDim Preserve rep(k)
                                        If j - 1 >= 0 Then
                                            If tmpeq(j - 1) <> "^" Then
                                                '(b*3)^2 : 3 -> 3^2=9 -> (b^2)*9
                                                rep(k) = app.Evaluate(tmpeq(j) & "^" & pwr)
                                            Else
                                                '(b^3)^2 : 3 -> 3*2=6 -> b^6
                                                rep(k) = app.Evaluate(tmpeq(j) & "*" & pwr)
                                            End If
                                        Else
                                            '(3/b)^2 : 3 -> 3^2=9 -> 9/(b^2)
                                            rep(k) = app.Evaluate(tmpeq(j) & "^" & pwr)
                                        End If
                                    ElseIf tmpeq(j) = unkvar Then
                                        If num_pwr = "1" Then
                                            '(b*3)^2 : b -> (b^2) -> (b^2)*9
                                            ReDim Preserve rep(k + 4)
                                            rep(k) = "(" : rep(k + 1) = unkvar : rep(k + 2) = "^" : rep(k + 3) = pwr : rep(k + 4) = ")"
                                            k = k + 4
                                        Else
                                            '(b^3)^2 : 3 -> 3*2=6 -> b^6
                                            ReDim Preserve rep(k + 2)
                                            rep(k) = unkvar : rep(k + 1) = "^" : rep(k + 2) = app.Evaluate(num_pwr & "*" & pwr)
                                            k = k + 2
                                            j = j + 2
                                        End If
                                    ElseIf Mid$(tmpeq(j), 1, 1) = Chr(158) Then
                                        '((b*3)/4)^2 -> z1/4 -> ((b*3)^2)/16 -> ((b^2)*9)/16
                                        ReDim Preserve rep(k + 4)
                                        rep(k) = "(" : rep(k + 1) = tmpeq(j) : rep(k + 2) = "^" : rep(k + 3) = pwr : rep(k + 4) = ")"
                                        k = k + 4
                                    Else
                                        ReDim Preserve rep(k)
                                        rep(k) = tmpeq(j)
                                    End If
                                    k = k + 1
                                Next j
                            End If
                        Else
                            For j = 0 To UBound(tmpeq)
                                If tmpeq(j) = "^" Then
                                    '(2^COS(b))^3 -> 2^(3*COS(b))
                                    ReDim Preserve rep(k + 3)
                                    rep(k) = tmpeq(j) : rep(k + 1) = "(" : rep(k + 2) = pwr : rep(k + 3) = "*"
                                    rep = add_arr1_to_arr(rep, pwr1)
                                    ReDim Preserve rep(UBound(rep) + 1)
                                    k = UBound(rep)
                                    rep(k) = ")"
                                    k = k + 1
                                    j = j + UBound(pwr1)
                                Else
                                    ReDim Preserve rep(k)
                                    rep(k) = tmpeq(j)
                                    k = k + 1
                                End If
                            Next j
                        End If
                        rep = disassemble_tempeq(temp, rep)
                        ret = replace_array_part_with_new_arr(ret, rep, p1, p2)
                        chkb = check_brackets_start_end(ret)
                        If chkb = True Then
                            ret = remove_brackets_start_end(ret)
                        End If
                        i = 0
                    End If
                End If
                chk = False
            End If
            i = i + 1
        Loop Until i > UBound(ret) Or i <= 0
        simplify_wrt_powers = ret
    End Function
    Function get_power(eq() As String, unkvar As String)
        Dim p As Integer
        Dim n() As String
        ReDim n(0)
        If eq.Contains("^") = True Then
            p = Array.IndexOf(eq, "^")
            If eq(p - 1) = unkvar Then
                n(0) = eq(p + 1)
            Else
                If eq(p + 1) = "(" Then
                    '3^(COS(b)) -> (COS(b))
                    n = get_eqn_part_after_opening_bracket(eq, p + 1)
                Else
                    'b^3 -> b
                    n(0) = eq(p + 1)
                End If
            End If
        Else
            n(0) = "1"
        End If
        get_power = n
    End Function
    Function simplify_division_multiplication(unkvar As String)
        Dim temp() As String, o() As String, main_arr(,) As String, tmpeq() As String, _
        teq() As String, pwr1() As String, ret1() As String, ret2() As String, arr(,) As String, temp_equation() As String, neq1(,) As String, neq2() As String, v() As String
        Dim str As String, ret() As String, base1 As String
        Dim i As Integer, n As Integer, k As Integer, j As Integer
        Dim chk As Boolean, check_operator As Boolean
        'b*(2^b) -> pw_ba=((b,1),(2,b))
        Call get_divisor_multiplier_powers_base(unkvar)
        temp_equation = get_compressed_equation(Operation.Divident_Multiplicand)
        o = get_operators2(temp_equation)
        If Len(Join(o)) > 0 Then
            If o(0) = "*" Or o(0) = "+" Or o(0) = "-" Or o(0) = "^" Or o(0) = "/" Then
                Common_Operator = o(0)
            End If
        ElseIf UBound(temp_equation) = 0 And temp_equation(0) = unkvar Then
            'b/(2^b) -> chk=False
            chk = check_if_power_isnumeric()
            If chk = True Then
                Common_Operator = "^"
            End If
        End If
        If Len(Common_Operator) > 0 Then
            If Common_Operator = "*" Or Common_Operator = "/" Or Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "^" Then
                check_operator = True
                v = get_operands2(temp_equation)
                main_arr = get_two_dim_array(Operation.Divident_Multiplicand)
            End If
        End If
        If check_operator = True Then
            If Common_Operator = "^" Then
                'b^2 -> base=b : power = 2
                base1 = get_base(temp_equation, unkvar)
                pwr1 = get_power(temp_equation, unkvar)
                If IsNumeric(join_arr(pwr1)) = True Then
                    chk = check_if_power_isnumeric()
                    If chk = True Then
                        If Operation.Base = base1 Then
                            If Operation.Type_of_Operation = "Division" Then
                                Operation.Check_Division = True
                            ElseIf Operation.Type_of_Operation = "Multiplication" Then
                                Operation.Check_Multiplication = True
                            End If
                            If Operation.Div_Mul_Part_Position = True Then
                                '(b+1)/(2/b) -> b*(b+1) : (b+1)*(2/b) -> (b+1)/b
                                ret = multiply_divisor_multiplier_and_divident_multiplicand()
                            Else
                                If Operation.Type_of_Operation = "Division" Then
                                    '(b^3)/b -> (b^2)
                                    neq1 = change_powers(unkvar)
                                ElseIf Operation.Type_of_Operation = "Multiplication" Then
                                    '(b^3)*b -> (b^4)
                                    neq2 = change_powers(unkvar)
                                End If
                            End If
                        End If
                    End If
                Else
                    chk = check_base_same(base1)
                    If chk = True Then
                        If Operation.Type_of_Operation = "Division" Then
                            Operation.Check_Division = True
                        ElseIf Operation.Type_of_Operation = "Multiplication" Then
                            Operation.Check_Multiplication = True
                        End If
                        If Operation.Div_Mul_Part_Position = True Then
                            '(2^b)/(5/(2^(COS(b)))) -> ((2^b)*(2^(COS(b))))/5 -> (2^(b+(COS(b))))/5
                            ret = multiply_divisor_multiplier_and_divident_multiplicand()
                        Else
                            '(2^b)/(2^(b*2)) -> 2^(b-(b*2)) : (2^b)*(2^(b*2)) -> 2^(b+(b*2))
                            ret = simplify_non_numeric_powers(pwr1)
                        End If
                    End If
                End If
            Else
                For i = 0 To UBound(v)
                    If Mid$(v(i), 1, 1) = Chr(158) Then
                        n = Mid$(v(i), 2)
                        temp = get_one_dim_array(main_arr, n)
                        arr = add_two_dim_array(arr, temp)
                        k = 0
                        Do
                            tmpeq = get_one_dim_array(arr, k)
                            For j = 0 To UBound(tmpeq)
                                If tmpeq(j) = "(" Or tmpeq(j) = unkvar Then
                                    If tmpeq(j) = "(" Then
                                        'gets part of equation upto closing bracket
                                        temp = get_eqn_part_after_opening_bracket(tmpeq, j)
                                        j = j + UBound(temp) - 1
                                        'check if part of divisor / multiplier can be multiplied / divided
                                        chk = check_operators_multiplier_divisor(temp, unkvar)
                                    ElseIf tmpeq(j) = unkvar Then
                                        chk = True
                                        ReDim temp(0)
                                        temp(0) = unkvar
                                    End If
                                    If chk = True Then
                                        pwr1 = get_power(temp, unkvar)
                                        If IsNumeric(join_arr(pwr1)) = False Then
                                            base1 = get_base(temp, unkvar)
                                            chk = check_base_same(base1)
                                        ElseIf IsNumeric(join_arr(pwr1)) = True Then
                                            'if power of divident is numeric check if power of divisor is also numeric
                                            chk = check_if_power_isnumeric()
                                        End If
                                        If chk = True Then
                                            If Operation.Type_of_Operation = "Division" Then
                                                Operation.Check_Division = True
                                            ElseIf Operation.Type_of_Operation = "Multiplication" Then
                                                Operation.Check_Multiplication = True
                                            End If
                                            If Operation.Div_Mul_Part_Position = True Then
                                                ret = multiply_divisor_multiplier_and_divident_multiplicand()
                                                Exit For
                                            Else
                                                If Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "*" Then
                                                    temp = add_to_divisible_multiplicable_part(temp_equation, v(i))
                                                    ret1 = add_arr1_to_arr(ret1, temp)
                                                    Exit Do
                                                ElseIf Common_Operator = "/" Then
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Else
                                        '((b^2)+3) -> (b^2)+3 -> z1+3 -> o = "+"
                                        temp = remove_brackets_start_end(temp)
                                        teq = get_compressed_equation(temp)
                                        o = get_operators2(teq)
                                        If Len(Join(o)) > 0 Then
                                            If o(0) = "+" Or o(0) = "-" Or o(0) = "*" Or o(0) = "/" Then
                                                arr = add_two_dim_array(arr, temp)
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                            k = k + 1
                        Loop Until k > UBound(arr, 1)
                        Erase arr
                        If chk = False Then
                            If Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "*" Then
                                temp = add_to_non_divisible_non_multiplicable_part(temp_equation, v(i))
                                ret2 = add_arr1_to_arr(ret2, temp)
                            End If
                        End If
                    ElseIf v(i) = unkvar Then
                        chk = check_if_power_isnumeric()
                        If chk = True Then
                            If Operation.Type_of_Operation = "Division" Then
                                Operation.Check_Division = True
                            ElseIf Operation.Type_of_Operation = "Multiplication" Then
                                Operation.Check_Multiplication = True
                            End If
                            If Operation.Div_Mul_Part_Position = True Then
                                '(b+1)/(2/b) -> b*(b+1) : (b+1)*(2/b) -> (b+1)/b
                                ret = multiply_divisor_multiplier_and_divident_multiplicand()
                                Exit For
                            Else
                                '(b^2)+(COS(b))+3 -> (b^2) is divisible_multiplicable_part & (COS(b))+3 is non-divisible_multiplicable_part
                                '(COS(b)-(b^2))/b -> divisible_multiplicable_part = -((b^2)/b)
                                If Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "*" Then
                                    temp = add_to_divisible_multiplicable_part(temp_equation, v(i))
                                    ret1 = add_arr1_to_arr(ret1, temp)
                                ElseIf Common_Operator = "/" Then
                                    Exit For
                                End If
                            End If
                        ElseIf chk = False Then
                            If Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "*" Then
                                temp = add_to_non_divisible_non_multiplicable_part(temp_equation, v(i))
                                ret2 = add_arr1_to_arr(ret2, temp)
                            End If
                        End If
                    Else
                        If Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "*" Then
                            temp = add_to_non_divisible_non_multiplicable_part(temp_equation, v(i))
                            ret2 = add_arr1_to_arr(ret2, temp)
                        End If
                    End If
                Next i
            End If
        End If
        If Operation.Check_Division = True Or Operation.Check_Multiplication = True Then
            If Operation.Div_Mul_Part_Position = False Then
                If Common_Operator = "+" Or Common_Operator = "-" Or Common_Operator = "*" Then
                    ret1 = disassemble_tempeq(Operation.Divident_Multiplicand, ret1)
                    If ret1(0) = "+" Or ret1(0) = "*" Then
                        '+(z1/b) -> (z1/b)
                        ret1 = reduce_arr(ret1, 0)
                    End If
                    If Len(Join(ret2)) > 0 Then
                        If Common_Operator = "+" Or Common_Operator = "-" Then
                            '-z1+z2-3 -> -z1-z2+3
                            If ret2(0) = "-" Then
                                For i = 1 To UBound(ret2)
                                    If ret2(i) = "+" Then
                                        ret2(i) = "-"
                                    ElseIf ret2(i) = "-" Then
                                        ret2(i) = "+"
                                    End If
                                Next i
                            End If
                            '-z1-z2+3 -> -(SIN(b))-(COS(b))+3 -> (SIN(b))-(COS(b))+3 -> ((SIN(b))-(COS(b))+3)
                            ret2 = disassemble_tempeq(Operation.Divident_Multiplicand, ret2)
                            str = ret2(0)
                            ret2 = reduce_arr(ret2, 0)
                            If UBound(ret2) <> 0 Then
                                ret2 = add_brackets_start_end(ret2)
                            End If
                            If Operation.Check_Division = True Then
                                '((SIN(b))-(COS(b))+3) -> ((SIN(b))-(COS(b))+3)/(b*3)
                                ret2 = add_element_to_array(ret2, UBound(ret2) + 1, "/")
                                ret2 = add_arr1_to_arr(ret2, Operation.Div_Mul_Part)
                            Else
                                '((SIN(b))-(COS(b))+3) -> ((SIN(b))-(COS(b))+3)*(b*3)
                                Operation.Div_Mul_Part = add_element_to_array(Operation.Div_Mul_Part, UBound(Operation.Div_Mul_Part) + 1, "*")
                                ret2 = replace_array_part_with_new_arr(ret2, Operation.Div_Mul_Part, -1, 0)
                            End If
                            '((SIN(b))-(COS(b))+3)/(b*3) -> (((SIN(b))-(COS(b))+3)/(b*3)) -> -(((SIN(b))-(COS(b))+3)/(b*3))
                            ret2 = add_brackets_start_end(ret2)
                            ret2 = add_element_to_array(ret2, 0, str)
                        ElseIf Common_Operator = "*" Then
                            ret2 = disassemble_tempeq(Operation.Divident_Multiplicand, ret2)
                        End If
                    End If
                    ret = add_arr1_to_arr(ret, ret1)
                    If Len(Join(ret2)) > 0 Then
                        ret = add_arr1_to_arr(ret, ret2)
                    End If
                ElseIf Common_Operator = "^" And IsNumeric(Operation.String_Power) = True Then
                    If Operation.Check_Division = True Then
                        '(b^2)/(b^5) -> neq(1,1)="1" : neq(1,2)=(b^3) -> 1/(b^3)
                        temp = get_one_dim_array(neq1, 0)
                        ret = add_arr1_to_arr(ret, temp)
                        temp = get_one_dim_array(neq1, 1)
                        If temp(0) <> "1" Then
                            ret = add_element_to_array(ret, UBound(ret) + 1, "/")
                            ret = add_arr1_to_arr(ret, temp)
                        End If
                    ElseIf Operation.Check_Multiplication = True Then
                        ReDim ret(UBound(neq2))
                        neq2.CopyTo(ret, 0)
                    End If
                ElseIf Common_Operator = "^" And IsNumeric(Operation.String_Power) = False Then
                    ret = disassemble_tempeq(Operation.Divident_Multiplicand, ret)
                ElseIf Common_Operator = "/" Then
                    If Operation.Check_Multiplication = True Then
                        '((b+2)/2)*b -> b*z1 -> (b*z1) -> (b*z1)/2 -> (b*(b+2))/2
                        ret = add_arr1_to_arr(ret, Operation.Div_Mul_Part)
                        ret = add_element_to_array(ret, UBound(ret) + 1, "*")
                        ret = add_element_to_array(ret, UBound(ret) + 1, v(0))
                    Else
                        '((b+2)/2)/b -> z1/b -> (z1/b) -> (z1/b)/2 -> ((b+2)/b)/2
                        ret = New String() {v(0)}
                        'ret = add_element_to_array(ret, UBound(ret) + 1, v(0))
                        ret = add_element_to_array(ret, UBound(ret) + 1, "/")
                        ret = add_arr1_to_arr(ret, Operation.Div_Mul_Part)
                    End If
                    ret = add_brackets_start_end(ret)
                    ret = add_element_to_array(ret, UBound(ret) + 1, "/")
                    ret = add_element_to_array(ret, UBound(ret) + 1, v(1))
                    ret = disassemble_tempeq(Operation.Divident_Multiplicand, ret)
                End If
            End If
            If UBound(Operation.Modified_Div_Mul) <> 0 And Operation.Modified_Div_Mul(0) <> "1" Then
                ret = add_brackets_start_end(ret)
                If Operation.Check_Division = True Then
                    '(b+2)/(b*5) -> 1+(2/b) -> (1+(2/b)) -> (1+(2/b))/5
                    ret = add_element_to_array(ret, UBound(ret) + 1, "/")
                ElseIf Operation.Check_Multiplication = True Then
                    '(b+2)*(b*5) -> (b^2)+(b*2) -> (1+(2/b)) -> (1+(2/b))*5
                    ret = add_element_to_array(ret, UBound(ret) + 1, "*")
                End If
                ret = add_arr1_to_arr(ret, Operation.Modified_Div_Mul)
            End If
            ret = simplify_mul_and_div_by_one(ret, unkvar)
        End If
        Common_Operator = ""
        simplify_division_multiplication = ret
    End Function
    Function add_to_non_divisible_non_multiplicable_part(eq() As String, v As String)
        Dim ret() As String
        Dim p As Integer
        ReDim ret(0)
        p = Array.IndexOf(eq, v)
        If p - 1 <= 0 Then
            If Common_Operator = "+" Or Common_Operator = "-" Then
                ret(0) = "+"
            ElseIf Common_Operator = "*" Then
                ret(0) = "*"
            End If
        Else
            ret(0) = eq(p - 1)
        End If
        ret = add_element_to_array(ret, UBound(ret) + 1, v)
        add_to_non_divisible_non_multiplicable_part = ret
    End Function
    Function add_to_divisible_multiplicable_part(eq() As String, v As String)
        Dim ret() As String
        Dim p As Integer
        ReDim ret(0)
        p = Array.IndexOf(eq, v)
        If p - 1 <= -1 Then
            If Common_Operator = "+" Or Common_Operator = "-" Then
                ret(0) = "+"
            ElseIf Common_Operator = "*" Then
                ret(0) = "*"
            End If
        Else
            ret(0) = eq(p - 1)
        End If
        If Operation.Check_Division = True Then
            '((b^5)+2)/b -> z1+2 -> (z1/b)
            ret = add_arr1_to_arr(ret, New String() {"(", v, "/"})
            ret = add_arr1_to_arr(ret, Operation.Div_Mul_Part)
        Else
            '((b^5)+2)*b -> z1+2 -> (b*z1)
            ret = add_element_to_array(ret, UBound(ret) + 1, "(")
            ret = add_arr1_to_arr(ret, Operation.Div_Mul_Part)
            ret = add_arr1_to_arr(ret, New String() {"*", v})
        End If
        ret = add_element_to_array(ret, UBound(ret) + 1, ")")
        add_to_divisible_multiplicable_part = ret
    End Function
    Function simplify_non_numeric_powers(Power() As String)
        Dim ret = New String() {Operation.Base, "^", "("}
        ret = add_arr1_to_arr(ret, Power)
        If Operation.Check_Division = True Then
            '(2^b)/(2^(b*2)) -> 2^(b-(b*2))
            ret = add_element_to_array(ret, UBound(ret) + 1, "-")
        ElseIf Operation.Check_Multiplication = True Then
            '(2^b)*(2^(b*2)) -> 2^(b+(b*2))
            ret = add_element_to_array(ret, UBound(ret) + 1, "+")
        End If
        ret = add_arr1_to_arr(ret, Operation.Power)
        ret = add_element_to_array(ret, UBound(ret) + 1, ")")
        simplify_non_numeric_powers = ret
    End Function
    Function check_base_same(base1 As String)
        Dim chk As Boolean
        Dim i As Integer
        For i = 0 To UBound(Operation.Base_Array)
            If base1 = Operation.Base_Array(i) Then
                Operation.Base = Operation.Base_Array(i)
                Operation.Power = get_one_dim_array(Operation.Power_Array, i)
                Operation.String_Power = join_arr(Operation.Power)
                ReDim Operation.Div_Mul_Part(2)
                Operation.Div_Mul_Part(0) = "(" : Operation.Div_Mul_Part(1) = Operation.Base_Array(i) : Operation.Div_Mul_Part(2) = "^"
                Operation.Div_Mul_Part = add_arr1_to_arr(Operation.Div_Mul_Part, Operation.Power)
                Operation.Div_Mul_Part = add_element_to_array(Operation.Div_Mul_Part, UBound(Operation.Div_Mul_Part) + 1, ")")
                chk = True
                Operation.Modified_Div_Mul = modify_multiplier_divisor()
                Exit For
            End If
        Next i
        check_base_same = chk
    End Function
    Function change_powers(unkvar As String)
        Dim ret1(,) As String, ret2() As String
        Dim chk As Boolean
        Dim pwr As Integer, pwr1 As Integer, pwr2 As Integer, p As Integer
        Dim numerator() As String, denominator() As String
        pwr2 = Operation.String_Power
        p = Array.IndexOf(Operation.Divident_Multiplicand, unkvar)
        If p + 2 <= UBound(Operation.Divident_Multiplicand) Then
            If Operation.Divident_Multiplicand(p + 1) = "^" And IsNumeric(Operation.Divident_Multiplicand(p + 2)) Then
                pwr1 = Operation.Divident_Multiplicand(p + 2)
                chk = True
            Else
                pwr1 = "1"
            End If
        Else
            pwr1 = "1"
        End If
        If Operation.Check_Division = True Then
            numerator = Operation.Divident_Multiplicand
            denominator = Operation.Div_Mul_Part
            If pwr1 > pwr2 Then
                pwr = pwr1 - pwr2
                If pwr = 1 Then
                    '(b^2)/b -> numerator = b
                    numerator = reduce_arr(numerator, p + 1)
                    numerator = reduce_arr(numerator, p + 1)
                Else
                    '(b^4)/(b^2) -> numerator = (b^2)
                    numerator(p + 2) = pwr
                End If
                ReDim denominator(0)
                'denominator = 1
                denominator(0) = "1"
            ElseIf pwr2 > pwr1 Or pwr2 = pwr1 Then
                If chk = True Then
                    'b^2 -> b2 -> b -> 1
                    numerator = reduce_arr(numerator, p + 1)
                    numerator = reduce_arr(numerator, p + 1)
                End If
                numerator(0) = "1"
                If pwr2 > pwr1 Then
                    pwr = pwr2 - pwr1
                    If pwr = 1 Then
                        'b/(b^2) -> denominator = b
                        ReDim denominator(0)
                        denominator(0) = unkvar
                    Else
                        'b/(b^3) -> denominator = (b^2)
                        denominator(3) = pwr2 - pwr1
                    End If
                Else
                    '(b^2)/(b^2) -> denominator = 1
                    ReDim denominator(0)
                    denominator(0) = "1"
                End If
            End If
            ret1 = add_two_dim_array(ret1, numerator)
            ret1 = add_two_dim_array(ret1, denominator)
            change_powers = ret1
        ElseIf Operation.Check_Multiplication = True Then
            pwr = pwr1 + pwr2
            If pwr1 <> "1" And UBound(Operation.Divident_Multiplicand) <> 0 Then
                'b*(b^2) -> (b^3)
                ReDim ret2(UBound(Operation.Divident_Multiplicand))
                Operation.Divident_Multiplicand.CopyTo(ret2, 0)
                ret2(p + 2) = pwr
            Else
                '(b^2)*b -> (b^3)
                ret2 = New String() {unkvar, "^", pwr}
            End If
            change_powers = ret2
        End If
    End Function
    Function multiply_divisor_multiplier_and_divident_multiplicand()
        Dim ret() As String
        Dim c As New Collection
        If Operation.Check_Division = True Then
            '(b+1)/(2/b) -> b*(b+1)
            c = add_array_to_collection(c, Operation.Div_Mul_Part)
            c.Add("*") : c.Add("(")
            c = add_array_to_collection(c, Operation.Divident_Multiplicand)
            c.Add(")")
        ElseIf Operation.Check_Multiplication = True Then
            '(b+1)*(2/b) -> (b+1)/b
            c.Add("(")
            c = add_array_to_collection(c, Operation.Divident_Multiplicand)
            c.Add(")") : c.Add("/")
            c = add_array_to_collection(c, Operation.Div_Mul_Part)
        End If
        ret = convert_collection_to_array(c)
        multiply_divisor_multiplier_and_divident_multiplicand = ret
    End Function
    Function check_if_power_isnumeric()
        Dim chk As Boolean
        Dim i As Integer
        Dim temp() As String
        For i = 0 To UBound(Operation.Power_Array, 1)
            temp = get_one_dim_array(Operation.Power_Array, i)
            If IsNumeric(join_arr(temp)) = True Then
                Operation.Base = Operation.Base_Array(i)
                Operation.Power = temp
                Operation.String_Power = join_arr(temp)
                If Operation.String_Power <> "1" Then
                    Operation.Div_Mul_Part = New String() {"(", Operation.Base_Array(i), "^", Operation.String_Power, ")"}
                Else
                    ReDim Operation.Div_Mul_Part(0)
                    Operation.Div_Mul_Part(0) = Operation.Base_Array(i)
                End If
                Operation.Modified_Div_Mul = modify_multiplier_divisor()
                chk = True
                Exit For
            End If
        Next i
        check_if_power_isnumeric = chk
    End Function
    Function modify_multiplier_divisor()
        Dim ret() As String, strt As String, str As String
        Dim i As Integer, k As Integer, n As Integer, l As Integer
        Dim temp() As String, teq() As String, o() As String
        Dim chkb As Boolean
        Dim mfunc = New String() {"+", "-", "LOG", "LN", "EXP", "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "SQRT"}
        strt = join_arr(Operation.Divisor_Multiplier)
        str = join_arr(Operation.Div_Mul_Part)
        n = 0
        If str = strt Then
            ReDim ret(0)
            ret(0) = "1"
            Operation.Div_Mul_Part_Position = False
        Else
            ReDim ret(UBound(Operation.Divisor_Multiplier))
            Operation.Divisor_Multiplier.CopyTo(ret, 0)
            'ret = Operation.Divisor_Multiplier
            chkb = check_brackets_start_end(ret)
            If chkb = True Then
                ret = remove_brackets_start_end(ret)
            End If
            For i = 0 To UBound(ret)
                If ret(i) = "(" Then
                    temp = get_eqn_part_after_opening_bracket(ret, i)
                    l = i + temp.Length - 1
                    k = i
                    strt = join_arr(temp)
                    If strt = str Then
                        Call check_equation_divisor_position(n)
                        '(b^2)*(COS(b)) : str=(b^2) -> )*(COS(b)) -> 1*(COS(b))
                        For j = 0 To UBound(temp) - 1
                            ret = reduce_arr(ret, k)
                        Next j
                        ret(k) = "1"
                        Exit For
                    Else
                        '(COS(b))*b -> temp=COS(b) -> COSz1 -> o=(COS) -> i will jump to ")"
                        temp = remove_brackets_start_end(temp)
                        teq = get_compressed_equation(temp)
                        o = get_operators2(teq)
                        If mfunc.Contains(o(0)) = True And o(0) <> "*" Then
                            i = l
                        End If
                    End If
                ElseIf ret(i) = str Then
                    Call check_equation_divisor_position(n)
                    ret(i) = "1"
                    Exit For
                ElseIf ret(i) = "/" Then
                    n = n + 1
                End If
            Next i
            If chkb = True Then
                ret = add_brackets_start_end(ret)
            End If
        End If
        modify_multiplier_divisor = ret
    End Function
    Function check_equation_divisor_position(n As Integer)
        If n > 0 Then
            If n Mod 2 <> 0 Then
                Operation.Div_Mul_Part_Position = True
            End If
        Else
            Operation.Div_Mul_Part_Position = False
        End If
    End Function
    Function get_divisor_multiplier_powers_base(unkvar As String)
        Dim temp() As String, o() As String, arr(,) As String, tmpeq() As String, teq() As String, temp1() As String
        Dim p As Integer, i As Integer, n As Integer, k As Integer
        Dim chk As Boolean
        Erase Operation.Power_Array
        'check if only operators in eq are a combination of *, /, ^ & only one copy on unkvar is present
        chk = check_operators_multiplier_divisor(Operation.Divisor_Multiplier, unkvar)
        If chk = True Then
            'divisor / multiplier consist only of "/", "*","^" and/or only one instance of unkvar
            ReDim Operation.Base_Array(0)
            If Operation.Divisor_Multiplier.Contains("^") = True Then
                p = Array.IndexOf(Operation.Divisor_Multiplier, "^")
                Operation.Base_Array(0) = Operation.Divisor_Multiplier(p - 1)
                If Operation.Divisor_Multiplier(p + 1) = "(" Then
                    temp = get_eqn_part_after_opening_bracket(Operation.Divisor_Multiplier, p + 1)
                Else
                    ReDim temp(0)
                    temp(0) = Operation.Divisor_Multiplier(p + 1)
                End If
                Operation.Power_Array = add_two_dim_array(Operation.Power_Array, temp)
            Else
                Operation.Base_Array(0) = unkvar
                ReDim temp(0)
                temp(0) = "1"
                Operation.Power_Array = add_two_dim_array(Operation.Power_Array, temp)
            End If
        ElseIf chk = False Then
            arr = add_two_dim_array(arr, Operation.Divisor_Multiplier)
            k = 0 : n = 0
            Do
                tmpeq = get_one_dim_array(arr, k)
                For i = 0 To UBound(tmpeq)
                    If tmpeq(i) = "(" Then
                        temp = get_eqn_part_after_opening_bracket(tmpeq, i)
                        i = i + UBound(temp) - 1
                        chk = check_operators_multiplier_divisor(temp, unkvar)
                        If chk = True Then
                            ReDim Preserve Operation.Base_Array(n)
                            Operation.Base_Array(n) = get_base(temp, unkvar)
                            n = n + 1
                            temp1 = get_power(temp, unkvar)
                            Operation.Power_Array = add_two_dim_array(Operation.Power_Array, temp1)
                        Else
                            '((b^2)*(3^b)) -> ret=((b^2),(3^b))
                            temp = remove_brackets_start_end(temp)
                            teq = get_compressed_equation(temp)
                            o = get_operators2(teq)
                            If o(0) = "^" Then
                                '2^(COS(b)) -> 2^z1 -> ret=(2,COS(b))
                                ReDim Preserve Operation.Base_Array(n)
                                Operation.Base_Array(n) = get_base(temp, unkvar)
                                n = n + 1
                                temp1 = get_power(temp, unkvar)
                                Operation.Power_Array = add_two_dim_array(Operation.Power_Array, temp1)
                            ElseIf o(0) = "/" Or o(0) = "*" Then
                                arr = add_two_dim_array(arr, temp)
                            End If
                        End If
                    ElseIf tmpeq(i) = unkvar Then
                        ReDim Preserve Operation.Base_Array(n)
                        Operation.Base_Array(n) = unkvar
                        n = n + 1
                        If i + 2 <= UBound(tmpeq) Then
                            If tmpeq(i + 1) = "^" And IsNumeric(tmpeq(i + 2)) = True Then
                                ReDim temp1(0)
                                temp1(0) = tmpeq(i + 2)
                            ElseIf tmpeq(i + 1) = "^" And tmpeq(i + 2) = "(" Then
                                temp1 = get_eqn_part_after_opening_bracket(tmpeq, i + 2)
                            Else
                                ReDim temp1(0)
                                temp1(0) = "1"
                            End If
                        Else
                            ReDim temp1(0)
                            temp1(0) = "1"
                        End If
                        Operation.Power_Array = add_two_dim_array(Operation.Power_Array, temp1)
                    End If
                Next i
                k = k + 1
            Loop Until k > UBound(arr, 1)
        End If
    End Function
    Function get_base(eq() As String, unkvar As String)
        Dim p As Integer
        Dim ret As String
        If eq.Contains("^") = True Then
            'b^3 -> base = b : 2^b -> base = 2
            p = Array.IndexOf(eq, "^")
            ret = eq(p - 1)
        Else
            'b*3 -> base = b
            ret = unkvar
        End If
        get_base = ret
    End Function
End Module