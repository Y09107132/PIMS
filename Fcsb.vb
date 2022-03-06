Imports Aspose.Cells
Imports System.Data.SqlClient
Module Fcsb
    Dim st() As String = Form0.st
    Public cmdstr As String, cmd As SqlCommand, dr As SqlDataReader, blph As Boolean
    Public cnctm As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & st(0) & ";password=" & st(1) & ";timeout=1")
    Function s1(ByRef name As String) As Integer
        Try
            cnctm.Open()
            Select Case CStr(New SqlCommand("select b.name from sys.database_principals a,sys.database_principals b,sys.database_role_members roles where roles.member_principal_id=a.principal_id and roles.role_principal_id=b.principal_id and a.name='" & name & "'", cnctm).ExecuteScalar)
                Case "db_owner"
                    s1 = 1
                Case "Soperator"
                    s1 = 2
                Case "DataReader"
                    s1 = 3
                Case "Joperator"
                    s1 = 4
                Case "MDataReader"
                    s1 = 5
                Case "PDataReader"
                    s1 = 6
                Case Else
                    s1 = CInt(name <> st(0))
            End Select
            cnctm.Close()
        Catch ex As Exception
            s1 = -1
            cnctm.Close()
        End Try
    End Function
    Function s2(mc As List(Of String), ByRef lx As String) As String
        s2 = "("
        For i = 0 To mc.Count - 2 Step 1
            s2 += lx & "='" & Replace(mc.Item(i), "'", "''") & "' or "
        Next
        If mc.Count >= 1 Then
            If mc(mc.Count - 1) = " " Then
                s2 += lx & " is NULL)"
            Else
                s2 += lx & "='" & Replace(mc(mc.Count - 1), "'", "''") & "')"
            End If
        End If
    End Function
    Sub s3(li As ListBox, ByRef cmdstr As String)
        li.Items.Clear()
        If li Is Form1.CL4 OrElse li Is Form1.CL2 Then li.Items.Add("全部")
        Try
            dr = New SqlCommand(cmdstr, Form1.cnct).ExecuteReader
            While dr.Read()
                li.Items.Add(dr(0))
            End While
            dr.Close()
        Catch ex As Exception
            dr.Close()
        End Try
    End Sub
    Sub s4(DA As DataGridView, ByRef str As String, Optional idt As List(Of Integer) = Nothing)
        For Each r As DataGridViewRow In DA.SelectedRows
            If Not r.IsNewRow Then
                If CInt(r.Cells(0).Value) > 0 Then
                    If Form1.sbl(0) OrElse idt IsNot Nothing AndAlso idt.Contains(CInt(r.Cells(0).Value)) Then
                        Try
                            Form1.cnct.Open()
                            Dim i As Integer = New SqlCommand("delete from " & str & " where Id=" & CStr(r.Cells(0).Value), Form1.cnct).ExecuteNonQuery()
                            Form1.cnct.Close()
                        Catch ex As Exception
                            Form1.cnct.Close()
                            DA.ClearSelection()
                            MsgBox("删除记录时有错误发生" & vbCrLf & ex.Message)
                            Return
                        End Try
                        If idt IsNot Nothing Then idt.Remove(CInt(r.Cells(0).Value))
                        For Each dac As DataGridViewCell In r.Cells
                            If Form1.dacl(DA).Contains(dac) Then Form1.dacl(DA).Remove(dac)
                        Next
                        DA.Rows.Remove(r)
                    End If
                Else
                    For Each dac As DataGridViewCell In r.Cells
                        If Form1.dacl(DA).Contains(dac) Then Form1.dacl(DA).Remove(dac)
                    Next
                    DA.Rows.Remove(r)
                End If
            End If
        Next
        DA.ClearSelection()
    End Sub
    Function s5(D1 As DateTimePicker, D2 As DateTimePicker, ByRef field As String) As String
        s5 = "("
        If D1.Checked Then
            If D2.Checked Then
                If D1.Value < D2.Value Then
                    s5 += field & " between '" & D1.Text & "' and '" & D2.Text & "')"
                Else
                    s5 += field & " between '" & D2.Text & "' and '" & D1.Text & "')"
                End If
            Else
                s5 += field & ">'" & D1.Text & "')"
            End If
        Else
            If D2.Checked Then s5 += field & "<'" & D2.Text & "')"
        End If
    End Function
    Sub s6(ByRef cmdstr As String, dgvc As DataGridViewComboBoxColumn)
        dr = New SqlCommand(cmdstr, Form1.cnct).ExecuteReader
        dgvc.Items.Clear()
        While dr.Read()
            dgvc.Items.Add(dr(0))
        End While
        dgvc.Items.Add("")
        dr.Close()
    End Sub
    Sub s7(B1 As Button, B2 As Button, DA As DataGridView, Optional idt As List(Of Integer) = Nothing)
        DA.ReadOnly = B1.Text = "锁定表格"
        B2.Enabled = B1.Text = "解锁表格"
        If B1.Text = "解锁表格" Then
            B1.Text = "锁定表格"
            For i = 0 To DA.Rows.Count - 2
                DA.Rows(i).ReadOnly = Form1.sbl(1) AndAlso idt IsNot Nothing AndAlso Not idt.Contains(CInt(DA.Rows(i).Cells(0).Value)) OrElse CInt(DA.Rows(i).Cells(0).Value) <= 0
            Next
        Else
            B1.Text = "解锁表格"
        End If
        DA.Columns(0).ReadOnly = True
    End Sub
    Sub s8(DA As DataGridView, ByRef bl As Boolean)
        If DA.SelectedRows.Count > 0 Then
            For Each row As DataGridViewRow In DA.SelectedRows
                If Not row.IsNewRow Then
                    For Each dac As DataGridViewCell In row.Cells
                        If Form1.dacl(DA).Contains(dac) Then Form1.dacl(DA).Remove(dac)
                    Next

                    DA.Rows.Remove(row)
                End If
            Next
        ElseIf bl Then
            For Each dar As DataGridViewRow In DA.Rows
                For Each dac As DataGridViewCell In dar.Cells
                    If Form1.dacl(DA).Contains(dac) Then Form1.dacl(DA).Remove(dac)
                Next
            Next
            DA.Rows.Clear()
        End If
        DA.ClearSelection()
    End Sub
    Function s9(ByRef n As Integer, DA As DataGridView, tb As String, ByRef cmdstr0 As String, Optional sqlprmt() As SqlParameter = Nothing) As Boolean
        Dim flag As Boolean
        cmd = New SqlCommand(String.Concat(cmdstr0, "select max(Id) from ", tb), Form1.cnct)
        If sqlprmt IsNot Nothing Then
            For i As Integer = 0 To sqlprmt.Length - 1
                If sqlprmt(i) IsNot Nothing Then cmd.Parameters.Add(sqlprmt(i))
            Next
        End If
        Try
            Form1.cnct.Open()
            DA.Rows(n).Cells(0).Value = cmd.ExecuteScalar
            Form1.cnct.Close()
        Catch ex As Exception
            Form1.cnct.Close()
            flag = True
            DA.Rows(n).Cells(0).Value = 0
            DA.Rows(n).ReadOnly = True
            MsgBox(String.Concat("记录提交未完全成功！" & vbCrLf & "", ex.Message))
        End Try
        Return flag
    End Function
    Function s10(ByRef str As String, ByRef bl As Boolean) As Byte
        Dim num As Byte
        If str IsNot Nothing Then
            Try
                Form1.cnct.Open()
                dr = New SqlCommand("select 正则公式,批号代码 from 批号代码", Form1.cnct).ExecuteReader()
                Dim str1 As String
                While dr.Read()
                    str1 = dr(0).ToString()
                    If bl Then
                        If Text.RegularExpressions.Regex.Match(CStr(IIf(IsNothing(str), "", str)), str1, Text.RegularExpressions.RegexOptions.IgnoreCase).Success Then
                            s28(str.ToCharArray(), 0, str, str1)
                            num = CByte(dr(1))
                            Exit While
                        End If
                    ElseIf Text.RegularExpressions.Regex.Match(CStr(IIf(IsNothing(str), "", str)), str1).Success Then
                        num = CByte(dr(1))
                        Exit While
                    End If
                End While
                Form1.cnct.Close()
            Catch exception As Exception
                Form1.cnct.Close()
            End Try
        End If
        Return num
    End Function
    Sub s11(ByRef id As Integer, DA As DataGridView, Optional ByRef bl As Boolean = True, Optional ByRef idd As Integer = 0, Optional ByRef bln As Boolean = True)
        Dim iid, n As Integer, dic As New Dictionary(Of Integer, String)
        Try
            cnctm.Open()
            cmd = New SqlCommand("select min(Id),count(Id) from 操作记录 where 记录Id=@id and 记录表=@记录表", cnctm)
            cmd.Parameters.Add(New SqlParameter("id", id))
            cmd.Parameters.Add(New SqlParameter("记录表", "物料数量"))
            dr = cmd.ExecuteReader
            While dr.Read
                iid = CInt(dr(0))
                n = CInt(dr(1))
            End While
            dr.Close()
            cmd = New SqlCommand("select SQL语句 from 操作记录 where id=@id", cnctm)
            cmd.Parameters.Add(New SqlParameter("id", iid))
            cmdstr = My.Settings.MR + Replace(CStr(cmd.ExecuteScalar), "insert into 物料数量 values(", "insert into @t values(",, 1) + "select * from @t"
            If bl Then
                DA.Rows.Add()
                idd = DA.Rows.Count - 2
            End If
            DA.Rows(idd).Cells(0).Value = -1
            DA.Rows(idd).Tag = New Integer() {n, id}
            DA.Rows(idd).Cells(0).Tag = dic
            cmd = New SqlCommand(cmdstr, cnctm)
            dr = cmd.ExecuteReader
            While dr.Read
                For i = 1 To dr.FieldCount
                    If (i = 3 OrElse i >= 6) AndAlso Not DirectCast(DA.Columns(i), DataGridViewComboBoxColumn).Items.Contains(IIf(IsDBNull(dr(i - 1)), "", dr(i - 1))) Then
                        If Not bl Then
                            MsgBox("要显示的项目: " & CStr(dr(i - 1)) & " 不在 " & DA.Columns(i).HeaderText & " 列表中！")
                        ElseIf dic.ContainsKey(i) Then
                            dic(i) = CStr(dr(i - 1))
                        Else
                            dic.Add(i, CStr(dr(i - 1)))
                        End If
                        DA.Rows(idd).Cells(i).Value = Nothing
                    Else
                        DA.Rows(idd).Cells(i).Value = IIf(IsDBNull(dr(i - 1)), Nothing, dr(i - 1))
                    End If
                Next
            End While
            cnctm.Close()
            DA.Rows(idd).ReadOnly = True
            If bln Then s45(DA, idd, bln)
        Catch ex As Exception
            cnctm.Close()
        End Try
    End Sub
    Sub s12(CL As CheckedListBox, e As ItemCheckEventArgs)
        Dim k As Integer
        For i = 1 To CL.Items.Count - 1
            If (CL.GetItemChecked(i) AndAlso i <> e.Index) OrElse (e.NewValue = CheckState.Checked AndAlso e.Index = i) Then k += 1
        Next
        If k > 0 Then
            If k = CL.Items.Count - 1 Then
                CL.SetItemCheckState(0, CheckState.Checked)
            Else
                CL.SetItemCheckState(0, CheckState.Indeterminate)
            End If
        Else
            CL.SetItemCheckState(0, CheckState.Unchecked)
        End If
    End Sub
    Function s13(ByRef str As String) As String
        Try
            cnctm.Open()
            cmd = New SqlCommand("批号识别", cnctm)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("批号", str))
            dr = cmd.ExecuteReader
            While dr.Read
                s13 = CStr(dr(0))
            End While
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
        End Try
    End Function
    Sub s14(Bn As Button, ByRef lix As Integer, DA As DataGridView, idt As List(Of Integer), CL As Color)
        DA.ReadOnly = Bn.Text = "解锁表格"
        If lix > -1 Then
            For i = lix To DA.Rows.Count - 2
                If idt.Contains(CInt(DA.Rows(i).Cells(0).Value)) Then
                    DA.Rows(i).Cells(0).Style.BackColor = CL
                    If DA Is Form1.DA1 AndAlso DA.Rows(i).Cells(0).Style.BackColor = Color.DarkViolet Then DA.Rows(i).Cells(0).Style.ForeColor = Color.White
                Else
                    DA.Rows(i).ReadOnly = Form1.sbl(1) OrElse CInt(DA.Rows(i).Cells(0).Value) < 1
                End If
            Next
        End If
    End Sub
    Function s15(ByRef yz() As String, Optional ByRef bl As Boolean = True) As Boolean
        Dim str(4) As String
        str(0) = CStr(IIf(yz(0) = "", "物料名称 is NULL", "物料名称='" & yz(0) & "'"))
        str(1) = CStr(IIf(yz(1) = "", "物料类型 is NULL", "物料类型='" & yz(1) & "'"))
        str(2) = CStr(IIf(yz(2) = "", "操作工序 is NULL", "操作工序='" & yz(2) & "'"))
        str(3) = CStr(IIf(yz(3) = "", "批号代码 is NULL", "批号代码=" & s10(yz(3), Form1.CH32.Checked And bl)))
        str(4) = CStr(IIf(yz(4) = "", "可用釜号 is NULL", "可用釜号='" & yz(4) & "'"))
        cmdstr = "select 1 from 工序类型 where "
        For i = 0 To 4
            cmdstr += str(i) & " and "
        Next
        cmdstr += " 可用性=1"
        Try
            Form1.cnct.Open()
            cmd = New SqlCommand(cmdstr, Form1.cnct)
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                Form1.cnct.Close()
                s15 = s55(yz(3)) <> "" AndAlso s55(yz(3)) <> yz(4)
            Else
                s15 = True
            End If
            Form1.cnct.Close()
        Catch ex As Exception
            Form1.cnct.Close()
            s15 = True
        End Try
    End Function
    Sub s16(DA As DataGridView, ByRef n As Integer, e As DataGridViewCellCancelEventArgs)
        Form1.skip(0) = True
        MsgBox(DA.Columns(n).HeaderText & "输入有误，请检查后重输！")
        e.Cancel = True
        DA.Columns(n).Visible = True
        DA.CurrentCell = DA.Rows(e.RowIndex).Cells(n)
        DA.BeginEdit(False)
    End Sub
    Sub s17(ByRef n As Integer)
        cmd = New SqlCommand("select dbo.储槽计算(@储槽名称,@液位,@时间)", Form1.cnct)
        cmd.Parameters.Add(New SqlParameter("储槽名称", If(Form1.DA2.Rows(n).Cells(0).Tag, Form1.DA2.Rows(n).Cells(2).Value)))
        cmd.Parameters.Add(New SqlParameter("液位", Form1.DA2.Rows(n).Cells(3).Value))
        cmd.Parameters.Add(New SqlParameter("时间", Form1.DA2.Rows(n).Cells(1).Value))
        Try
            Form1.cnct.Open()
            Form1.DA2.Rows(n).Cells(4).Value = IIf(IsDBNull(cmd.ExecuteScalar()), Nothing, cmd.ExecuteScalar())
            Form1.cnct.Close()
        Catch ex As Exception
            Form1.DA2.Rows(n).Cells(4).Value = Nothing
            Form1.cnct.Close()
        End Try
    End Sub
    Sub s18(ByRef n As Integer)
        cmdstr = "select distinct 储槽特性.物料名称,储槽特性.操作工序 from 储槽特性,储槽液位 where 储槽特性.储槽名称=储槽液位.储槽名称 and 储槽液位.储槽名称='" & Replace(CStr(If(Form1.DA2.Rows(n).Cells(0).Tag, Form1.DA2.Rows(n).Cells(2).Value)), "'", "''") & "'"
        Try
            Form1.cnct.Open()
            dr = New SqlCommand(cmdstr, Form1.cnct).ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    Form1.DA2.Rows(n).Cells(5).Value = IIf(IsDBNull(dr(0)), Nothing, dr(0))
                    Form1.DA2.Rows(n).Cells(6).Value = IIf(IsDBNull(dr(1)), Nothing, dr(1))
                End While
            Else
                Form1.DA2.Rows(n).Cells(5).Value = Nothing
                Form1.DA2.Rows(n).Cells(6).Value = Nothing
            End If
            Form1.cnct.Close()
        Catch ex As Exception
            Form1.DA2.Rows(n).Cells(5).Value = Nothing
            Form1.DA2.Rows(n).Cells(6).Value = Nothing
            Form1.cnct.Close()
        End Try
    End Sub
    Function s19(ByRef str As String, ByRef table As String) As String
        If table = "物料数量" OrElse table = "储槽液位" Then
            str = Replace(str, "insert into " & table & " values", "",, 1)
        ElseIf table <> "" Then
            str = Replace(str, "update [" & table & "] ", "",, 1)
            str = Replace(str, "insert into [" & table & "](", "insert into " & table & "(",, 1)
            str = Replace(str, "delete from [" & table & "] where ", "delete from " & table & " where ",, 1)
            str = Replace(str, "[" & Left(table, Len(table) - 2) & "批号]", Left(table, Len(table) - 2) & "批号")
        End If
        str = Replace(str, "'''", "''")
        str = Replace(str, "''", vbBack)
        str = Replace(str, "'", "")
        Return Replace(str, vbBack, "'")
    End Function
    Sub s20(li1 As ListBox, li2 As ListBox, ByRef nm As String, ByRef ky As Boolean)
        Dim i As Integer
        cmdstr = ""
        For i = 0 To li1.Items.Count - 1
            cmdstr += "update " & nm & "特性 set 可用性='" & ky & "' where " & nm & "名称='" & Replace(CStr(li1.Items(i)), "'", "''") & "'"
            li2.Items.Add(li1.Items(i))
        Next i
        Try
            Form1.cnct.Open()
            i = New SqlCommand(cmdstr, Form1.cnct).ExecuteNonQuery()
            Form1.cnct.Close()
            li1.Items.Clear()
        Catch ex As Exception
            Form1.cnct.Close()
            MsgBox("试图改变所有可用性失败!表:" & nm & "特性,程序已退出!")
            Return
        End Try
    End Sub
    Sub s21(li As ListBox, ByRef nm As String)
        Dim i As Integer
        cmdstr = ""
        For i = 0 To li.Items.Count - 1
            cmdstr += "update " & nm & "特性 set id=" & i + 1 & "where " & nm & "名称='" & Replace(CStr(li.Items.Item(i)), "'", "''") & "'"
        Next
        Try
            Form1.cnct.Open()
            i = New SqlCommand(cmdstr, Form1.cnct).ExecuteNonQuery()
            Form1.cnct.Close()
            MsgBox("成功执行！")
        Catch ex As Exception
            Form1.cnct.Close()
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub s22(li As ListBox, ByRef i As Integer)
        Dim itm As String
        If li.SelectedItems.Count = 0 Then Return
        li.SelectionMode = SelectionMode.One
        If li.SelectedItems.Count > 1 OrElse li.SelectedIndex = 0 AndAlso i = -1 OrElse li.SelectedIndex = li.Items.Count - 1 AndAlso i = 1 Then Return
        itm = CStr(li.Items.Item(li.SelectedIndex))
        li.Items.Item(li.SelectedIndex) = li.Items.Item(li.SelectedIndex + i)
        li.Items.Item(li.SelectedIndex + i) = itm
        li.SetSelected(li.SelectedIndex + i, True)
        li.SelectionMode = SelectionMode.MultiExtended
    End Sub
    Function s23() As Object
        Select Case Form1.CH4.Text
            Case "数据库值"
                Return DBNull.Value
            Case "月初转存"
                Return 0
            Case "用户输入"
                Return 1
            Case Else
                Return 2
        End Select
    End Function
    Function s24(Optional ByRef ec As String = "", Optional ByRef ec1 As String = "", Optional ByRef ec2 As String = "", Optional D As DateTimePicker = Nothing, Optional ByRef id As Boolean = False) As String
        Dim str As String
        Try
            Form1.cnct.Open()
            cmd = New SqlCommand("班组识别", Form1.cnct) With
                {.CommandType = CommandType.StoredProcedure}
            If D IsNot Nothing Then
                ec1 = CStr(D.Value)
                cmd.Parameters.Add(New SqlParameter("period", 0))
                cmd.Parameters.Add(New SqlParameter("mode", 1))
            End If
            cmd.Parameters.Add(New SqlParameter("日期", ec1))
            cmd.Parameters.Add(New SqlParameter("操作工序", ec2))
            cmd.Parameters.Add(New SqlParameter("id", id))
            If IsDBNull(cmd.ExecuteScalar) Then
                str = ""
            Else
                str = CStr(cmd.ExecuteScalar)
            End If
            Form1.cnct.Close()
        Catch ex As Exception
            Form1.cnct.Close()
            str = ec
        End Try
        If str = "" Then
            If D IsNot Nothing Then
                str = "N/A"
            Else
                str = Nothing
            End If
        End If
        Return str
    End Function
    Sub s25(DA As DataGridView, ByRef cmdstrn As Integer, ByRef table As String)
        Dim bl As Boolean
        Dim cnctn As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & st(0) & ";password=" & st(1) & ";timeout=1")
        Try
            DA.Columns.Clear()
            cnctm.Open()
            dr = New SqlCommand("select 操作员,操作时间,SQL语句,记录表 as 表,记录Id as Id,Id as RId,计算机名 from 操作记录 where 记录Id=" & cmdstrn & " and 记录表 ='" & table & "'", cnctm).ExecuteReader
            For i = 0 To dr.FieldCount - 1
                DA.Columns.Add(dr.GetName(i), dr.GetName(i))
            Next
            RemoveHandler DA.CellValueChanged, AddressOf Form1.DA11_CellValueChanged
            While dr.Read
                DA.Rows.Add()
                For i = 0 To dr.FieldCount - 1
                    DA.Rows(DA.Rows.Count - 2).Cells(i).Value = s19(CStr(dr(i)), table)
                    If i = 2 AndAlso IsNumeric(s19(CStr(dr(i)), table)) Then
                        bl = True
                        cnctn.Open()
                        DA.Rows(DA.Rows.Count - 2).Cells(2).Tag = New SqlCommand("select SQL语句 from 操作记录 where Id=" & s19(CStr(dr(i)), table), cnctn).ExecuteScalar
                    End If
                Next
            End While
            AddHandler DA.CellValueChanged, AddressOf Form1.DA11_CellValueChanged
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
        Finally
            cnctn.Dispose()
        End Try
        If table <> "" Then
            If table = "物料数量" OrElse table = "储槽液位" Then
                DA.Columns(0).Width = 90
                DA.Columns(1).Width = 130
                DA.Columns(2).Width = 570
            Else
                DA.Columns(0).Width = 90
                DA.Columns(1).Width = 130
                DA.Columns(2).Width = 554
            End If
            For i = 1 To DA.Columns.Count
                DA.Columns(i - 1).Visible = False
            Next
            DA.Columns(0).Visible = True
            DA.Columns(1).Visible = True
            DA.Columns(2).Visible = True
        End If
        s56(DA)
    End Sub
    Function s26(ByRef ml As String, ByRef gx As String, ByRef lx As String, ByRef gd As String, ByRef fd As String) As Boolean
        cnctm.Open()
        cmdstr = "select count(*) from 工序类型 where 物料名称=" & ml & " and 操作工序" & CStr(IIf(gx = "NULL", " is ", "=")) & gx & " and 物料类型" & CStr(IIf(lx = "NULL", " is ", "=")) & lx & " and 批号代码" & CStr(IIf(gd = "NULL", " is ", "=")) & gd & " and 可用釜号" & CStr(IIf(fd = "NULL", " is ", "=")) & fd
        cmd = New SqlCommand(cmdstr, cnctm)
        dr = cmd.ExecuteReader
        While dr.Read
            s26 = CInt(dr(0)) > 0
        End While
        cnctm.Close()
    End Function
    Function s27(ByRef sa As Char()) As String
        For i = 0 To UBound(sa, 1)
            s27 += sa(i)
        Next
    End Function
    Sub s28(ByRef sa As Char(), ByRef i As Integer, ByRef str As String, ByRef dr0 As String)
        Dim num As Integer
        If i <> sa.Length Then
            If sa(i) >= "a" AndAlso sa(i) <= "z" OrElse sa(i) >= "A" AndAlso sa(i) <= "Z" Then
                Dim sb As Char() = CType(sa.Clone(), Char())
                sb(i) = Chr(Asc(sb(i)) + Math.Sign(93 - Asc(sa(i))) * 32)
                num = i + 1
                s28(sb, num, str, dr0)
            End If
            num = i + 1
            s28(sa, num, str, dr0)
        Else
            Dim str1 As String = s27(sa)
            If Text.RegularExpressions.Regex.Match(CStr(IIf(IsNothing(str1), "", str1)), dr0).Success Then
                str = Trim(str1)
                Return
            End If
        End If
    End Sub
    Sub s29()
        Form1.DA6.EndEdit()
        Form1.DA6.Rows.Clear()
        For i = 6 To Form1.DA6.Columns.Count - 1
            Form1.DA6.Columns.RemoveAt(6)
        Next
        Form1.DA6.Rows.Add()
        DirectCast(Form1.DA6.Columns(5), DataGridViewComboBoxColumn).Items.Clear()
        dr = New SqlCommand("select distinct 物料类型 from 工序类型 where 操作工序 is NULL and 可用性=1", Form1.cnct).ExecuteReader
        While dr.Read
            If Form1.LI3.Items.Contains(dr(0)) Then
                DirectCast(Form1.DA6.Columns.Item(5), DataGridViewComboBoxColumn).Items.Add(dr(0))
            End If
        End While
        dr.Close()
        DirectCast(Form1.DA6.Columns.Item(5), DataGridViewComboBoxColumn).Items.Add("")
    End Sub
    Sub s30(ByRef blct As Boolean, ByRef sender As Object, Optional ByRef bl As Boolean = True)
        Dim txtpd, txt As String
        If bl Then
            Form1.lbl(sender)(0) = False
        Else
            Form1.lbl126 = False
            If Not Form1.lbl.ContainsKey(sender) Then Form1.lbl.Add(sender, New Object() {False, Form2.DA1})
        End If
        Form1.SFD.Filter = "Excel 99-03文件|*.xls|Excel 2007文件|*.xlsx|pdf文档|*.pdf"
        Form1.SFD.FileName = CStr(DirectCast(sender, Control).Tag)
        If blct Then
            Do
                If Form1.SFD.ShowDialog = DialogResult.OK Then
                    txt = Form1.SFD.FileName
                    txtpd = Right(LCase(txt), txt.Length - txt.LastIndexOf(".") - 1)
                    If txtpd = "xls" OrElse txtpd = "xlsx" OrElse txtpd = "pdf" Then
                        Exit Do
                    Else
                        Form1.SFD.FileName = CStr(DirectCast(sender, Control).Tag)
                        MsgBox("不支持的文件格式！，请输入正确的扩展名！")
                    End If
                Else
                    Return
                End If
            Loop
        Else
            txt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & CStr(DirectCast(sender, Control).Tag) & ".xls"
        End If
        Dim append As Boolean, xlbook As Workbook, DA As DataGridView = DirectCast(Form1.lbl(sender)(1), DataGridView), a(DA.Columns.Count - 1) As Integer
        If FileIO.FileSystem.FileExists(txt) Then
            Dim msr As MsgBoxResult = MsgBox("文件已存在，是否追加？（不追加将覆盖）", MsgBoxStyle.YesNoCancel)
            If msr = MsgBoxResult.Yes Then
                append = True
            ElseIf msr = MsgBoxResult.Cancel Then
                Return
            End If
        End If
        Try
            If append Then
                xlbook = New Workbook(txt)
            Else
                xlbook = New Workbook
            End If
            Form1.s5(xlbook.Worksheets(0), DA)
            xlbook.Save(txt)
            xlbook = Nothing
            MsgBox("已经导出，名称为：" & txt)
        Catch ex As Exception
            MsgBox(txt & "导出错误！" & vbCrLf & ex.Message)
        End Try
    End Sub
    Sub s32(ct As Control, dt As Dictionary(Of String, String), dit As Dictionary(Of Control, String), ByRef ph As String)
        For Each tb As Control In ct.Controls
            If tb.Controls.Count = 0 Then
                If tb.Tag IsNot Nothing Then
                    Dim tx As TextBox = DirectCast(tb, TextBox)
                    If dt.ContainsKey(CStr(tb.Tag)) Then
                        tx.BackColor = Color.White
                        tx.ForeColor = Color.FromName("WindowText")
                        tx.Text = dt(CStr(tb.Tag))
                    Else
                        tx.BackColor = Color.FromArgb(165, 215, 175)
                        tx.ForeColor = Color.FromName("WindowText")
                        tx.Text = ""
                    End If
                End If
                If TypeOf tb Is TextBox Then
                    Dim i As Integer = Form2.s4(Form2.TC2.SelectedIndex, 0)
                    If i > -1 AndAlso DirectCast(Form2.rg(2, i), TextBox).Text <> ph Then
                        If dit.ContainsKey(tb) Then dit.Remove(tb)
                    End If
                End If
            Else
                s32(tb, dt, dit, ph)
            End If
        Next
    End Sub
    Function s33(ByRef str As String, Tn As TextBox, Bn As Button, ByRef j As Boolean) As Boolean
        Try
            Dim str1 As String = Bn.Text
            Form1.cnct.Open()
            Bn.Text = CStr(New SqlCommand("begin tran select Id from[" & str & "工艺]with(tablockx)where[" & str & "批号]='" & Replace(Tn.Text, "'", "''") & "'", Form1.cnct).ExecuteScalar)
            Form1.cnct.Close()
            If Bn.Text = "" Then
                If str1 <> "" Then
                    Bn.Text = str1
                    Dim msgbr As MsgBoxResult = MsgBox("该" & str & "批次已删除，是否终止更新？", MsgBoxStyle.YesNo)
                    If msgbr = MsgBoxResult.Yes Then
                        Form1.cnct.Open()
                        str1 = CStr(New SqlCommand("begin tran commit tran", Form1.cnct).ExecuteNonQuery)
                        Form1.cnct.Close()
                        Dim dtr() As DataRow = Form1.pdt.Select("BN='" & Replace(Tn.Text, "'", "''") & "'")
                        For Each row As DataRow In dtr
                            Form1.pdt.Rows.Remove(row)
                        Next
                        Bn.Text = ""
                        j = False
                        Return False
                    End If
                End If
                Return True
            Else
                Form1.cnct.Open()
                str1 = CStr(New SqlCommand("begin tran commit tran", Form1.cnct).ExecuteNonQuery)
                Form1.cnct.Close()
            End If
        Catch ex As Exception
            Form1.cnct.Close()
        End Try
    End Function
    Sub s34(tn As TextBox, ByRef rg9 As String, ByRef rg10 As Object, ByRef bl As Boolean)
        Dim msb As MsgBoxResult, str As String, CH32 As Boolean = Form1.CH32.Checked, CH33 As Boolean = Form1.CH33.Checked
        Do
            str = rg9 & "工序批号格式不"
            If CH32 AndAlso s10(tn.Text, True) = 0 Then
                str += "正确!"
            ElseIf CH33 AndAlso (CH32 AndAlso s10(tn.Text, True) <> CByte(rg10) AndAlso s10(tn.Text, True) > 0 OrElse Not CH32 AndAlso s10(tn.Text, False) <> CByte(rg10) AndAlso s10(tn.Text, False) > 0) Then
                str += "匹配!"
            Else
                Exit Do
            End If
            msb = MsgBox(str, DirectCast(MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.DefaultButton1, MsgBoxStyle))
            If msb = MsgBoxResult.Abort Then
                tn.Focus()
                bl = False
                Return
            ElseIf msb = MsgBoxResult.Ignore Then
                Exit Do
            End If
        Loop
    End Sub
    Function s35(ct As Control) As Boolean
        For Each bt As Control In ct.Controls
            If bt.Controls.Count = 0 Then
                If bt.ForeColor = Color.FromName("WindowText") AndAlso bt.Text <> "" AndAlso CStr(bt.Tag) <> "" AndAlso DirectCast(bt, TextBox).BorderStyle = BorderStyle.Fixed3D Then Return True
            ElseIf s35(bt) Then
                Return True
            End If
        Next
    End Function
    Function s36(ByRef strn As String, ByRef obj() As Object, Optional ByRef bl As Boolean = False) As Boolean
        Dim m As Boolean
        Dim str As String = strn
        Try
            If CStr(obj(1)) = "money" Then
                If Not IsNumeric(str) AndAlso Len(str) > 1 AndAlso IsNumeric(Right(str, Len(str) - 1)) Then
                    If Left(str, 1) <> Left(CStr(obj(0)), 1) Then Return bl
                    str = Right(str, Len(str) - 1)
                End If
            End If
            If obj(0).GetType.FullName = "System.DBNull" Then
                Return str = ""
            Else
                If str = CStr(obj(0)) Then Return True
                Select Case obj(0).GetType.FullName
                    Case "System.String"
                        If CStr(obj(1)) = "money" Then
                            m = True
                        Else
                            Return str = CStr(obj(0))
                        End If
                    Case "System.DateTime"
                        Dim d As Date
                        If bl Then
                            Select Case CStr(obj(1))
                                Case "smalldatetime"
                                    Return Date.TryParse(str, d) AndAlso CDate(obj(0)) = CDate(Format(DateAdd(DateInterval.Second, 30, d), "yyyy-MM-dd HH:mm"))
                                Case "date"
                                    Return Date.TryParse(str, d) AndAlso CDate(obj(0)) = CDate(Format(d, "yyyy-MM-dd"))
                                Case Else
                                    Return Date.TryParse(str, d) AndAlso CDate(obj(0)) = d
                            End Select
                        Else
                            Return Date.TryParse(str, d) AndAlso CDate(obj(0)) = d
                        End If
                    Case Else
                        m = True
                End Select
            End If
            If m Then
                Dim b As Boolean
                Dim c As Integer
                If CStr(obj(1)) = "bit" Then
                    If Integer.TryParse(CStr(CInt(CBool(str))), c) Then b = CBool(c)
                    Return CBool(obj(0)) = b
                Else
                    Dim d As Decimal
                    If bl Then
                        Dim decstr As String = CStr(obj(0))
                        If decstr.IndexOf(".") > -1 Then
                            Dim i As Integer = Len(decstr) - decstr.IndexOf(".") - 1
                            decstr = "0."
                            For j = 1 To i
                                decstr += "0"
                            Next
                        Else
                            decstr = "0"
                        End If
                        Return str <> "" AndAlso Decimal.TryParse(Format(CDec(str), decstr), d) AndAlso CDec(obj(0)) = d
                    Else
                        Return Decimal.TryParse(str, d) AndAlso CDec(obj(0)) = d
                    End If
                End If
            End If
        Catch ex As Exception
            strn = str
            Return False
        End Try
    End Function
    Sub s37(bx As TextBox, ByRef num1 As Decimal)
        Dim i, num, num2 As Integer
        Dim ts As Decimal
        Dim str3 As String
        Dim flag As Boolean
        Dim str2 As String = "0."
        Dim tt As String = bx.Text
        Dim st As String = bx.SelectedText
        Dim ss As Integer = bx.SelectionStart
        Dim sl As Integer = bx.SelectionLength
        bx.SelectionLength = Len(RTrim(bx.SelectedText))
        Dim str As String = Left(bx.Text, bx.SelectionStart)
        Dim str1 As String = Right(bx.Text, bx.TextLength - bx.SelectionStart - bx.SelectionLength)
        Try
            If bx.Text <> "" Then
                If IsNumeric(bx.Text) AndAlso Not Right(bx.Text, 1) = "-" AndAlso Not bx.Text.Contains("+") Then
                    If CDec(bx.Text) < 0 AndAlso bx.SelectionStart = 0 AndAlso bx.SelectionLength = 0 Then ss = 1
                    If bx.SelectionLength <= 0 Then
                        If bx.Text.Contains(".") Then
                            If Left(bx.Text, 1) = "." OrElse Left(bx.Text, 2) = "-." Then str2 = "."
                            num2 = Len(bx.Text) - bx.Text.IndexOf(".")
                            i = 2
                            While i <= num2
                                str2 = String.Concat(str2, "0")
                                i += 1
                            End While
                        End If
                        sl = bx.SelectionLength
                        num = If(Not bx.Text.Contains("."), bx.TextLength - ss - sl + 1, bx.Text.IndexOf(".") - ss - sl + 1)
                        If Not bx.Text.Contains(".") Then
                            If CDec(bx.Text) = 0 Then
                                bx.Text = Format(num1, str2)
                            ElseIf CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl) = 0 Then
                                bx.Text = CStr(CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl))
                            ElseIf Not (CDec(bx.Text) < 0 Xor CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl) < 0) Then
                                bx.Text = CStr(CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl))
                            Else
                                bx.Text = CStr(-CDec(bx.Text))
                            End If
                        ElseIf bx.SelectedText.Contains(".") Then
                            If CDec(bx.Text) = 0 Then
                                bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)
                            ElseIf CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)) = 0 Then
                                bx.Text = Format(0, str2)
                            ElseIf Not (CDec(bx.Text) < 0 Xor CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)) < 0) Then
                                bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)
                            Else
                                bx.Text = CStr(-CDec(bx.Text))
                            End If
                        ElseIf CDec(bx.Text) = 0 Then
                            bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") < ss, 1, 0))), str2)
                        ElseIf CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") < ss, 1, 0))), str2)) = 0 Then
                            bx.Text = Format(0, str2)
                        ElseIf Not (CDec(bx.Text) < 0 Xor CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") = ss - 1, 1, 0))), str2)) < 0) Then
                            bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") < ss, 1, 0))), str2)
                        Else
                            bx.Text = Format(-CDec(bx.Text), str2)
                        End If
                        bx.SelectionLength = sl
                        If Not bx.Text.Contains(".") Then
                            bx.SelectionStart = bx.TextLength - num - sl + 1
                        Else
                            bx.SelectionStart = bx.Text.IndexOf(".") - num - sl + 1
                        End If
                    Else
                        flag = True
                    End If
                ElseIf bx.SelectedText = "" AndAlso IsNumeric(Mid(bx.Text, bx.SelectionStart, 1)) Then
                    Dim tl As Integer = bx.TextLength - ss
                    While ss <> 0 AndAlso IsNumeric(Mid(bx.Text, ss, 1))
                        ss -= 1
                    End While
                    If ss > 0 OrElse IsNumeric(Mid(bx.Text, 1, 1)) Then
                        str3 = Mid(bx.Text, ss + 1, bx.SelectionStart - ss)
                        num2 = CInt(Math.Round(CDec(str3) + num1))
                        If num2 < 0 Then
                            str2 = ""
                            Dim length As Integer = str3.Length
                            i = 1
                            While i <= length
                                str2 = String.Concat(str2, "9")
                                i += 1
                            End While
                            bx.Text = String.Concat(Left(bx.Text, ss), str2, Right(bx.Text, bx.TextLength - bx.SelectionStart))
                        ElseIf Left(str3, 1) <> "0" Then
                            bx.Text = String.Concat(Left(bx.Text, ss), CStr(num2), Right(bx.Text, bx.TextLength - bx.SelectionStart))
                        Else
                            str2 = ""
                            Dim length1 As Integer = str3.Length
                            i = 1
                            While i <= length1
                                str2 = String.Concat(str2, "0")
                                i += 1
                            End While
                            bx.Text = String.Concat(Left(bx.Text, ss), Format(num2, str2), Right(bx.Text, bx.TextLength - bx.SelectionStart))
                        End If
                    End If
                    bx.SelectionStart = bx.TextLength - tl
                ElseIf Not Decimal.TryParse(bx.SelectedText, ts) OrElse Right(bx.SelectedText, 1) = "-" OrElse bx.SelectedText.Contains("+") Then
                    str2 = If(bx.SelectedText <> "", Right(bx.SelectedText, 1), Mid(bx.Text, bx.SelectionStart, 1))
                    num = AscW(str2)
                    If num >= 65 AndAlso num <= 90 Then
                        If num <> 65 OrElse num1 <> -1 Then
                            str2 = If(num <> 90 OrElse num1 <> 1, Chr(CInt(num + num1)), "A")
                        Else
                            str2 = "Z"
                        End If
                    ElseIf num >= 97 AndAlso num <= 122 Then
                        If num <> 97 OrElse num1 <> -1 Then
                            str2 = If(num <> 122 OrElse num1 <> 1, Chr(CInt(num + num1)), "a")
                        Else
                            str2 = "z"
                        End If
                    ElseIf num >= 48 AndAlso num <= 57 Then
                        If num <> 48 OrElse num1 <> -1 Then
                            str2 = If(num <> 57 OrElse num1 <> 1, Chr(CInt(num + num1)), "0")
                        Else
                            str2 = "9"
                        End If
                    End If
                    If bx.SelectedText <> "" Then
                        bx.Text = String.Concat(Left(bx.Text, bx.SelectionStart + bx.SelectionLength - 1), str2, Right(bx.Text, Len(bx.Text) - bx.SelectionStart - bx.SelectionLength))
                    Else
                        bx.Text = String.Concat(Left(bx.Text, bx.SelectionStart - 1), str2, Right(bx.Text, Len(bx.Text) - bx.SelectionStart))
                    End If
                    bx.SelectionStart = ss
                    bx.SelectionLength = sl
                Else
                    flag = True
                End If
                If flag Then
                    str3 = Mid(bx.Text, bx.SelectionStart + 1, bx.SelectionLength)
                    If str3.Contains(".") Then
                        num2 = Len(str3) - str3.IndexOf(".")
                        i = 2
                        str2 = "."
                        While i <= num2
                            str2 = String.Concat(str2, "0")
                            i += 1
                        End While
                        num2 = str3.IndexOf(".")
                        If Left(str3, 1) = "-" Then num2 -= 1
                        i = 1
                        While i <= num2
                            str2 = String.Concat("0", str2)
                            i += 1
                        End While
                        If CDec(str3) = 0 AndAlso num1 = -1 Then
                            If str3.Contains("-") Then
                                str3 = Left(str3, Len(str3) - 1) + "1"
                            Else
                                str3 = Right(String.Concat(Format(CDec(String.Concat("1", str3)) - Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), CStr(IIf(Right(str3, 1) = ".", ".", ""))), Len(str3))
                            End If
                        ElseIf num1 <> -1 Then
                            str3 = If(Right(str3, 1) <> ".", Right(Format(CDec(str3) + Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), Len(str3)), Right(String.Concat(Format(CDec(str3) + Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), "."), Len(str3)))
                        Else
                            str3 = Right(String.Concat(Format(CDec(str3) - Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), CStr(IIf(Right(str3, 1) = ".", ".", ""))), CInt(IIf(str3.Contains("-"), Len(str3) + 1, Len(str3))))
                        End If
                    ElseIf CDec(str3) = 0 AndAlso num1 = -1 Then
                        str3 = CStr(CDec(String.Concat("1", str3)) - 1)
                    ElseIf num1 = -1 Then
                        str2 = ""
                        num2 = Len(str3)
                        If str3.Contains("-") Then num2 -= 1
                        i = 1
                        While i <= num2
                            str2 = String.Concat(str2, "0")
                            i += 1
                        End While
                        str3 = Format(CDec(str3) - 1, str2)
                    ElseIf Len(CStr(CDec(str3) + 1)) <> Len(str3) + 1 Then
                        str2 = ""
                        num2 = Len(str3)
                        If str3.Contains("-") Then num2 -= 1
                        If Math.Log10(-CDec(str3)) Mod 1 = 0 Then num2 -= 1
                        i = 1
                        While i <= num2
                            str2 = String.Concat(str2, "0")
                            i += 1
                        End While
                        str3 = Format(CDec(str3) + 1, str2)
                    Else
                        num2 = Len(str3)
                        str3 = ""
                        i = 1
                        While i <= num2
                            str3 = String.Concat(str3, "0")
                            i += 1
                        End While
                    End If
                    bx.Text = String.Concat(str, str3, str1)
                    bx.SelectionStart = ss
                    bx.SelectionLength = sl
                    Dim k As Boolean
                    If Decimal.TryParse(st, ts) Then
                        If ts < 0 Then
                            If Math.Log10(-CDec(str3)) Mod 1 = 0 AndAlso num1 < 0 Then
                                bx.SelectionLength += 1
                            ElseIf num1 > 0 Then
                                str3 = CStr(-CDec(str3))
                                For j = 1 To Len(str3)
                                    If Mid(str3, j, 1) <> "9" AndAlso Mid(str3, j, 1) <> "." Then
                                        k = True
                                        Exit For
                                    End If
                                Next
                                If Not k OrElse CDec(str3) = 0 Then
                                    If bx.SelectionLength = sl Then bx.SelectionLength -= 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                bx.Text = "0"
                bx.SelectionStart = 1
            End If
        Catch ex As Exception
            bx.Text = tt
        End Try
    End Sub
    Sub s38(dit As Dictionary(Of Control, String), bt As Control, tx As TextBox)
        Dim str As String
        If bt IsNot tx Then
            Dim strArrays As String() = CStr(bt.Tag).Split(CChar("|"))
            If strArrays(1) <> "物料数量" Then
                str = If(strArrays(1) <> "物料含量", "max(", "avg(")
            Else
                str = "sum("
            End If
            Try
                cnctm.Open()
                cmdstr = String.Concat(New String() {"select ", str, strArrays(1), ") from 物料数量 where 批号=@批号 and 物料名称='", strArrays(0), "' and 物料类型='", strArrays(2), "'"})
                If UBound(strArrays) = 3 Then cmdstr += " and 反应釜号 in" + strArrays(3)
                cmd = New SqlCommand(cmdstr, cnctm)
                cmd.Parameters.AddWithValue("批号", tx.Text)
                bt.Text = CStr(IIf(IsDBNull(cmd.ExecuteScalar), "", cmd.ExecuteScalar))
                If dit.ContainsKey(bt) Then
                    dit(bt) = bt.Text
                Else
                    dit.Add(bt, bt.Text)
                End If
                bt.BackColor = Color.FromArgb(165, 215, 175)
                cnctm.Close()
            Catch ex As Exception
                cnctm.Close()
            End Try
        End If
    End Sub
    Sub s39(ByRef blct As Boolean)
        Dim T52D, T53D, T39D, T40D As Decimal, T52B, T53B, T39B, T40B As Boolean, k0, k1, k2, k3, k4 As New List(Of String)
        If Form1.LI1.Items.Count = 0 AndAlso Form1.LI2.Items.Count = 0 Then Return
        If Form1.LI3.Items.Count = 0 AndAlso Form1.LI4.Items.Count = 0 Then Return
        Dim lix As Integer = Form1.DA1.Rows.Count - 1
        Dim cmdstr6 As String = "(" : Dim cmdstr7 As String = "("
        Dim cmdstr8 As String = "("
        s54(Form1.LI1, Form1.LI2, k0) : s54(Form1.LI3, Form1.LI4, k1)
        If Form1.CL2.CheckedItems.Count = 0 AndAlso Form1.LI6.Items.Count = 0 Then
            k2.Add(" ")
        Else
            s54(Form1.LI5, Form1.LI6, k2)
        End If
        For x = 1 To Form1.T1.Items.Count
            If UCase(Form1.T1.Text).Contains(CStr(Form1.T1.Items(x - 1))) Then k3.Add(CStr(Form1.T1.Items(x - 1)))
        Next
        For x = 1 To Form1.CO9.Items.Count
            If UCase(Form1.CO9.Text).Contains(CStr(Form1.CO9.Items(x - 1))) Then k4.Add(CStr(Form1.CO9.Items(x - 1)))
        Next
        Dim cmdstr0 As String = s2(k0, "物料名称") : Dim cmdstr1 As String = s2(k1, "物料类型")
        Dim cmdstr2 As String = s2(k2, "操作工序") : Dim cmdstr3 As String = s2(k3, "班别班组")
        Dim cmdstr4 As String = s2(k4, "反应釜号") : Dim cmdstr5 As String = s5(Form1.D1, Form1.D2, "日期")
        If Form1.T38.Text <> "" Then
            If Form1.LI15.SelectedItems.Count = 0 OrElse Form1.LI15.Visible = False Then
                cmdstr6 = CStr(IIf(blph, "批号='" & Replace(Form1.T38.Text, "'", "''") & "'", "批号 like '%" & Replace(Form1.T38.Text, "'", "''") & "%'"))
            Else
                cmdstr6 = "(批号 like '%"
                For k = 0 To Form1.LI15.SelectedItems.Count - 2
                    cmdstr6 += Replace(CStr(Form1.LI15.SelectedItems(k)), "'", "''") & "%' or  批号 like '%"
                Next
                cmdstr6 += Replace(CStr(Form1.LI15.SelectedItems(Form1.LI15.SelectedItems.Count - 1)), "'", "''") & "%')"
            End If
        End If
        Form1.T52.Text = s49(Form1.T52.Text, T52B, T52D) : Form1.T52.Tag = Form1.T52.Text
        Form1.T53.Text = s49(Form1.T53.Text, T53B, T53D) : Form1.T53.Tag = Form1.T53.Text
        If T52B Then
            If T53B Then
                cmdstr7 = "(物料数量 between " & Math.Min(T52D, T53D) & " and " & Math.Max(T52D, T53D) & ")"
            Else
                cmdstr7 = "(物料数量>=" & T52D & ")"
            End If
        ElseIf T53B Then
            cmdstr7 = "(物料数量<=" & T53D & ")"
        End If
        Form1.T39.Text = s49(Form1.T39.Text, T39B, T39D) : Form1.T39.Tag = Form1.T39.Text
        Form1.T40.Text = s49(Form1.T40.Text, T40B, T40D) : Form1.T40.Tag = Form1.T40.Text
        If T39B Then
            If T40B Then
                cmdstr8 = "(物料含量 between " & Math.Min(T39D, T40D) & " and " & Math.Max(T39D, T40D) & ")"
            Else
                cmdstr8 = "(物料含量>=" & T39D & ")"
            End If
        ElseIf T40B Then
            cmdstr8 = "(物料含量<=" & T40D & ")"
        End If
        Application.DoEvents()
        cmdstr = "select * from 物料数量 where"
        If cmdstr0 <> "(" Then cmdstr += " and " & cmdstr0
        If cmdstr1 <> "(" Then cmdstr += " and " & cmdstr1
        If cmdstr2 <> "(" Then cmdstr += " and " & cmdstr2
        If cmdstr3 <> "(" Then cmdstr += " and " & cmdstr3
        If cmdstr4 <> "(" Then cmdstr += " and " & cmdstr4
        If Not blct Then
            If cmdstr5 <> "(" Then cmdstr += " and " & cmdstr5
        End If
        If cmdstr6 <> "(" Then cmdstr += " and " & cmdstr6
        If cmdstr7 <> "(" Then cmdstr += " and " & cmdstr7
        If cmdstr8 <> "(" Then cmdstr += " and " & cmdstr8
        If Not blct Then cmdstr += " order by 日期"
        cmdstr = Left(cmdstr, 19) + Replace(cmdstr, "where and", "where", 20, 1)
        If InStr(26, cmdstr, "and") = 0 Then cmdstr = Left(cmdstr, 19) + Replace(cmdstr, " where", "", 20, 1)
        If blct Then cmdstr = "select * from (select top 6 * from (" & cmdstr & ") as A order by 日期 desc) as B order by 日期"
        Try
            Form1.cnct.Open()
            dr = New SqlCommand(cmdstr, Form1.cnct).ExecuteReader
            While dr.Read
                Form1.DA1.Rows.Add()
                For i = 0 To 9
                    Form1.DA1.Rows(Form1.DA1.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                Next
                Form1.DA1.Rows(Form1.DA1.Rows.Count - 2).Cells(1).Value = Format(CDate(dr(1)), "yyyy-MM-dd HH:mm")
            End While
            Form1.cnct.Close()
        Catch ex As Exception
            Form1.cnct.Close()
            MsgBox("查询过程中有错误。" & vbCrLf & ex.Message)
            Return
        End Try
        Form1.LI15.Hide() : s14(Form1.B16, lix, Form1.DA1, Form1.idt1, Color.DarkViolet) : Form1.DA1.ClearSelection()
        Form1.ctbl = False
    End Sub
    Sub s40(dit As Dictionary(Of Control, String), T As Control, TM As TextBox)
        For Each ct As Control In T.Controls
            If ct.Controls.Count = 0 Then
                If TypeOf ct Is TextBox AndAlso DirectCast(ct, TextBox).Tag IsNot Nothing AndAlso CStr(DirectCast(ct, TextBox).Tag).Contains("|") Then
                    s38(dit, ct, TM)
                End If
            Else
                s40(dit, ct, TM)
            End If
        Next
    End Sub
    Sub s41(tp As Control)
        If tp.Controls.Count = 0 Then
            If TypeOf tp Is TextBox AndAlso tp.Tag IsNot Nothing Then
                AddHandler tp.TextChanged, AddressOf Form2.TC_TextChanged
                AddHandler tp.LostFocus, AddressOf Form2.T_LostFocus
            End If
        Else
            For Each ct As Control In tp.Controls
                s41(ct)
            Next
        End If
    End Sub
    Sub s42(ByRef blct As Boolean)
        Dim k0, k1 As New List(Of String)
        cmdstr = "select * from 原料入库 where"
        s44(Form2.CB1, k0)
        s44(Form2.CB2, k1)
        Dim cmdstr1 As String = s2(k0, "检验人员")
        Dim cmdstr2 As String = s2(k1, "原料名称")
        Dim cmdstr3 As String = "原料批号 like '%" & Replace(Form2.T32.Text, "'", "''") & "%'"
        If cmdstr1 <> "(" Then cmdstr += " and " & cmdstr1
        If cmdstr2 <> "(" Then cmdstr += " and " & cmdstr2
        If blct Then
            cmdstr = Left(cmdstr, 17) + Replace(cmdstr, " where and", " where", 18, 1)
            If Right(cmdstr, 5) = " where" Then cmdstr = Left(cmdstr, Len(cmdstr) - 6)
            cmdstr = " select * from(" & Replace(cmdstr, "*", "top 6 *") & " order by 原料批号 desc)A order by 原料批号"
        Else
            If cmdstr3 <> "(" Then cmdstr += " and " & cmdstr3
            cmdstr = Left(cmdstr, 17) + Replace(cmdstr, "where and", "where", 18, 1)
            If Right(cmdstr, 5) = "where" Then cmdstr = Left(cmdstr, Len(cmdstr) - 6)
            cmdstr += " order by 原料批号"
        End If
        Dim lix As Integer = Form2.DA1.Rows.Count - 1
        Try
            cnctm.Open()
            cmd = New SqlCommand(cmdstr, cnctm)
            dr = cmd.ExecuteReader
            While dr.Read
                If DirectCast(Form2.DA1.Columns(1), DataGridViewComboBoxColumn).Items.Contains(dr(1)) Then
                    Form2.DA1.Rows.Add()
                    For i = 0 To dr.FieldCount - 1
                        Form2.DA1.Rows(Form2.DA1.Rows.Count - 2).Cells(i).Value = dr(i)
                    Next
                End If
            End While
            cnctm.Close()
            s14(Form2.B79, lix, Form2.DA1, Form1.idt3, Color.Pink)
        Catch ex As Exception
            cnctm.Close()
            MsgBox(ex.Message)
        End Try
        Form2.DA1.ClearSelection()
        Form1.ccbl2 = False
    End Sub
    Function s43(ByRef datestr1 As String, ByRef datestr2 As String) As String
        Dim date1, date2 As Date
        Try
            date1 = CDate(datestr1)
            date2 = CDate(datestr2)
        Catch ex As Exception
            Return "N/A"
        End Try
        Dim m As Integer = CInt(Math.Round(DateDiff(DateInterval.Minute, date1, date2)))
        If Math.Abs(m) >= 1440 Then
            s43 = m \ 1440 & "天" & (Math.Abs(m) - (Math.Abs(m) \ 1440) * 1440) \ 60 & "时" & Math.Abs(m) - (Math.Abs(m) \ 60) * 60 & "分"
        ElseIf Math.Abs(m) >= 60 Then
            s43 = m \ 60 & "时" & Math.Abs(m) - (Math.Abs(m) \ 60) * 60 & "分"
        Else
            s43 = m & "分"
        End If
    End Function
    Sub s44(CB As ComboBox, k As List(Of String))
        If CB.Text = "" Then
            For i = 1 To CB.Items.Count - 1
                k.Add(CStr(CB.Items(i - 1)))
            Next
        Else
            k.Add(CB.Text)
        End If
    End Sub
    Sub s45(DA As DataGridView, ByRef er As Integer, Optional ByRef bl As Boolean = True)
        Dim n As Integer, xn As Decimal, yn As Boolean, sql As String
        DA.ClearSelection()
        If CInt(DA.Rows(er).Cells(0).Value) <> 0 Then
            Do
                DA.Rows(er).Cells(0).Value = CInt(DA.Rows(er).Cells(0).Value) - 1
                If CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 0 OrElse DirectCast(DA.Rows(er).Tag, Integer())(0) = Nothing Then
                    s11(DirectCast(DA.Rows(er).Tag, Integer())(1), DA, False, er, False)
                    DA.Rows(er).Cells(s50(sql, DirectCast(DA.Rows(er).Tag, Integer())(0) - 1, DirectCast(DA.Rows(er).Tag, Integer())(1), "物料数量")).Selected = True
                    Return
                End If
                n = s50(sql, DirectCast(DA.Rows(er).Tag, Integer())(0) + CInt(DA.Rows(er).Cells(0).Value) + 1, DirectCast(DA.Rows(er).Tag, Integer())(1), "物料数量")
                Dim needsql As String = Right(sql, sql.Length - sql.IndexOf("=") - 1)
                If Not IsNumeric(sql) Then
                    needsql = Left(needsql, needsql.LastIndexOf(" where Id="))
                    If needsql = "NULL" Then needsql = Nothing
                    If n = 3 OrElse n >= 6 Then
                        If DirectCast(DA.Columns(n), DataGridViewComboBoxColumn).Items.Contains(If(needsql, "")) Then
                            DA.Rows(er).Cells(n).Value = needsql
                            If DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).ContainsKey(n) Then
                                DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).Remove(n)
                            End If
                        Else
                            DA.Rows(er).Cells(n).Value = Nothing
                            If Not bl Then
                                MsgBox("要显示的项目: " & needsql & " 不在 " & DA.Columns(n).HeaderText & " 列表中！")
                            ElseIf DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).ContainsKey(n) Then
                                DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String))(n) = needsql
                            Else
                                DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).Add(n, needsql)
                            End If
                        End If
                    Else
                        DA.Rows(er).Cells(n).Value = needsql
                    End If
                    If CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) > 1 Then DA.Rows(er).Cells(s50(sql, DirectCast(DA.Rows(er).Tag, Integer())(0) + CInt(DA.Rows(er).Cells(0).Value), DirectCast(DA.Rows(er).Tag, Integer())(1), "物料数量")).Selected = True
                    s49(CStr(DA.Rows(er).Cells(4).Value), yn, xn)
                    If yn Then DA.Rows(er).Cells(4).Value = xn
                    yn = False
                    s49(CStr(DA.Rows(er).Cells(5).Value), yn, xn)
                    If yn Then DA.Rows(er).Cells(5).Value = xn
                End If
            Loop Until Not bl OrElse bl AndAlso CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 1
            If bl AndAlso CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 1 Then
                For Each key In DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).Keys
                    MsgBox("要显示的项目: " & DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String))(key) & " 不在 " & DA.Columns(key).HeaderText & " 列表中！")
                Next
            End If
        End If
    End Sub
    Sub s46(DA As DataGridView, ByRef er As Integer, Optional ByRef bl As Boolean = True)
        Dim sql As String, n As Integer, xn As Decimal
        If CInt(DA.Rows(er).Cells(0).Value) <> 0 Then
            Do
                DA.ClearSelection()
                DA.Rows(er).Cells(0).Value = CInt(DA.Rows(er).Cells(0).Value) - 1
                If CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 0 OrElse DirectCast(DA.Rows(er).Tag, Integer())(0) = Nothing Then
                    s47(DirectCast(DA.Rows(er).Tag, Integer())(1), DA, False, er, False)
                    DA.Rows(er).Cells(s50(sql, DirectCast(DA.Rows(er).Tag, Integer())(0) - 1, DirectCast(DA.Rows(er).Tag, Integer())(1), "储槽液位")).Selected = True
                    Return
                End If
                n = s50(sql, DirectCast(DA.Rows(er).Tag, Integer())(0) + CInt(DA.Rows(er).Cells(0).Value) + 1, DirectCast(DA.Rows(er).Tag, Integer())(1), "储槽液位")
                Dim needsql As String = Right(sql, sql.Length - sql.IndexOf("=") - 1)
                needsql = Left(needsql, needsql.LastIndexOf(" where Id="))
                If Not IsNumeric(sql) Then
                    Select Case n
                        Case 1
                            DA.Rows(er).Cells(1).Value = needsql
                        Case 2
                            If DirectCast(DA.Columns(2), DataGridViewComboBoxColumn).Items.Contains(needsql) Then
                                DA.Rows(er).Cells(2).Value = needsql
                                DA.Rows(er).Cells(0).Tag = Nothing
                            Else
                                DA.Rows(er).Cells(2).Value = Nothing
                                DA.Rows(er).Cells(0).Tag = needsql
                            End If
                        Case 3
                            s49(needsql, True, xn)
                            DA.Rows(er).Cells(3).Value = xn
                    End Select
                    If CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) > 1 Then DA.Rows(er).Cells(s50(sql, DirectCast(DA.Rows(er).Tag, Integer())(0) + CInt(DA.Rows(er).Cells(0).Value), DirectCast(DA.Rows(er).Tag, Integer())(1), "储槽液位")).Selected = True
                    s17(er)
                    s18(er)
                    If DA.Rows(er).Cells(0).Tag IsNot Nothing AndAlso Not bl Then MsgBox("要显示的项目: " & needsql & " 不在 储槽名称 列表中！")
                End If
            Loop Until Not bl OrElse bl AndAlso CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 1
            If bl AndAlso CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 1 Then
                If DA.Rows(er).Cells(0).Tag IsNot Nothing Then MsgBox("要显示的项目: " & CStr(DA.Rows(er).Cells(0).Tag) & " 不在 储槽名称 列表中！")
            End If
        End If
    End Sub
    Sub s47(ByRef id As Integer, DA As DataGridView, Optional ByRef bl As Boolean = True, Optional ByRef idd As Integer = 0, Optional ByRef bln As Boolean = True)
        Dim iid, n As Integer
        Try
            cnctm.Open()
            cmd = New SqlCommand("select min(Id),count(Id) from 操作记录 where 记录Id=@id and 记录表=@记录表", cnctm)
            cmd.Parameters.Add(New SqlParameter("id", id))
            cmd.Parameters.Add(New SqlParameter("记录表", "储槽液位"))
            dr = cmd.ExecuteReader
            While dr.Read
                iid = CInt(dr(0))
                n = CInt(dr(1))
            End While
            dr.Close()
            cmd = New SqlCommand("select SQL语句 from 操作记录 where id=@id", cnctm)
            cmd.Parameters.Add(New SqlParameter("id", iid))
            cmdstr = My.Settings.TR + Replace(CStr(cmd.ExecuteScalar), "insert into 储槽液位 values(", "insert into @t values(",, 1) + "select * from @t"
            If bl Then
                DA.Rows.Add()
                idd = DA.Rows.Count - 2
            End If
            DA.Rows(idd).Cells(0).Value = -1
            DA.Rows(idd).Cells(0).Style.ForeColor = Color.Black
            DA.Rows(idd).Tag = New Integer() {n, id}
            cmd = New SqlCommand(cmdstr, cnctm)
            dr = cmd.ExecuteReader
            While dr.Read
                If Not DirectCast(DA.Columns(2), DataGridViewComboBoxColumn).Items.Contains(dr(1)) Then
                    DA.Rows(idd).Cells(2).Value = Nothing
                    DA.Rows(idd).Cells(0).Tag = dr(1)
                Else
                    DA.Rows(idd).Cells(0).Tag = Nothing
                    DA.Rows(idd).Cells(2).Value = dr(1)
                End If
                DA.Rows(idd).Cells(1).Value = dr(0)
                DA.Rows(idd).Cells(3).Value = dr(2)
            End While
            s17(idd)
            s18(idd)
            If DA.Rows(idd).Cells(0).Tag IsNot Nothing AndAlso Not bln Then MsgBox("所选储槽: " & CStr(DA.Rows(idd).Cells(0).Tag) & " 不在储槽列表中！")
            cnctm.Close()
            DA.Rows(idd).Cells(3).Value = CDec(Format(DA.Rows(idd).Cells(3).Value, "0.000"))
            DA.Rows(idd).ReadOnly = True
            If bln Then s46(DA, idd, bln)
        Catch ex As Exception
            cnctm.Close()
        End Try
    End Sub
    Function s48(ByRef original As String, Optional ByRef offset As String = Nothing) As String
        Dim meta, meta1, meta2, offset1 As String, rst As Date
        original = Trim(original)
        original = original.Replace("：", ":")
        meta = original
        Dim i As Integer = meta.IndexOf(" ")
        meta = Left(meta, i + 1) & Replace(meta, ".", ":", i + 2, 2)
        meta2 = meta
        If meta <> "" Then
            If InStr(meta, ":") = 0 Then
                meta = meta + CStr(IIf(IsNothing(offset), "", " " + offset))
            ElseIf InStr(meta, "/") = 0 AndAlso InStr(meta, "-") = 0 AndAlso InStr(meta, ".") = 0 Then
                meta = Format(Now(), "yyyy-MM-dd ") & meta
                If Date.TryParse(meta, rst) Then
                    If CDate(meta) > Now() Then
                        meta1 = meta
                        meta = CStr(DateAdd(DateInterval.Day, -1, CDate(meta)))
                    End If
                End If
            End If
        ElseIf offset IsNot Nothing Then
            meta = Format(Now(), "yyyy-MM-dd") + CStr(IIf(IsNothing(offset), "", " " + offset))
        End If
        If IsNothing(offset) Then
            offset = If(meta1, meta)
        Else
            offset1 = " " + offset
        End If
        If Date.TryParse(meta, rst) Then
            If offset.LastIndexOf(":") = -1 Then
                s48 = Format(rst, "yyyy-MM-dd")
            ElseIf offset.LastIndexOf(":") = offset.IndexOf(":") AndAlso meta2.LastIndexOf(":") = meta2.IndexOf(":") Then
                s48 = Format(DateAdd(DateInterval.Second, 30, rst), "yyyy-MM-dd HH:mm")
            Else
                s48 = Format(rst, "yyyy-MM-dd HH:mm:ss")
            End If
        ElseIf Date.TryParse(original, rst) AndAlso CDate(original) = CDate(Format(CDate(original), "yyyy-MM-dd")) Then
            s48 = Format(CDate(original), "yyyy-MM-dd") + offset1
        Else
            s48 = original
        End If
    End Function
    Function s49(ByRef original As String, ByRef yn As Boolean, ByRef meta As Decimal) As String
        Try
            original = Replace(original, "--", "")
            original = Replace(original, "（", "(")
            original = Replace(original, "）", ")")
            If original = "" Then Return ""
            cmdstr = "select " & original
            Form1.cnctk.Open()
            cmd = New SqlCommand(cmdstr, Form1.cnctk)
            dr = cmd.ExecuteReader
            s49 = Replace(Replace(original, dr.GetName(0), ""), " ", "")
            While dr.Read
                yn = Decimal.TryParse(CStr(dr(0)), meta)
            End While
            Form1.cnctk.Close()
        Catch ex As Exception
            Form1.cnctk.Close()
            yn = False
            Return original
        End Try
    End Function
    Function s50(ByRef sql As String, ByRef topn As Integer, ByRef id As Integer, ByRef table As String) As Integer
        cnctm.Open()
        cmdstr = "select top 1 A.Id,A.SQL语句 from(select top(" & topn & ")Id,SQL语句 from 操作记录 where 记录Id=@id and 记录表='" & table & "' order by Id desc)as A order by A.Id"
        cmd = New SqlCommand(cmdstr, cnctm)
        cmd.Parameters.Add(New SqlParameter("id", id))
        dr = cmd.ExecuteReader
        While dr.Read
            sql = s19(CStr(dr(1)), table)
        End While
        dr.Close()
        dr = New SqlCommand("select*from " & table & " where 1=0", cnctm).ExecuteReader
        For i = 1 To dr.FieldCount - 1
            If InStr(13, sql, dr.GetName(i)) > 0 Then cnctm.Close() : Return i
        Next
        cnctm.Close()
        Return 0
    End Function
    Sub s51(T As TextBox, e As KeyEventArgs, ByRef bl As Boolean)
        If e.KeyCode = Keys.Enter Then
            Try
                Form1.cnctk.Open()
                cmdstr = "select " & T.Text
                cmd = New SqlCommand(cmdstr, Form1.cnctk)
                dr = cmd.ExecuteReader
                While dr.Read
                    T.Text = CStr(dr(0))
                End While
                Form1.cnctk.Close()
            Catch ex As Exception
                Form1.cnctk.Close()
                If bl Then T.Text = ""
            End Try
            T.SelectionStart = T.TextLength
        ElseIf e.KeyCode = Keys.Escape Then
            T.Text = CStr(T.Tag)
            T.SelectionStart = T.TextLength
        Else
            T.Tag = T.Text
        End If
    End Sub
    Sub s52(DA As DataGridView, ByRef a(,) As String, ByRef T As TextBox)
        a(0, 1) = "-2"
        Dim j As Integer
        Try
            For i = 0 To DA.SelectedCells.Count - 1
                If DA.SelectedCells.Count = 1 Then
                    If DA.SelectedCells(0).RowIndex = 0 Then
                        a(0, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex).Cells(1).Value)
                        a(0, 1) = CStr(DA.SelectedCells.Item(0).RowIndex)
                        a(1, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex + 1).Cells(1).Value)
                        a(1, 1) = CStr(DA.SelectedCells.Item(0).RowIndex + 1)
                    ElseIf DA.Rows.Count > 2 AndAlso DA.SelectedCells(0).RowIndex > DA.Rows.Count - 2 Then
                        a(0, 0) = CStr(DA.Rows(DA.Rows.Count - 3).Cells(1).Value)
                        a(0, 1) = CStr(DA.Rows.Count - 3)
                        a(1, 0) = CStr(DA.Rows(DA.Rows.Count - 2).Cells(1).Value)
                        a(1, 1) = CStr(DA.Rows.Count - 2)
                    Else
                        a(0, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex).Cells(1).Value)
                        a(0, 1) = CStr(DA.SelectedCells.Item(0).RowIndex)
                        a(1, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex - 1).Cells(1).Value)
                        a(1, 1) = CStr(DA.SelectedCells.Item(0).RowIndex - 1)
                    End If
                Else
                    If DA.SelectedCells.Item(i).RowIndex <> CInt(a(0, 1)) Then
                        If DA.SelectedCells(i).RowIndex = DA.Rows.Count - 1 Then
                            a(j, 0) = CStr(DA.Rows(DA.SelectedCells.Item(i).RowIndex - 1).Cells(1).Value)
                            a(j, 1) = CStr(DA.SelectedCells.Item(i).RowIndex - 1)
                            a(1 - j, 0) = CStr(DA.Rows(DA.SelectedCells.Item(i).RowIndex - 2).Cells(1).Value)
                            a(1 - j, 1) = CStr(DA.SelectedCells.Item(i).RowIndex - 2)
                            Exit For
                        Else
                            a(j, 0) = CStr(DA.Rows(DA.SelectedCells.Item(i).RowIndex).Cells(1).Value)
                            a(j, 1) = CStr(DA.SelectedCells.Item(i).RowIndex)
                        End If
                        j += 1
                        If j = 2 Then Exit For
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
        Dim bl As Boolean = CInt(a(0, 1)) > CInt(a(1, 1))
        T.Text = Fcsb.s43(a(-CInt(bl), 0), a(-CInt(Not bl), 0))
        If T.Text.Contains("-") Then
            T.BackColor = Color.FromArgb(255, 100, 100)
        ElseIf T.Text = "0分" Then
            T.BackColor = Color.FromArgb(255, 255, 192)
        Else
            T.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub
    Sub s53(dt As DataTable, cmd As SqlCommand)
        dt.Reset()
        Dim i = New SqlDataAdapter(cmd).Fill(dt)
        dt.Columns.Add("Id", Type.GetType("System.Int32"))
        For i = 0 To dt.Rows.Count - 1
            dt.Rows(i)(dt.Columns.Count - 1) = i
        Next
    End Sub
    Sub s54(li1 As ListBox, li2 As ListBox, k As List(Of String))
        If li2.Items.Count = 0 Then
            For i = 1 To li1.Items.Count
                k.Add(CStr(li1.Items(i - 1)))
            Next
            k.Add(" ")
        Else
            For i = 1 To li2.Items.Count
                k.Add(CStr(li2.Items(i - 1)))
            Next
        End If
    End Sub
    Function s55(ByRef d As String) As String
        Dim str As String
        Dim num As Byte = s10(d, Form1.CH32.Checked)
        If IsNothing(d) OrElse num = 0 Then
            str = ""
        Else
            Try
                Form1.cnct.Open()
                cmd = New SqlCommand("批号正则", Form1.cnct) With {.CommandType = CommandType.StoredProcedure}
                cmd.Parameters.Add(New SqlParameter("批号", d))
                cmd.Parameters.Add(New SqlParameter("代码", num))
                str = CStr(cmd.ExecuteScalar）
                Form1.cnct.Close()
            Catch ex As Exception
                Form1.cnct.Close()
            End Try
            Try
                Form1.cnct.Open()
                If Not CBool(New SqlCommand(String.Concat("select-1from 反应釜号 where 反应釜号 ='", Replace(str, "'", "''"), "'"), Form1.cnct).ExecuteScalar) Then str = ""
                Form1.cnct.Close()
            Catch ex As Exception
                Form1.cnct.Close()
            End Try
        End If
        Return str
    End Function
    Sub s56(DA As DataGridView)
        Form1.dacw(DA).Clear()
        Form1.dacw(DA).Add(DA.RowHeadersWidth)
        For i = 0 To DA.Columns.Count - 1
            Form1.dacw(DA).Add(DA.Columns(i).Width)
        Next
    End Sub
    Sub s57(DA As DataGridView)
        DA.RowHeadersWidth = Form1.dacw(DA)(0)
        For i = 0 To DA.Columns.Count - 1
            DA.Columns(i).Width = Form1.dacw(DA)(i + 1)
            DA.Columns(i).Visible = True
        Next
    End Sub
    Sub s58(DA As DataGridView)
        If DA.CurrentRow.Cells(0).Value IsNot Nothing AndAlso DA.CurrentRow.IsNewRow Then
            Dim row As Object
            Dim j As Integer = DA.CurrentCell.ColumnIndex
            DA.Rows.Add()
            For i = 0 To DA.Columns.Count - 1
                row = DA.Rows(DA.Rows.Count - 1).Cells(i).Value
                DA.Rows(DA.Rows.Count - 1).Cells(i).Value = DA.Rows(DA.Rows.Count - 2).Cells(i).Value
                DA.Rows(DA.Rows.Count - 2).Cells(i).Value = row
            Next
            row = DA.Rows(DA.Rows.Count - 1).Cells(0).Style
            DA.Rows(DA.Rows.Count - 1).Cells(0).Style = DA.Rows(DA.Rows.Count - 2).Cells(0).Style
            DA.Rows(DA.Rows.Count - 2).Cells(0).Style = DirectCast(row, DataGridViewCellStyle)
            If DA Is Form1.DA1 Then
                row = DA.Rows(DA.Rows.Count - 1).Cells(7).Style.ForeColor
                DA.Rows(DA.Rows.Count - 1).Cells(7).Style.ForeColor = DA.Rows(DA.Rows.Count - 2).Cells(7).Style.ForeColor
                DA.Rows(DA.Rows.Count - 2).Cells(7).Style.ForeColor = DirectCast(row, Color)
            End If
            DA.ClearSelection()
            DA.Rows(DA.Rows.Count - 2).Cells(j).Selected = True
        End If
    End Sub
    Sub s59(ByRef flag1 As Boolean, ByRef flag2 As Boolean)
        If flag1 Then
            If Form1.LI3.Items.Contains("入库") OrElse flag2 Then
                Form2.G7.Enabled = True
                Form2.DA1.Enabled = True
            Else
                Form2.G7.Enabled = False
                Form2.DA1.Enabled = False
            End If
        End If
    End Sub
    Sub s60()
        Form2.CB2.Items.Clear()
        DirectCast(Form2.DA1.Columns(1), DataGridViewComboBoxColumn).Items.Clear()
        For Each item In Form1.LI1.Items
            Form2.CB2.Items.Add(item)
            DirectCast(Form2.DA1.Columns(1), DataGridViewComboBoxColumn).Items.Add(item)
        Next
        For Each item In Form1.LI2.Items
            Form2.CB2.Items.Add(item)
            DirectCast(Form2.DA1.Columns(1), DataGridViewComboBoxColumn).Items.Add(item)
        Next
        Form2.CB2.Items.Add("")
        DirectCast(Form2.DA1.Columns(1), DataGridViewComboBoxColumn).Items.Add("")
    End Sub
End Module