Imports System.Data.SqlClient
Public Class Form2
    Public rg(,) As Object
    Dim xc As Object, n As Integer, bx As TextBox, at As Control, b(1, 1) As Object, otext, cmdstr As String, dgvcell As DataGridViewCell, pdt As DataTable = Form1.pdt, cnct As SqlConnection = Form1.cnct
    Public Sub Form2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form1.Show()
        Form1.WindowState = FormWindowState.Normal
        Form1.lct = Location
        Form1.dacl.Remove(DA1)
    End Sub
    Private Sub TA_KeyUp(sender As Object, e As KeyEventArgs)
        Dim i As Integer = s4(sender, 2)
        If e.KeyData = Keys.Enter Then
            s1(DirectCast(sender, TextBox), DirectCast(rg(3, i), Button), rg(10, i), DirectCast(rg(4, i), Button), DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), TC2.TabPages(CInt(rg(0, i))))
            LI1.Hide()
        ElseIf e.KeyData = Keys.Escape Then
            LI1.Hide()
        End If
    End Sub
    Public Sub T_TextChanged(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 2)
        If DirectCast(rg(2, i), TextBox).Text = "" Then
            s12(TC2.TabPages(CInt(rg(0, i))), DirectCast(sender, TextBox))
            DirectCast(rg(7, i), Dictionary(Of Control, String)).Clear()
            DirectCast(rg(3, i), Button).Enabled = False
            DirectCast(rg(4, i), Button).Enabled = False
            DirectCast(rg(5, i), Button).Text = ""
        Else
            s2(rg(8, i), DirectCast(sender, TextBox), DirectCast(rg(3, i), Button), DirectCast(rg(4, i), Button), rg(0, i))
            s3(DirectCast(sender, Control), CStr(rg(9, i)))
        End If
    End Sub
    Public Sub B_EnabledChanged(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 3)
        If DirectCast(rg(6, i), TextBox) IsNot Nothing Then
            If DirectCast(sender, Button).Enabled Then
                AddHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
            Else
                RemoveHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
            End If
        End If
    End Sub
    Public Sub TB_TextChanged(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 6)
        Dim j As Byte = CByte(rg(11, i))
        j = CByte(s4(j, 10))
        If j > -1 Then s3(DirectCast(sender, Control), CStr(rg(9, j)), -CInt(rg(0, i)))
    End Sub
    Public Sub TC_TextChanged(sender As Object, e As EventArgs)
        Dim tx As TextBox = DirectCast(sender, TextBox)
        If tx.Focused Then
            Dim i As Integer = s4(TC2.SelectedIndex, 0)
            Dim rg7 As Dictionary(Of Control, String) = DirectCast(rg(7, i), Dictionary(Of Control, String))
            RemoveHandler tx.TextChanged, AddressOf TC_TextChanged
            If Not DirectCast(rg(3, i), Button).Enabled Then
                MsgBox("请先输入批号初始化！")
                If tx.BackColor <> Color.White Then
                    If rg7.ContainsKey(tx) Then
                        tx.Text = rg7(tx)
                    Else
                        tx.Text = ""
                    End If
                End If
            ElseIf Form1.sbl(0) AndAlso tx.BorderStyle = BorderStyle.Fixed3D OrElse Form1.sbl(1) AndAlso (tx.ForeColor = Color.Red OrElse Not rg7.ContainsKey(tx) OrElse rg7(tx) = "") Then
                tx.BackColor = Color.White
            ElseIf rg7.ContainsKey(tx) Then
                tx.Text = rg7(tx)
            Else
                tx.Text = ""
            End If
            If tx.BackColor <> Color.White Then tx.SelectionStart = tx.TextLength
            AddHandler tx.TextChanged, AddressOf TC_TextChanged
        End If
    End Sub
    Public Sub T_LostFocus(sender As Object, e As EventArgs)
        Dim TXBX As TextBox = DirectCast(sender, TextBox)
        If TXBX.BorderStyle = BorderStyle.Fixed3D Then
            Dim TC As Control = TXBX.Parent
            Do
                If TypeOf TC Is TabPage Then
                    Exit Do
                Else
                    TC = TC.Parent
                End If
            Loop
            Try
                cnct.Open()
                If CStr(New SqlCommand("select b.name from syscolumns a,systypes b,sysobjects d where a.xtype=b.xusertype and a.id=d.id and d.name='" & Replace(CStr(TC.Tag), "'", "''") & "工艺' and a.name='" & Replace(Replace(Replace(CStr(TXBX.Tag), "[", ""), "]", ""), "'", "''") & "'", cnct).ExecuteScalar).Contains("date") AndAlso DirectCast(rg(3, s4(CStr(TC.Tag), 9)), Button).Enabled Then TXBX.Text = s48(TXBX.Text)
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        End If
        b(0, -CInt(b(1, 1))) = TXBX.Text
        Dim str As String = s43(CStr(b(0, -CInt(b(1, 0)))), CStr(b(0, -CInt(b(1, 1)))))
        If str.Contains("-") Then
            Text = otext & "████" & str & "████"
        Else
            Text = otext & str
        End If
        b(1, 0) = Not CBool(b(1, 0))
        b(1, 1) = Not CBool(b(1, 1))
    End Sub
    Private Sub BA_Click(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 1), TB As TextBox, str As String, bl() As Object = ｛rg(10, i), rg(11, i)｝
        For l = 0 To 1
            TB = DirectCast(rg(2, s4(sender, 1)), TextBox)
            If Not IsDBNull(bl(l)) Then
                i = s4(bl(l), 11 - l)
                Do Until i = -1
                    If CBool(l) Then
                        str = CStr(rg(9, s4(TB, 2)))
                        cmdstr = "select 上一工序批号 from [" + str + "工艺] where [" + str + "批号]=@批号"
                    Else
                        str = CStr(rg(9, i))
                        cmdstr = "select [" + str + "批号] from [" + str + "工艺] where 上一工序批号=@批号"
                    End If
                    cmd = New SqlCommand(cmdstr, cnct)
                    cmd.Parameters.AddWithValue("批号", TB.Text)
                    RemoveHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf T_TextChanged
                    Try
                        cnct.Open()
                        DirectCast(rg(2, i), TextBox).Text = CStr(IIf(IsDBNull(cmd.ExecuteScalar), "", cmd.ExecuteScalar))
                        cnct.Close()
                    Catch ex As Exception
                        cnct.Close()
                    End Try
                    s2(rg(8, i), DirectCast(rg(2, i), TextBox), DirectCast(rg(3, i), Button), DirectCast(rg(4, i), Button), rg(0, i))
                    If Not DirectCast(rg(3, i), Button).Enabled AndAlso Form1.CL2.Items.Contains(CStr(rg(9, i))) Then s1(DirectCast(rg(2, i), TextBox), DirectCast(rg(3, i), Button), rg(10, i), DirectCast(rg(4, i), Button), DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), TC2.TabPages(CInt(rg(0, i))))
                    AddHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf T_TextChanged
                    TB = DirectCast(rg(2, i), TextBox)
                    If IsDBNull(rg(10 + l, i)) Then
                        Exit Do
                    Else
                        i = s4(CInt(rg(10 + l, i)), 11 - l)
                    End If
                Loop
            End If
        Next
        LI1.Hide()
    End Sub
    Private Sub BB_Click(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 5)
        Dim bln As Boolean
        Dim Bm As Button = DirectCast(sender, Button)
        Dim Bn As Button = DirectCast(rg(3, i), Button)
        Dim Bo As Button = DirectCast(rg(4, i), Button)
        Dim Tn As TextBox = DirectCast(rg(2, i), TextBox)
        Dim tb As Control = TC2.TabPages(CInt(rg(0, i)))
        Dim rg8 As String = CStr(rg(8, i))
        Dim rg9 As String = CStr(rg(9, i))
        Dim n As Object = rg(10, i)
        Dim bl As Boolean = True
        Dim dtr() As DataRow
        If Bm.Text > "" Then
            If Bn.Enabled Then
                s25(Form1.DA11, CInt(Bm.Text), rg9 & "工艺")
                Form1.TC1.SelectedIndex = 4
                Form1.Show()
                Form1.WindowState = FormWindowState.Normal
                Form1.Activate()
            ElseIf s7(tb) Then
                If Tn.Text > "" Then
                    s34(Tn, rg9, n, bl)
                    If Fcsb.s10(Tn.Text, True) = CByte(n) OrElse bl Then
                        Try
                            cnct.Open()
                            Dim str(1) As String
                            dr = New SqlCommand("select Id,[" & rg9 & "批号]from[" & rg9 & "工艺]With(tablockx)where[" & rg9 & "批号]='" & Replace(rg8, "'", "''") & "'union all select Id,[" & rg9 & "批号]from[" & rg9 & "工艺]where Id=" & Bm.Text, cnct).ExecuteReader
                            While dr.Read
                                str(0) = CStr(dr(0))
                                str(1) = CStr(dr(1))
                            End While
                            dr.Close()
                            If IsNothing(str(0)) Then
                                Bn.Enabled = True
                                Bo.Enabled = True
                                LI1.Hide()
                                RemoveHandler Tn.TextChanged, AddressOf T_TextChanged
                                Tn.Text = rg8
                                AddHandler Tn.TextChanged, AddressOf T_TextChanged
                                MsgBox("该" & rg9 & "批次已删除，更新已终止！")
                                Return
                            Else
                                Bm.Text = str(0)
                            End If
                            If Tn.Text = str(1) Then
                                str(0) = "begin tran commit tran"
                            Else
                                str(0) = "update [" & rg9 & "工艺] set [" & rg9 & "批号]='" & Replace(Tn.Text, "'", "''") & "' where Id=" & Bm.Text
                            End If
                            Dim j = New SqlCommand(str(0), cnct).ExecuteNonQuery()
                            s8(DirectCast(rg(7, i), Dictionary(Of Control, String)), tb, Tn)
                            dtr = pdt.Select("BN ='" & Replace(rg8, "'", "''") & "'")
                            For h = 1 To dtr.Count
                                dtr(h - 1)(0) = Tn.Text
                            Next
                            Bn.Enabled = True
                            Bo.Enabled = True
                            rg(8, i) = Tn.Text
                            s14(TC2.SelectedTab, Tn, Bm, rg9, DirectCast(rg(7, i), Dictionary(Of Control, String)))
                            MsgBox("批号目前为: " & Tn.Text & " ,详细信息请点击Id按钮查询")
                        Catch ex As Exception
                            cnct.Close()
                            MsgBox(ex.Message)
                        End Try
                    End If
                Else
                    bln = True
                End If
            ElseIf Tn.Text = "" Then
                bln = True
            Else
                MsgBox("更改批次失败，可能的原因是权限不足！")
            End If
        Else
            bln = rg8 > ""
        End If
        If bln Then
            RemoveHandler Tn.TextChanged, AddressOf T_TextChanged
            Tn.Text = rg8
            AddHandler Tn.TextChanged, AddressOf T_TextChanged
            Bn.Enabled = True
            Bo.Enabled = True
            s1(DirectCast(rg(2, i), TextBox), Bn, rg(10, i), Bo, DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), TC2.TabPages(CInt(rg(0, i))))
        End If
        LI1.Hide()
    End Sub
    Private Sub BC_Click(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 3)
        s1(DirectCast(rg(2, i), TextBox), DirectCast(sender, Button), rg(10, i), DirectCast(rg(4, i), Button), DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), TC2.TabPages(CInt(rg(0, i))))
        Dim Tn As TextBox = DirectCast(rg(2, i), TextBox)
        If s7(TC2.TabPages(CInt(rg(0, i)))) Then
            Dim ct As Control = TC2.TabPages(CInt(rg(0, i)))
            Dim Bn As Button = DirectCast(rg(5, i), Button)
            Dim str As String = CStr(rg(9, i))
            cnct.Open()
            cmdstr = "begin tran select-1from[" & str & "工艺]with(tablockx)where[" & str & "批号]='" & Replace(Tn.Text, "'", "''") & "'"
            If CBool(New SqlCommand(cmdstr, cnct).ExecuteScalar) Then cmdstr = "delete from [" & str & "工艺] where [" & str & "批号]='" & Replace(Tn.Text, "'", "''") & "'"
            cnct.Close()
            If Strings.Left(cmdstr, 1) = "d" Then
                Try
                    cmd = New SqlCommand(cmdstr, cnct)
                    cnct.Open()
                    cmd.ExecuteNonQuery()
                    cnct.Close()
                    Dim dtr() As DataRow = pdt.Select("BN='" & Replace(Tn.Text, "'", "''") & "'")
                    For Each row As DataRow In dtr
                        pdt.Rows.Remove(row)
                    Next
                    MsgBox("批次删除已成功！")
                    s12(ct, Tn)
                    Bn.Text = ""
                Catch ex As Exception
                    cnct.Close()
                    MsgBox(ex.Message)
                End Try
            Else
                Try
                    cnct.Open()
                    Dim j As Integer = New SqlCommand("begin tran commit tran", cnct).ExecuteNonQuery()
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                End Try
                If Bn.Text > "" Then
                    DirectCast(rg(4, i), Button).Enabled = False
                    DirectCast(rg(3, i), Button).Enabled = False
                    s12(ct, Tn)
                    Bn.Text = ""
                End If
                MsgBox("无法对 批次 " & Tn.Text & " 执行 删除，因为它不存在，或者您没有所需的权限。")
            End If
        Else
            MsgBox("无法对 批次 " & Tn.Text & " 执行 删除，因为它不存在，或者您没有所需的权限。")
        End If
        LI1.Hide()
    End Sub
    Private Sub BD_Click(sender As Object, e As EventArgs)
        Dim i As Integer = s4(sender, 4)
        If DirectCast(rg(2, i), TextBox).Text > "" Then
            Dim dgvc As DataGridViewComboBoxColumn = DirectCast(Form1.DA1.Columns(6), DataGridViewComboBoxColumn)
            Form1.T39.Text = ""
            Form1.T40.Text = ""
            Form1.T52.Text = ""
            Form1.T53.Text = ""
            RemoveHandler Form1.T38.TextChanged, AddressOf Form1.T38_TextChanged
            Form1.T38.Text = DirectCast(rg(2, i), TextBox).Text
            AddHandler Form1.T38.TextChanged, AddressOf Form1.T38_TextChanged
            RemoveHandler Form1.D2.ValueChanged, AddressOf Form1.D_Change
            Form1.D2.Checked = True
            Form1.D2.Value = DateAdd(DateInterval.Day, 1, Now)
            AddHandler Form1.D2.ValueChanged, AddressOf Form1.D_Change
            Form1.LI3.Items.Clear()
            Form1.LI4.Items.Clear()
            Try
                cnct.Open()
                Fcsb.s6("select 物料类型 from 物料类型 where 可用性=1 or 物料类型 in('消耗','产出','回收') order by Id", dgvc)
                Fcsb.s3(Form1.LI3, "select 物料类型 from 物料类型 where 可用性=1 or 物料类型 in('消耗','产出','回收') order by Id")
                s29()
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
            Form1.TC1.SelectedIndex = 0
            If Form1.CL2.Items.IndexOf(rg(9, i)) > 0 Then Form1.CL2.SetItemChecked(Form1.CL2.Items.IndexOf(rg(9, i)), True)
            Form1.CL2_MouseUp(Form1.CL2, e)
            Form1.T1.Text = ""
            Form1.CO9.Text = ""
            Form1.D1.Checked = False
            blph = True
            s39(False)
            blph = False
            Form1.Show()
            Form1.WindowState = FormWindowState.Normal
            Form1.Activate()
        End If
    End Sub
    Private Sub LI1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles LI1.MouseDoubleClick
        If LI1.SelectedItems.Count = 0 Then Return
        Dim i As Integer = s4(Math.Abs(n), 0)
        If n > 0 Then
            s2(rg(8, i), DirectCast(rg(2, i), TextBox), DirectCast(rg(3, i), Button), DirectCast(rg(4, i), Button))
            If Not DirectCast(rg(3, i), Button).Enabled Then s1(DirectCast(rg(2, i), TextBox), DirectCast(rg(3, i), Button), rg(10, i), DirectCast(rg(4, i), Button), DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), TC2.TabPages(CInt(rg(0, i))))
        Else
            LI1.Hide()
            Dim Tn As TextBox = DirectCast(rg(6, i), TextBox)
            RemoveHandler Tn.TextChanged, AddressOf TB_TextChanged
            Tn.Text = CStr(LI1.SelectedItem)
            AddHandler Tn.TextChanged, AddressOf TB_TextChanged
            Tn.Focus()
            Tn.SelectionStart = Tn.TextLength
        End If
    End Sub
    Private Sub TC2_KeyDown(sender As Object, e As KeyEventArgs) Handles TC2.KeyDown
        Dim i As Integer = s4(DirectCast(sender, TabControl).SelectedIndex, 0)
        If i = -1 OrElse DirectCast(rg(2, i), TextBox).Focused Then Return
        If TC2.SelectedIndex > 0 AndAlso i > -1 Then
            If DirectCast(rg(3, i), Button).Enabled = True Then
                If (e.KeyData = 13 OrElse e.KeyData = 262162) AndAlso TypeOf ActiveControl Is TextBox AndAlso DirectCast(ActiveControl, TextBox).Multiline = False OrElse e.KeyData = 262162 AndAlso TypeOf ActiveControl Is TextBox AndAlso DirectCast(ActiveControl, TextBox).Multiline = True Then
                    If TypeOf ActiveControl Is TextBox Then
                        Dim TB As TextBox = DirectCast(ActiveControl, TextBox)
                        DirectCast(rg(2, i), TextBox).Focus()
                        DirectCast(rg(2, i), TextBox).SelectionLength = 0
                        TB.Focus()
                        TB.SelectionLength = 0
                        TB.SelectionStart = TB.TextLength
                    End If
                    If IsDBNull(rg(11, i)) OrElse IsNothing(rg(6, i)) Then
                        s9(DirectCast(rg(2, i), TextBox), CStr(rg(9, i)), TC2.TabPages(CInt(rg(0, i))), DirectCast(rg(7, i), Dictionary(Of Control, String)), DirectCast(rg(5, i), Button))
                    Else
                        Dim bl As Boolean = True
                        Dim msb As MsgBoxResult
                        Dim str As String
                        Dim CH32 As Boolean = Form1.CH32.Checked
                        Dim CH33 As Boolean = Form1.CH33.Checked
                        RemoveHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
                        Do
                            str = "上一工序批号格式不"
                            If DirectCast(rg(6, i), TextBox).Text > "" Then
                                If CH32 AndAlso Fcsb.s10(DirectCast(rg(6, i), TextBox).Text, True) = 0 Then
                                    str += "正确！"
                                ElseIf CH33 AndAlso (CH32 AndAlso Fcsb.s10(DirectCast(rg(6, i), TextBox).Text, True) <> CByte(rg(11, i)) AndAlso Fcsb.s10(DirectCast(rg(6, i), TextBox).Text, True) > 0 OrElse Not CH32 AndAlso Fcsb.s10(DirectCast(rg(6, i), TextBox).Text, False) <> CInt(rg(11, i)) AndAlso Fcsb.s10(DirectCast(rg(6, i), TextBox).Text, False) > 0) Then
                                    str += "匹配！"
                                Else
                                    Exit Do
                                End If
                                msb = MsgBox(str, DirectCast(MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.DefaultButton1, MsgBoxStyle))
                                If msb = MsgBoxResult.Abort Then
                                    DirectCast(rg(6, i), TextBox).Focus()
                                    bl = False
                                    Exit Do
                                ElseIf msb = MsgBoxResult.Ignore Then
                                    Exit Do
                                End If
                            Else
                                Exit Do
                            End If
                        Loop
                        If Fcsb.s10(DirectCast(rg(6, i), TextBox).Text, CH32) = CByte(rg(11, i)) OrElse DirectCast(rg(6, i), TextBox).Text = "" OrElse bl Then
                            s9(DirectCast(rg(2, i), TextBox), CStr(rg(9, i)), TC2.TabPages(CInt(rg(0, i))), DirectCast(rg(7, i), Dictionary(Of Control, String)), DirectCast(rg(5, i), Button))
                        End If
                        AddHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
                    End If
                    LI1.Hide()
                Else
                    If TypeOf ActiveControl Is TextBox Then
                        Dim dit As Dictionary(Of Control, String) = DirectCast(rg(7, i), Dictionary(Of Control, String))
                        If ActiveControl.BackColor = Color.White Then
                            If e.KeyCode = Keys.Escape Then
                                If LI1.Visible Then
                                    LI1.Items.Clear()
                                    LI1.Hide()
                                Else
                                    If rg(6, i) IsNot Nothing Then RemoveHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
                                    ActiveControl.Text = ""
                                    If dit.ContainsKey(ActiveControl) Then ActiveControl.Text = dit(ActiveControl)
                                    If rg(6, i) IsNot Nothing Then AddHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
                                    ActiveControl.BackColor = Color.FromArgb(165, 215, 175)
                                    DirectCast(ActiveControl, TextBox).SelectionStart = DirectCast(ActiveControl, TextBox).TextLength
                                End If
                            End If
                        ElseIf Not Form1.sbl(1) Then
                            If Form1.sbl(0) AndAlso e.KeyCode = Keys.Escape Then
                                If LI1.Visible Then
                                    LI1.Items.Clear()
                                    LI1.Hide()
                                ElseIf DirectCast(ActiveControl, TextBox).BorderStyle = BorderStyle.Fixed3D Then
                                    ActiveControl.BackColor = Color.White
                                End If
                            End If
                        ElseIf (ActiveControl.Text = "" OrElse ActiveControl.ForeColor = Color.Red) AndAlso e.KeyCode = Keys.Escape Then
                            If LI1.Visible Then
                                LI1.Items.Clear()
                                LI1.Hide()
                            ElseIf DirectCast(ActiveControl, TextBox).BorderStyle = BorderStyle.Fixed3D Then
                                ActiveControl.BackColor = Color.White
                            End If
                        End If
                    End If
                End If
            ElseIf e.KeyCode = Keys.Escape OrElse e.KeyCode = Keys.Enter Then
                MsgBox("请先输入批号初始化！")
                DirectCast(ActiveControl, TextBox).SelectionStart = DirectCast(ActiveControl, TextBox).TextLength
            End If
        End If
    End Sub
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form1.flct Then
            Location = Form1.lct
        Else
            Form1.flct = True
        End If
        otext = Text & "—" & Form1.usr & "        时间间隔:"
        Text = otext
        b(1, 0) = False
        b(1, 1) = True
        If Form1.sbl(3) Then B79.Enabled = False
        If Form1.suer = 4 Then
            DA1.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            L126.Enabled = False
        End If
        CB1.Items.Add(Form1.st(0))
        Try
            cnctm.Open()
            cmd = New SqlCommand("sp_helpuser", cnctm)
            dr = cmd.ExecuteReader
            While dr.Read
                If CInt(dr(5)) > 4 Then
                    CB1.Items.Add(dr(0))
                End If
            End While
            CB1.Items.Add("")
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
        End Try
        s60()
        cmdstr = "select 班别班组,id from 班别班组 order by id"
        Try
            cnct.Open()
            Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(4), DataGridViewComboBoxColumn))
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        DA1.Rows(0).Cells(4).Value = Form1.usr
        ReDim Preserve rg(11, 0)
        For k = 1 To TC2.TabPages.Count - 1
            rg(0, UBound(rg, 2)) = k
            For Each ct As Control In TC2.TabPages(k).Controls
                If ct.Text = "批号：" Then
                    rg(1, UBound(rg, 2)) = ct
                    AddHandler ct.Click, AddressOf BA_Click
                ElseIf TypeOf ct Is TextBox Then
                    If CStr(ct.Tag) = "" Then
                        rg(2, UBound(rg, 2)) = ct
                        AddHandler ct.TextChanged, AddressOf T_TextChanged
                        AddHandler ct.KeyUp, AddressOf TA_KeyUp
                        AddHandler ct.MouseDown, AddressOf T_MouseDown
                        If Form1.suer = 4 Then DirectCast(ct, TextBox).Enabled = False
                    ElseIf CStr(ct.Tag) = "上一工序批号" Then
                        rg(6, UBound(rg, 2)) = ct
                    End If
                ElseIf ct.Text = "删除内容" Then
                    rg(3, UBound(rg, 2)) = ct
                    AddHandler ct.EnabledChanged, AddressOf B_EnabledChanged
                    AddHandler ct.Click, AddressOf BC_Click
                ElseIf ct.Text = "链接物料" Then
                    rg(4, UBound(rg, 2)) = ct
                    AddHandler ct.Click, AddressOf BD_Click
                ElseIf TypeOf ct Is Button AndAlso ct.Text = "" Then
                    rg(5, UBound(rg, 2)) = ct
                    AddHandler ct.Click, AddressOf BB_Click
                End If
            Next
            rg(7, UBound(rg, 2)) = New Dictionary(Of Control, String)
            rg(9, UBound(rg, 2)) = TC2.TabPages(k).Tag
            Dim i As Integer = s4(k, 0)
            Try
                cnct.Open()
                cmd = New SqlCommand("select distinct 批号代码.批号代码,前道代码 from 批号代码,工序类型 where 批号代码.批号代码=工序类型.批号代码 and 操作工序='" & Replace(CStr(rg(9, i)), "'", "''") & "' and 批号代码.批号代码 is not NULL", cnct)
                dr = cmd.ExecuteReader
                While dr.Read
                    rg(10, i) = dr(0)
                    rg(11, i) = dr(1)
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
            ReDim Preserve rg(11, UBound(rg, 2) + 1)
            s41(TC2.TabPages(k))
        Next
        ReDim Preserve rg(11, UBound(rg, 2) - 1)
        If Form1.suer = 4 Then
            L126.Enabled = False
            CB1.Enabled = False
            T32.Enabled = False
        End If
        Form1.dacl.Add(DA1, New List(Of DataGridViewCell))
        AddHandler DA1.MouseWheel, AddressOf Form1.MouseWheel
        AddHandler DA1.RowPostPaint, AddressOf Form1.RowPostPaint
        AddHandler DA1.CellMouseEnter, AddressOf Form1.CellMouseEnter
        AddHandler DA1.CellMouseLeave, AddressOf Form1.CellMouseLeave
        Form1.dacw.Add(DA1, New List(Of Integer)) : s56(DA1)
        If Screen.PrimaryScreen.Bounds.Width <= Width OrElse Screen.PrimaryScreen.Bounds.Height <= Height Then MsgBox("屏幕分辨率不得小于974×665！")
    End Sub
    Private Sub TC2_MouseUp(sender As Object, e As MouseEventArgs) Handles TC2.MouseUp
        LI1.Hide()
        LI1.Items.Clear()
    End Sub
    Public Sub B78_Click(sender As Object, e As EventArgs) Handles B78.Click
        Dim blct As Boolean
        If CB2.Items.Contains(CB2.Text) Then
            If Not Form1.ccbl2 Then
                Form1.ni = 0
                Form1.TM1.Interval = SystemInformation.DoubleClickTime
                Form1.TM1.Enabled = True
                Form1.ccbl2 = True
                Return
            Else
                Form1.TM1.Interval = 1
                Form1.TM1.Enabled = False
                blct = True
            End If
            s42(blct)
        End If
    End Sub
    Private Sub B79_Click(sender As Object, e As EventArgs) Handles B79.Click
        Fcsb.s7(DirectCast(sender, Button), B17, DA1, Form1.idt3)
        If Not Form1.sbl(0) Then DA1.Columns(4).ReadOnly = True
    End Sub
    Private Sub B17_Click(sender As Object, e As EventArgs) Handles B17.Click
        Fcsb.s4(DA1, "原料入库", Form1.idt3)
    End Sub
    Public Sub L126_Click(sender As Object, e As EventArgs) Handles L126.Click
        If Not Form1.lbl126 Then
            Form1.ni = 0
            Form1.TM1.Interval = SystemInformation.DoubleClickTime
            Form1.TM1.Enabled = True
            Form1.lbl126 = True
            Return
        Else
            Form1.TM1.Interval = 1
            Form1.TM1.Enabled = False
        End If
        s30(True, sender, False)
    End Sub
    Private Sub DA1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.SelectedCells.Count = 0 OrElse DA.SelectedRows.Count > 0 Then Return
        If DA.SelectedCells.Count = 1 AndAlso e.RowIndex > -1 Then
            DA.Columns(DA.SelectedCells(0).ColumnIndex).Visible = True
            If e.Button = Windows.Forms.MouseButtons.Left Then DA.BeginEdit(True)
        End If
    End Sub
    Public Sub DA1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA1.CellBeginEdit
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        xc = Nothing
        G7.Enabled = False
        For i = 1 To DA.Columns.Count
            DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        If CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then Return
        Try
            cnct.Open()
            dr = New SqlCommand("select * from 原料入库 where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value), cnct).ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    xc = IIf(IsDBNull(dr(e.ColumnIndex)), Nothing, dr(e.ColumnIndex))
                    For Each col As DataGridViewColumn In DA.Columns
                        DA.Rows(e.RowIndex).Cells(col.Index).Value = IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))
                    Next
                End While
            Else
                e.Cancel = True
                DA.Rows(e.RowIndex).Cells(0).Value = 0
                DA.Rows(e.RowIndex).ReadOnly = True
            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Public Sub DA1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA1.CellEndEdit
        Dim cmdstrb As String, DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.Rows.Count = 1 OrElse CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) > "" Then
            G7.Enabled = True
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        cmdstr = ""
        Form1.skip(1) = False : Form1.skip(0) = False
        If DA.Rows(e.RowIndex).Cells(0).Value IsNot Nothing AndAlso CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CStr(xc) Then
            cmdstr = "update 原料入库 set "
            cmdstrb = DA.Columns(e.ColumnIndex).HeaderText
            cmdstr += cmdstrb & "=@" & cmdstrb & " where Id=@Id"
        End If
        If cmdstr > "" Then
            Try
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.Parameters.Add(New SqlParameter("Id", DA.Rows(e.RowIndex).Cells(0).Value))
                cmd.Parameters.Add(New SqlParameter(cmdstrb, DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value))
                cnct.Open()
                cmd.ExecuteNonQuery()
                cnct.Close()
                If CInt(DA.Rows(e.RowIndex).Cells(0).Value) <> 0 Then
                    If DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(200, 200, 200) Then
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(175, 175, 175)
                    ElseIf DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(175, 175, 175) Then
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(150, 150, 150)
                    Else
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(200, 200, 200)
                    End If
                End If
                If Not Form1.idt3.Contains(CInt(DA.Rows(e.RowIndex).Cells(0).Value)) Then Form1.idt3.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
            Catch ex As Exception
                cnct.Close()
                Form1.skip(0) = True
                s10(e)
                MsgBox("原料入库信息更改失败，请重试！" & vbCrLf & ex.Message)
                Return
            End Try
        End If
    End Sub
    Public Sub DA1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex = -1 Then
            If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                If e.Button = MouseButtons.Middle Then
                    s57(DA)
                ElseIf e.Button = MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                    DA.Columns.Item(e.ColumnIndex).Visible = False
                End If
            End If
        ElseIf CStr(DA.Rows(e.RowIndex).Cells(0).Value) > "" Then
            If e.ColumnIndex = 0 Then
                If DA.Rows.Count > 1 Then
                    If e.Button = Windows.Forms.MouseButtons.Middle AndAlso CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) > "" AndAlso Not Form1.sbl(3) Then
                        DA.EndEdit()
                        If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                            Dim en As EventArgs
                            If B79.Text = "解锁表格" Then B79_Click(B79, en)
                            DA.Rows.Add()
                            For i = 1 To 3
                                DA.Rows(DA.Rows.Count - 2).Cells(i).Value = DA.Rows(e.RowIndex).Cells(i).Value
                            Next
                            DA.Rows(DA.Rows.Count - 2).Cells(4).Value = Form1.usr
                            RemoveHandler DA.RowValidating, AddressOf DA1_RowValidating
                            RemoveHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                            DA.Columns(1).Visible = True
                            DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(2)
                            DA.Rows(DA.Rows.Count - 2).ReadOnly = False
                            DA.BeginEdit(True)
                            DA.Rows(e.RowIndex).Cells(0).Selected = True
                            AddHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                            AddHandler DA.RowValidating, AddressOf DA1_RowValidating
                            Form1.dttm = CStr(DA.Rows(e.RowIndex).Cells(1).Value)
                        End If
                    ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
                        DA.ClearSelection()
                        DA.Rows(e.RowIndex).Cells(0).Selected = True
                        Form1.B105_Click(Form1.B105, e)
                        Form1.B105_Click(Form1.B105, e)
                        If Form1.LI1.Items.Contains(DA.Rows(e.RowIndex).Cells(1).Value) Then
                            Form1.LI1.Items.Remove(DA.Rows(e.RowIndex).Cells(1).Value)
                            Form1.LI2.Items.Add(DA.Rows(e.RowIndex).Cells(1).Value)
                        End If
                        Form1.LI3.Items.Remove("入库")
                        Form1.LI4.Items.Add("入库")
                        RemoveHandler Form1.T38.TextChanged, AddressOf Form1.T38_TextChanged
                        Form1.T38.Text = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
                        AddHandler Form1.T38.TextChanged, AddressOf Form1.T38_TextChanged
                        Form1.DA1.Rows.Clear()
                        Form1.B14_Click(Form1.B14, e)
                        Form1.B14_Click(Form1.B14, e)
                        Form1.Show()
                        Form1.WindowState = FormWindowState.Normal
                        Form1.Activate()
                    End If
                End If
            End If
        End If
    End Sub
    Public Sub DA1_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA1.RowValidating
        Dim an As Date, DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            DA.EndEdit()
            cmdstr = "insert into 原料入库 values("
            For i = 1 To DA.Columns.Count - 1
                cmdstr += "@" + DA.Columns(i).HeaderText + ","
            Next
            cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) + ")"
            For i = 1 To 4
                If CStr(DA.Rows(e.RowIndex).Cells(i).Value) = "" Then
                    s16(DA, i, e)
                    Return
                End If
            Next
            Dim sqlpamt(3) As SqlParameter
            For i = 1 To DA.Columns.Count - 1
                sqlpamt(i - 1) = New SqlParameter(DA.Columns(i).HeaderText, CStr(DA.Rows(e.RowIndex).Cells(i).Value))
            Next
            If Not Fcsb.s9(DA.Rows.Count - 2, DA, "原料入库", cmdstr, sqlpamt) Then
                Form1.skip(0) = False
                If Form1.suer = 4 Then
                    DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Pink
                ElseIf Form1.sbl(0) Then
                    DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Pink
                ElseIf Form1.suer = 2 Then
                    DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Pink
                End If
                Form1.idt3.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
                G7.Enabled = True
                For i = 1 To DA.Columns.Count
                    DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
                Next
            End If
        End If
    End Sub
    Private Sub DA1_KeyDown(sender As Object, e As KeyEventArgs) Handles DA1.KeyDown
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.KeyCode = Keys.Escape Then
            Try
                If IsNothing(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) Then
                    DA.Rows.RemoveAt(DA.Rows.Count - 2)
                    G7.Enabled = True
                    For Each col As DataGridViewColumn In DirectCast(sender, DataGridView).Columns
                        col.SortMode = DataGridViewColumnSortMode.Automatic
                    Next
                    Form1.skip(0) = False
                End If
            Catch ex As Exception
            End Try
        Else
            Form1.dabl = e.KeyCode = Keys.ShiftKey
        End If
    End Sub
    Private Sub DA1_KeyUp(sender As Object, e As KeyEventArgs) Handles DA1.KeyUp
        Form1.dabl = False
    End Sub
    Public Sub DA1_SelectionChanged(sender As Object, e As EventArgs) Handles DA1.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If Form1.skip(1) Then
            RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
            If DA.CurrentCell IsNot dgvcell Then
                RemoveHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                RemoveHandler DA.RowValidating, AddressOf DA1_RowValidating
                RemoveHandler DA.CellBeginEdit, AddressOf DA1_CellBeginEdit
                DA.Columns(dgvcell.ColumnIndex).Visible = True
                DA.CurrentCell = dgvcell
                DA.BeginEdit(False)
                AddHandler DA.CellBeginEdit, AddressOf DA1_CellBeginEdit
                AddHandler DA.RowValidating, AddressOf DA1_RowValidating
                AddHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
            End If
        End If
    End Sub
    Protected Overloads Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Dim DA As DataGridView, bl As Boolean
        If TypeOf ActiveControl Is DataGridViewTextBoxEditingControl Then
            DA = DirectCast(ActiveControl, DataGridViewTextBoxEditingControl).EditingControlDataGridView
        ElseIf TypeOf ActiveControl Is DataGridViewComboBoxEditingControl Then
            DA = DirectCast(ActiveControl, DataGridViewComboBoxEditingControl).EditingControlDataGridView
        Else
            Exit Function
        End If
        If keyData = Keys.Escape Then
            RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
            If Form1.skip(1) Then
                dgvcell.Value = xc
                DA.CancelEdit() : DA.EndEdit()
                Form1.skip(1) = False
                bl = DA.Rows(dgvcell.RowIndex).Cells(0).Value IsNot Nothing
            Else
                DA.CancelEdit() : DA.EndEdit()
                bl = DA.Rows.Count = 1 OrElse IsNothing(DA.Rows(DA.Rows.Count - 1).Cells(0).Value) AndAlso DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing
            End If
            If bl OrElse DA.CurrentRow.Cells(0).Value IsNot Nothing AndAlso DA.CurrentRow.IsNewRow Then
                G7.Enabled = True
                For i = 1 To DA.Columns.Count
                    DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
                Next
                Form1.skip(0) = False
            End If
            s58(DA)
            AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
        Else
            Dim flag As Boolean, a As Integer
            Do
                For c = 0 To 1
                    For d = 0 To 1
                        flag = keyData = 9 + 65536 * a + 131072 * c + 262144 * d OrElse keyData = 13 + 65536 * a + 131072 * c + 262144 * d
                        If flag Then Exit Do
                        For i = 33 To 40
                            flag = keyData = i + 65536 * a + 131072 * c + 262144 * d
                            If flag Then Exit Do
                        Next
                    Next
                Next
                a += 1
            Loop Until a = 2
            If flag AndAlso Not (TypeOf ActiveControl Is DataGridViewComboBoxEditingControl AndAlso Math.Abs(keyData - 39) = 1) Then
                If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                    DA.EndEdit()
                    Return True
                End If
            End If
        End If
    End Function
    Private Sub B80_Click(sender As Object, e As EventArgs) Handles B80.Click
        Fcsb.s8(DA1, True)
    End Sub
    Public Sub TC2_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TC2.Selecting
        If Form1.skip(0) Then
            e.Cancel = True
            DA1.BeginEdit(False)
            at = ActiveControl
            AddHandler DA1.GotFocus, AddressOf DA1_GotFocus
        Else
            For Each item In Form1.CL2.Items
                If CStr(e.TabPage.Tag) = CStr(item) Then Return
            Next
            e.Cancel = e.TabPageIndex > 0
        End If
    End Sub
    Private Sub DA1_GotFocus(sender As Object, e As EventArgs)
        RemoveHandler DA1.GotFocus, AddressOf DA1_GotFocus
        ActiveControl = at
    End Sub
    Private Sub DA1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DA1.RowsAdded
        DA1.Rows(e.RowIndex).Cells(4).Value = Form1.usr
    End Sub
    Private Sub T_MouseDown(sender As Object, e As EventArgs)
        Dim T As TextBox = DirectCast(sender, TextBox)
        If T.Focused Then
            T.ContextMenuStrip = Nothing
        ElseIf Form1.suer <> 4 Then
            T.ContextMenuStrip = Form1.CMS1
        End If
    End Sub
    Private Sub TC2_MouseWheel(sender As Object, e As MouseEventArgs) Handles TC2.MouseWheel
        If DirectCast(sender, TabControl).SelectedIndex = 0 Then
            If TypeOf ActiveControl Is TextBox Then
                bx = DirectCast(ActiveControl, TextBox)
                s37(bx, Math.Sign(e.Delta))
            End If
        ElseIf Not Form1.sbl(3) Then
            Dim bx As TextBox
            If TypeOf ActiveControl Is TextBox Then
                Dim at As TextBox = DirectCast(ActiveControl, TextBox)
                If Not (Form1.sbl(1) AndAlso at.ForeColor.R = 0 AndAlso at.Text > "" AndAlso at.BackColor = Color.FromArgb(165, 215, 175)) Then bx = at
            End If
            Dim i As Integer = s4(TC2.SelectedIndex, 0)
            If bx IsNot Nothing Then
                If bx Is DirectCast(rg(2, i), TextBox) Then
                    s37(bx, Math.Sign(e.Delta))
                    s2(rg(8, i), bx, DirectCast(rg(3, i), Button), DirectCast(rg(4, i), Button), rg(0, i))
                ElseIf bx.BorderStyle = BorderStyle.Fixed3D Then
                    If DirectCast(rg(3, i), Button).Enabled Then
                        bx.BackColor = Color.White
                        s37(bx, Math.Sign(e.Delta))
                    ElseIf bx.BackColor = Color.White Then
                        s37(bx, Math.Sign(e.Delta))
                    Else
                        MsgBox("请先输入批号初始化！")
                    End If
                End If
            End If
        End If
    End Sub
    Sub s1(T As TextBox, B1 As Button, ByRef n As Object, B2 As Button, B3 As Button, dit As Dictionary(Of Control, String), ByRef ph As Object, ct As Control)
        Dim flag As Boolean = True
        Dim i As Integer = s4(n, 10)
        Dim str As String = CStr(rg(9, i))
        If T.Text > "" Then
            s34(T, str, n, flag)
            If flag Then
                B1.Enabled = True
                B2.Enabled = True
                s40(dit, T.Parent, T)
                Try
                    cnct.Open()
                    B3.Text = CStr(New SqlCommand("select Id from [" & str & "工艺] where [" & str & "批号]='" & Replace(T.Text, "'", "''") & "'", cnct).ExecuteScalar)
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                End Try
                If rg(6, i) IsNot Nothing Then RemoveHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
                s5(ct, T, str, dit, B3, False,, B3.Text > "")
                If rg(6, i) IsNot Nothing Then AddHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf TB_TextChanged
                ph = T.Text
            End If
        End If
    End Sub
    Sub s2(ByRef ph As Object, Tn As TextBox, B1 As Button, B2 As Button, Optional ByRef m As Object = Nothing)
        If IsNothing(m) Then
            LI1.Hide()
            RemoveHandler Tn.TextChanged, AddressOf T_TextChanged
            Tn.Text = CStr(LI1.SelectedItem)
            Tn.Focus()
            Tn.SelectionStart = Tn.TextLength
            AddHandler Tn.TextChanged, AddressOf T_TextChanged
        Else
            n = CInt(m)
        End If
        If CStr(ph) <> Tn.Text Then
            B1.Enabled = False
            B2.Enabled = False
        ElseIf CStr(ph) > "" Then
            B1.Enabled = True
            B2.Enabled = True
        End If
    End Sub
    Sub s3(Tn As Control, ByRef str As String, Optional ByRef i As Integer = 0)
        If i <> 0 Then n = i
        LI1.Items.Clear()
        If Tn.Text = "" OrElse Tn.BackColor = Color.FromArgb(165, 215, 175) Then
            LI1.Hide()
        Else
            Try
                cnctm.Open()
                cmdstr = String.Concat("select top 10批号 from(select[", str, "批号]as 批号 from[", str, "工艺]where[", str, "批号]COLLATE Chinese_PRC_CI_AS like '%'+@批号+'%'")
                If i = 0 Then cmdstr = String.Concat(cmdstr, " union select 批号 from 物料数量 where 批号 COLLATE Chinese_PRC_CI_AS like'%'+@批号+'%' and 操作工序='" & Replace(str, "'", "''") & "'")
                cmdstr = String.Concat(cmdstr, ")as T order by 批号 desc")
                cmd = New SqlCommand(cmdstr, cnctm)
                cmd.Parameters.AddWithValue("批号", Tn.Text)
                dr = cmd.ExecuteReader()
                While dr.Read()
                    LI1.Items.Add(dr(0))
                End While
                cnctm.Close()
                If LI1.Items.Count = 0 Then
                    LI1.Hide()
                Else
                    LI1.SetBounds(Tn.Left + 16, Tn.Top + 63, Tn.Size.Width, 116)
                    LI1.Show()
                End If
            Catch ex As Exception
                cnctm.Close()
                MsgBox(String.Concat("无法搜索相关信息" & vbCrLf & "", ex.Message), MsgBoxStyle.OkOnly, Nothing)
            End Try
        End If
    End Sub
    Function s4(sender As Object, ByRef j As Integer) As Integer
        For i = 0 To UBound(rg, 2)
            Try
                If rg(j, i) Is sender OrElse CStr(sender) = CStr(rg(j, i)) Then Return i
            Catch ex As Exception
            End Try
        Next
        Return -1
    End Function
    Sub s5(ct As Control, tx As TextBox, ByRef str As String, dit As Dictionary(Of Control, String), ByRef Bn As Button, Optional ByRef B As Boolean = True, Optional ByRef exflag As Boolean = False, Optional ByRef sign As Boolean = False, Optional ByRef mark As Boolean = False)
        For Each current As Control In ct.Controls
            If current.Controls.Count > 0 Then
                s5(current, tx, str, dit, Bn, B, exflag, sign, mark)
            Else
                If TypeOf current Is TextBox AndAlso DirectCast(current, TextBox).Tag IsNot Nothing Then
                    If CStr(DirectCast(current, TextBox).Tag).Contains("|") Then
                        s38(dit, current, tx)
                    Else
                        If Not B Then current.BackColor = Color.FromArgb(165, 215, 175)
                        If current.BackColor = Color.FromArgb(165, 215, 175) Then
                            s6(dit, current, tx, Bn, str, sign)
                        Else
                            s11(dit, current, tx, Bn, str, exflag, sign, mark)
                        End If
                    End If
                End If
            End If
        Next
    End Sub
    Sub s6(dit As Dictionary(Of Control, String), bt As Control, tx As TextBox, ByRef Bn As Button, ByRef str As String, ByRef sign As Boolean)
        Dim str2(1) As Object, dtr As DataRow()
        cmdstr = "select " & CStr(bt.Tag) & ",(select b.name from syscolumns a,systypes b,sysobjects d where a.xtype=b.xusertype and a.id=d.id and d.name='" & Replace(str, "'", "''") & "工艺' and a.name='" & Replace(Replace(Replace(CStr(bt.Tag), "[", ""), "]", ""), "'", "''") & "') from [" & str & "工艺] where [" & str & "批号]='" & Replace(tx.Text, "'", "''") & "'" & CStr(IIf(Bn.Text = "", "", " and Id=" & Bn.Text))
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            While dr.Read
                str2(0) = dr(0)
                str2(1) = dr(1)
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        If IsNothing(str2(1)) Then
            If sign Then
                MsgBox("该" & str & "批次已删除，更新已中止！")
                dtr = pdt.Select("BN='" & Replace(tx.Text, "'", "''") & "'")
                For Each row As DataRow In dtr
                    pdt.Rows.Remove(row)
                Next
                Bn.Text = ""
                sign = False
            ElseIf Bn.Text = "" Then
                bt.Text = ""
                bt.ForeColor = Color.FromName("WindowText")
            End If
        Else
            bt.Text = CStr(IIf(IsDBNull(str2(0)), "", str2(0)))
            If pdt.Rows.Count = 0 Then
                bt.ForeColor = Color.FromName("WindowText")
            Else
                dtr = pdt.Select(String.Concat(New String() {"BN = '", Replace(tx.Text, "'", "''"), "' and Name = '", bt.Name, "'"}))
                If dtr.Count = 0 Then
                    bt.ForeColor = Color.FromName("WindowText")
                ElseIf s36(dtr(0)(2).ToString, str2, True) Then
                    bt.ForeColor = Color.Red
                Else
                    bt.ForeColor = Color.FromName("WindowText")
                    pdt.Rows.Remove(dtr(0))
                End If
            End If
        End If
        bt.BackColor = Color.FromArgb(165, 215, 175)
        If dit.ContainsKey(bt) Then
            dit(bt) = bt.Text
        Else
            dit.Add(bt, bt.Text)
        End If
    End Sub
    Function s7(ct As Control) As Boolean
        Return Form1.sbl(1) AndAlso Not s35(ct) OrElse Form1.sbl(0)
    End Function
    Sub s8(dit As Dictionary(Of Control, String), tb As Control, Tn As TextBox)
        For Each ct As Control In tb.Controls
            If ct.Controls.Count = 0 Then
                If ct.Tag IsNot Nothing AndAlso CStr(ct.Tag).Contains("|") Then s38(dit, ct, Tn)
            Else
                s8(dit, ct, Tn)
            End If
        Next
    End Sub
    Sub s9(Tn As TextBox, ByRef str As String, tcpg As Control, dit As Dictionary(Of Control, String), Bn As Button)
        Dim i As Boolean
        Dim j As Boolean = True
        If Not Form1.sbl(3) Then
            Try
                Do While s33(str, Tn, Bn, j)
                    cnct.Open()
                    i = CBool(New SqlCommand("insert into [" & str & "工艺]([" & str & "批号]) values('" & Replace(Tn.Text, "'", "''") & "')", cnct).ExecuteNonQuery)
                    cnct.Close()
                Loop
            Catch ex As Exception
                cnct.Close()
                MsgBox(ex.Message)
            End Try
            s5(tcpg, Tn, str, dit, Bn,, True, j, i)
        End If
    End Sub
    Sub s10(e As DataGridViewCellEventArgs)
        RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
        RemoveHandler DA1.CellMouseClick, AddressOf DA1_CellMouseClick
        RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
        RemoveHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
        DA1.Columns(e.ColumnIndex).Visible = True
        DA1.CurrentCell = DA1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        DA1.BeginEdit(False)
        dgvcell = DA1.Rows(e.RowIndex).Cells(e.ColumnIndex) : Form1.skip(1) = True
        G7.Enabled = False
        For i = 1 To DA1.Columns.Count
            DA1.Columns(i - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        AddHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
        AddHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
        AddHandler DA1.CellMouseClick, AddressOf DA1_CellMouseClick
        AddHandler DA1.RowValidating, AddressOf DA1_RowValidating
    End Sub
    Sub s11(dit As Dictionary(Of Control, String), bt As Control, tx As TextBox, ByRef Bn As Button, ByRef str As String, ByRef exflag As Boolean, ByRef sign As Boolean, ByRef mark As Boolean)
        Dim bl As Boolean, str2(1) As Object, str1 As String, dtr As DataRow(), flag As Boolean = True
        cmdstr = "begin tran select " & CStr(bt.Tag) & ",(select b.name from syscolumns a,systypes b,sysobjects d where a.xtype=b.xusertype and a.id=d.id and d.name='" & Replace(str, "'", "''") & "工艺'and a.name='" & Replace(Replace(Replace(CStr(bt.Tag), "[", ""), "]", ""), "'", "''") & "')from[" & str & "工艺]with(tablockx)where[" & str & "批号]='" & Replace(tx.Text, "'", "''") & "'" & CStr(IIf(Bn.Text = "", "", " and Id=" & Bn.Text))
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            While dr.Read
                str2(0) = dr(0)
                str2(1) = dr(1)
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox(ex.Message)
            Return
        End Try
        If IsNothing(str2(1)) Then
            If sign Then
                MsgBox("该" & str & "批次已删除，更新已中止！")
                dtr = pdt.Select("BN='" & Replace(tx.Text, "'", "''") & "'")
                For Each row As DataRow In dtr
                    pdt.Rows.Remove(row)
                Next
                sign = False
                Bn.Text = ""
            End If
            Try
                cnct.Open()
                bl = CBool(New SqlCommand("begin tran commit tran", cnct).ExecuteNonQuery)
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        Else
            If pdt.Rows.Count > 0 Then dtr = pdt.Select(String.Concat(New String() {"BN = '", Replace(tx.Text, "'", "''"), "' and Name = '", bt.Name, "'"}))
            If dtr IsNot Nothing AndAlso dtr.Count > 0 Then
                bl = s36(dtr(0)(2).ToString, str2, True)
            ElseIf dit.ContainsKey(bt) Then
                bl = s36(dit(bt), str2, True) OrElse mark AndAlso IsDBNull(str2(0))
            Else
                bl = IsDBNull(str2(0)) OrElse s36(bt.Text, str2)
            End If
            If Not bl Then
                If Form1.sbl(1) Then
                    flag = True
                ElseIf Form1.sbl(0) Then
                    flag = s36(bt.Text, str2) OrElse MsgBox("从服务器检索到更新的数据，是否跳过？，新数据为：" & vbCrLf & "" & CStr(bt.Tag) & ":" & CStr(IIf(IsDBNull(str2(0)), "NULL", str2(0))) & "。", MsgBoxStyle.YesNo) = MsgBoxResult.Yes
                End If
                If Not flag Then
                    bl = True
                Else
                    bt.Text = CStr(IIf(IsDBNull(str2(0)), "", str2(0)))
                    bt.ForeColor = Color.FromName("WindowText")
                    bt.BackColor = Color.FromArgb(165, 215, 175)
                    If pdt.Rows.Count > 0 AndAlso dtr.Count > 0 Then pdt.Rows.Remove(dtr(0))
                End If
            End If
            If bl Then
                str1 = If(bt.Text = "", "NULL", "'" + Replace(bt.Text, "'", "''") + "'")
                If Not s36(bt.Text, str2) Then
                    cmdstr = String.Concat(New String() {"update [", str, "工艺] set ", CStr(bt.Tag), "=", str1, " where [", str, "批号]='", Replace(tx.Text, "'", "''"), "'"})
                    cmd = New SqlCommand(cmdstr, cnct)
                    Try
                        cnct.Open()
                        cmd.ExecuteNonQuery()
                        cnct.Close()
                    Catch ex As Exception
                        cnct.Close()
                        If mark Then dit(bt) = ""
                        If exflag OrElse Not flag Then
                            exflag = False
                            MsgBox(String.Concat(CStr(bt.Tag), " 输入不正确，请更改后再输！", vbCrLf, ex.Message), MsgBoxStyle.OkOnly)
                            bt.Focus()
                        End If
                        Return
                    End Try
                    If pdt.Rows.Count = 0 OrElse dtr.Count = 0 Then
                        pdt.Rows.Add(tx.Text, bt.Name, bt.Text)
                    Else
                        dtr(0)(2) = bt.Text
                    End If
                Else
                    bt.Text = CStr(IIf(IsDBNull(str2(0)), "", str2(0)))
                    If bt.Focused Then DirectCast(bt, TextBox).SelectionStart = DirectCast(bt, TextBox).TextLength
                End If
                If pdt.Rows.Count > 0 AndAlso pdt.Select(String.Concat(New String() {"BN = '", Replace(tx.Text, "'", "''"), "' and Name = '", bt.Name, "'"})).Count > 0 Then bt.ForeColor = Color.Red
                bt.BackColor = Color.FromArgb(165, 215, 175)
            Else
                Try
                    cnct.Open()
                    bl = CBool(New SqlCommand("begin tran commit tran", cnct).ExecuteNonQuery)
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                End Try
            End If
        End If
        If dit.ContainsKey(bt) Then
            dit(bt) = bt.Text
        Else
            dit.Add(bt, bt.Text)
        End If
    End Sub
    Sub s12(ct As Control, tx As TextBox)
        For Each bt As Control In ct.Controls
            If bt.Controls.Count = 0 Then
                If bt.Tag IsNot Nothing Then
                    RemoveHandler bt.TextChanged, AddressOf TC_TextChanged
                    bt.Text = ""
                    AddHandler bt.TextChanged, AddressOf TC_TextChanged
                    bt.BackColor = Color.FromArgb(165, 215, 175)
                End If
            Else
                s12(bt, tx)
            End If
        Next
    End Sub
    Sub s14(ct As Control, tx As TextBox, Bn As Button, ByRef str As String, dit As Dictionary(Of Control, String))
        For Each bt As Control In ct.Controls
            If bt.Controls.Count > 0 Then
                s14(bt, tx, Bn, str, dit)
            Else
                If TypeOf bt Is TextBox AndAlso DirectCast(bt, TextBox).BorderStyle = BorderStyle.Fixed3D AndAlso bt.BackColor = Color.FromArgb(165, 215, 175) Then s6(dit, bt, tx, Bn, str, False)
            End If
        Next
    End Sub
    Private Sub Form2_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        s59(Visible, Form1.LI4.Items.Contains("入库"))
    End Sub
End Class