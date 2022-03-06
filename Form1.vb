Imports System.Data.SqlClient
Imports Aspose.Cells
Public Class Form1
    Friend WithEvents TSMI, TSMIN, TSMIA, TSMIB, TSMIC, TSMID As ToolStripMenuItem
    Dim cmdstrgx As String, er, ec, ex, tci As Integer, CR, CG, CB As Byte, TBEC, cbl As TextBox, sv As Object, lgc, ctlbl, rb, rc, whbl, clbl As Boolean, tb1, tb2 As New DataTable, dacell As New List(Of DataGridViewCell), dgvcell As DataGridViewCell, gd As New Dictionary(Of String, String), ary As New Dictionary(Of Object, DataTable), TN As New Dictionary(Of TextBox, Object())
    Public st() As String = Form0.st, cnctk As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=msdb;user id=calc" & ";timeout=1"), lct As Point, ni As Integer, dttm As String, dacl As New Dictionary(Of DataGridView, List(Of DataGridViewCell)), usr As String = Form0.C2.Text, suer As Integer = Fcsb.s1(usr), pswd As String = Form0.T4.Text, da As SqlDataAdapter, dtn, dto, pdt As New DataTable, ri As Integer = -1, fc, nn, flg, tcs, bbbl, sbl(3), skip(1), flct, ctbl, b105bl, L124bl, ckbl, ttbl, ccbl, ccbl2, lbl126, bcbl, dabl As Boolean, scm As New Dictionary(Of Color, Byte), bn As New Dictionary(Of String, String), lbl As New Dictionary(Of Object, Object()), idt1, idt2, idt3, idt4 As New List(Of Integer), dacw As New Dictionary(Of DataGridView, List(Of Integer)), cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & usr & ";password=" & pswd & ";timeout=1")
    Public Sub B14_Click(sender As Object, e As EventArgs) Handles B14.Click
        If Not ctbl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            ctbl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        Fcsb.s39(True)
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim na, nb, nc As Boolean
        na = DA3.Rows.Count > 1
        nb = DA5.Rows.Count > 1
        nc = DA6.Columns.Count > 6
        If Not nc Then
            For Each cell As DataGridViewCell In DA6.Rows(0).Cells
                If CStr(cell.Value) <> "" Then
                    nc = True
                    Exit For
                End If
            Next
        End If
        If (na OrElse nb OrElse nc) AndAlso Not fc Then
            TC1.SelectedIndex = 2
            Dim msgr As MsgBoxResult = MsgBox("有未提交的行，是否继续退出？", MsgBoxStyle.OkCancel)
            If msgr = MsgBoxResult.Cancel Then
                e.Cancel = True : lgc = False : Return
            ElseIf lgc Then
                Hide()
                Form0.Show()
            End If
        ElseIf lgc Then
            lgc = False
            Hide()
            Form0.Show()
        End If
        Form2.Close()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim k0 As New List(Of String)
        Text += "—" & usr
        clbl = Form0.CL1.CheckedItems.Count = 0
        If clbl Then
            For Each r In Form0.CL1.Items
                CL2.Items.Add(r)
                If suer <> 4 AndAlso suer <> 5 Then CL4.Items.Add(r)
                CB6.Items.Add(r)
            Next
        Else
            G2.Enabled = False
            For Each r In Form0.CL1.CheckedItems
                CL2.Items.Add(r)
                If suer <> 4 AndAlso suer <> 5 Then CL4.Items.Add(r)
                CB6.Items.Add(r)
            Next
        End If
        CB6.Items.Remove("全部")
        CB6.Items.Add("")
        CL2.SetItemChecked(0, True)
        If suer <> 4 AndAlso suer <> 5 Then CL4.SetItemChecked(0, True)
        For Each r In CL2.CheckedItems
            k0.Add(CStr(r))
            LI5.Items.Add(r)
            DirectCast(DA1.Columns.Item(8), DataGridViewComboBoxColumn).Items.Add(r)
        Next
        LI5.Items.Remove("全部") : k0.Remove("全部")
        DirectCast(DA1.Columns.Item(8), DataGridViewComboBoxColumn).Items.Remove("全部")
        DirectCast(DA1.Columns.Item(8), DataGridViewComboBoxColumn).Items.Add("")
        cmdstrgx = Fcsb.s2(k0, "操作工序.操作工序")
        DA6.Rows.Add()
        Try
            cnct.Open()
            cmdstr = "select 物料类型 from 物料类型 where 可用性=1 order by id"
            Fcsb.s3(LI3, cmdstr)
            Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(6), DataGridViewComboBoxColumn))
            tb1.Columns.Add("物料名称")
            tb2.Columns.Add("物料名称")
            For Each r In LI3.Items
                tb1.Rows.Add(r)
            Next
            For Each r In CL2.CheckedItems
                tb2.Rows.Add(r)
            Next
            If LI3.Items.Count > 0 Then
                cmd = New SqlCommand("消耗产量", cnct)
                cmd.Parameters.Add(New SqlParameter("操作工序", tb2))
                cmd.Parameters.Add(New SqlParameter("物料类型", tb1))
                cmd.Parameters.Add(New SqlParameter("类型", 1))
                cmd.CommandType = CommandType.StoredProcedure
                Fcsb.s53(dtn, cmd)
                dr = cmd.ExecuteReader
                While dr.Read
                    LI1.Items.Add(dr(0))
                    DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Add(dr(0))
                End While
                dr.Close()
                DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Add("")
            End If
            cmd = New SqlCommand("消耗产量", cnct)
            cmd.Parameters.Add(New SqlParameter("操作工序", tb2))
            cmd.Parameters.Add(New SqlParameter("物料类型", tb1))
            cmd.Parameters.Add(New SqlParameter("类型", CByte(IIf(clbl, 2, 1))))
            cmd.CommandType = CommandType.StoredProcedure
            dr = cmd.ExecuteReader
            While dr.Read
                CL1.Items.Add(dr(0))
            End While
            dr.Close()
            cmdstr = "select 反应釜号,id from 反应釜号 order by id"
            Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(9), DataGridViewComboBoxColumn))
            Fcsb.s6(cmdstr, DirectCast(DA5.Columns.Item(3), DataGridViewComboBoxColumn))
            Fcsb.s6(cmdstr, DirectCast(DA6.Columns.Item(3), DataGridViewComboBoxColumn))
            cmdstr = "select 班别班组,id from 班别班组 order by id"
            Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(7), DataGridViewComboBoxColumn))
            Fcsb.s6(cmdstr, DirectCast(DA5.Columns.Item(2), DataGridViewComboBoxColumn))
            Fcsb.s6(cmdstr, DirectCast(DA6.Columns.Item(2), DataGridViewComboBoxColumn))
            If suer <> 6 Then
                If Form0.CL1.CheckedItems.Count > 0 Then
                    cmdstr = "select 储槽名称,位号 from 储槽特性,操作工序 where " & cmdstrgx & " and 储槽特性.操作工序=操作工序.操作工序 and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.物料名称"
                Else
                    cmdstr = "select 储槽名称,位号 from 储槽特性,操作工序 where 储槽特性.操作工序=操作工序.操作工序 and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.物料名称"
                End If
                Fcsb.s53(dto, New SqlCommand(cmdstr, cnct))
                Fcsb.s3(LI7, cmdstr)
                Fcsb.s6(cmdstr, DirectCast(DA2.Columns.Item(2), DataGridViewComboBoxColumn))
                s31(CO8)
                For a = 0 To 4
                    DA5.Columns(a).Frozen = True
                    DA6.Columns(a).Frozen = True
                Next
                DA6.Columns(5).Frozen = True
                cmdstr = "select 报表名称 from 报表配置 where 报表类型=0"
                Fcsb.s6(cmdstr, DirectCast(DA9.Columns.Item(2), DataGridViewComboBoxColumn))
            End If
            If sbl(0) Then
                cmdstr = "select 物料名称 from 物料特性 where 可用性=0 order by id"
                Fcsb.s3(LI9, cmdstr)
                cmdstr = "select 物料名称 from 物料特性 where 可用性=1 order by id"
                Fcsb.s3(LI12, cmdstr)
                cmdstr = "select 储槽名称 from 储槽特性 where 可用性=0 order by id"
                Fcsb.s3(LI10, cmdstr)
                cmdstr = "select 储槽名称 from 储槽特性 where 可用性=1 order by id"
                Fcsb.s3(LI13, cmdstr)
                cmdstr = "select 储槽名称 from 储槽特性 order by id"
                s6(CB2)
                cmdstr = "select 物料名称 from 物料特性 order by id"
                s6(CB1)
                s6(CB3)
                cmdstr = "select 操作工序 from 操作工序 order by id"
                s6(CB4)
                s6(CB5)
            End If
            If suer > 0 Then
                B50.Enabled = False
                B103.Enabled = False
                B104.Enabled = False
                L128.Enabled = False
                DA11.ReadOnly = True
                DA11.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            End If
            If sbl(2) Then
                DA1.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
                DA2.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            End If
            If suer = 4 OrElse suer = 5 Then
                B92.Enabled = False
                B97.Enabled = False
                CL4.Enabled = False
                T45.Enabled = False
                LI19.Enabled = False
                CH39.Enabled = False
                L125.Enabled = False
            End If
            If sbl(0) OrElse sbl(1) Then
                If Form0.CL1.CheckedItems.Count > 0 Then
                    cmdstr = "select 储槽名称,储槽特性.id from 储槽特性,操作工序 where 储槽特性.操作工序=操作工序.操作工序 and " & cmdstrgx & " and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.id"
                Else
                    cmdstr = "select 储槽名称,储槽特性.id from 储槽特性,操作工序 where 储槽特性.操作工序=操作工序.操作工序 and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.id"
                End If
                s14()
                s15(tb1, tb2)
                Fcsb.s29()
            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox(ex.Message)
        End Try
        D1.Text = Format(DateAdd(DateInterval.Day, -2, Now), "yyyy-MM-dd 00:00")
        D2.Text = Format(DateAdd(DateInterval.Day, 1, Now), "yyyy-MM-dd 00:00")
        D3.Text = Format(DateAdd(DateInterval.Minute, -3359, Now), "yyyy-MM-dd 07:59")
        D4.Text = Format(DateAdd(DateInterval.Minute, 961, Now), "yyyy-MM-dd 07:59")
        D6.Value = DateAdd(DateInterval.Day, -1, Now)
        D7.Value = DateAdd(DateInterval.Day, -1, Now)
        D9.Text = Format(DateAdd(DateInterval.Day, -2, Now), "yyyy-MM-dd 00:00")
        D10.Text = Format(DateAdd(DateInterval.Day, 1, Now), "yyyy-MM-dd 00:00")
        If sbl(3) Then
            B16.Enabled = False : B26.Enabled = False
            G1.Enabled = False : B80.Enabled = False
            D9.Enabled = False : D10.Enabled = False
            G2.Enabled = False : G3.Enabled = False
        End If
        pdt.Columns.Add("BN", Type.GetType("System.String"))
        pdt.Columns.Add("Name", Type.GetType("System.String"))
        pdt.Columns.Add("CT", Type.GetType("System.String"))
        If suer = 4 Then
            T38.Enabled = False : B28.Enabled = False
            D1.Enabled = False : D2.Enabled = False
            D3.Enabled = False : D4.Enabled = False
            D1.Checked = False : D2.Checked = False
            D3.Checked = False : D4.Checked = False
            T39.Enabled = False : T40.Enabled = False
            T52.Enabled = False : T53.Enabled = False
            B99.Enabled = False : B13.Enabled = False
            B13.Text = "班别班组"
        End If
        If suer > 2 Then
            CH32.Enabled = False
            CH33.Enabled = False
        End If
        T28.Enabled = sbl(0)
        If sbl(0) Then
            s26(CL5, "操作工序")
            s26(CL6, "物料类型")
        End If
        SFD.Filter = "Excel 99-03文件|*.xls|Excel 2007文件|*.xlsx|pdf文档|*.pdf"
        SFD.DefaultExt = "xls"
        cnct.Open()
        cmd = New SqlCommand("select 反应釜号,id from 反应釜号 order by id", cnct)
        dr = cmd.ExecuteReader
        While dr.Read
            CO9.Items.Add(dr(0))
            If suer <> 6 Then CO11.Items.Add(dr(0))
        End While
        cnct.Close()
        CO9.Items.Add(" ")
        If suer <> 6 Then CO11.Items.Add(" ")
        cnct.Open()
        cmd = New SqlCommand("select 班别班组,id from 班别班组 order by id", cnct)
        dr = cmd.ExecuteReader
        While dr.Read
            T1.Items.Add(dr(0))
            If suer <> 6 Then CO10.Items.Add(dr(0))
        End While
        cnct.Close()
        T1.Items.Add(" ")
        If suer <> 6 Then CO10.Items.Add(" ")
        dacw.Add(DA1, New List(Of Integer)) : Fcsb.s56(DA1)
        dacw.Add(DA2, New List(Of Integer)) : Fcsb.s56(DA2)
        dacw.Add(DA9, New List(Of Integer)) : Fcsb.s56(DA9)
        dacw.Add(DA11, New List(Of Integer))
        If sbl(2) Then L126.Enabled = False : L127.Enabled = False
        If suer = 4 Then L128.Enabled = False
        lbl.Add(L126, {False, DA1})
        lbl.Add(L127, {False, DA2})
        lbl.Add(L128, {False, DA11})
        lbl.Add(L130, {False})
        lbl.Add(L125, {False})
        flg = True
        If suer = 4 OrElse suer = 5 Then L125.Enabled = False
        TN.Add(T46, {T47, "@液位", "储槽"}) : TN.Add(T47, {T46, "@物料数量", "液位"})
        If UBound(st) = 4 Then T62.Text = "班别(以" & st(4) & "为准):" & Fcsb.s24(,, st(4), D11)
        s41()
        Try
            cnct.Open()
            dr = New SqlCommand("select 报表名称,报表标签 from 报表配置 where 报表类型=1", cnct).ExecuteReader
            While dr.Read
                gd.Add(CStr(dr(0)), CStr(dr(1)))
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        CH20.Tag = {"原料动态表"}
        CH21.Tag = {""}
        CH22.Tag = {""}
        For Each G As Control In G9.Controls
            For Each R As RadioButton In G.Controls
                AddHandler R.MouseDown, AddressOf R_MouseDown
                AddHandler R.MouseUp, AddressOf R_MouseUp
            Next
        Next
        dacl.Add(DA1, New List(Of DataGridViewCell))
        dacl.Add(DA2, New List(Of DataGridViewCell))
        dacl.Add(DA3, New List(Of DataGridViewCell))
        dacl.Add(DA5, New List(Of DataGridViewCell))
        dacl.Add(DA6, New List(Of DataGridViewCell))
        dacl.Add(DA9, New List(Of DataGridViewCell))
        dacl.Add(DA10, New List(Of DataGridViewCell))
        dacl.Add(DA11, New List(Of DataGridViewCell))
        dacl.Add(DA12, New List(Of DataGridViewCell))
        For Each DA As DataGridView In dacl.Keys
            AddHandler DA.MouseWheel, AddressOf MouseWheel
            AddHandler DA.RowPostPaint, AddressOf RowPostPaint
            AddHandler DA.CellMouseEnter, AddressOf CellMouseEnter
            AddHandler DA.CellMouseLeave, AddressOf CellMouseLeave
        Next
        If Screen.PrimaryScreen.Bounds.Width <= Width OrElse Screen.PrimaryScreen.Bounds.Height <= Height Then MsgBox("屏幕分辨率不得小于1114×663！")
    End Sub
    Private Sub B1_Click(sender As Object, e As EventArgs) Handles B1.Click, LI1.DoubleClick
        s1(LI1, LI2)
    End Sub
    Private Sub B2_Click(sender As Object, e As EventArgs) Handles B2.Click, LI2.DoubleClick
        s1(LI2, LI1)
    End Sub
    Private Sub B5_Click(sender As Object, e As EventArgs) Handles B5.Click, LI3.DoubleClick
        s1(LI3, LI4)
    End Sub
    Private Sub B6_Click(sender As Object, e As EventArgs) Handles B6.Click, LI4.DoubleClick
        s1(LI4, LI3)
    End Sub
    Private Sub B9_Click(sender As Object, e As EventArgs) Handles B9.Click, LI5.DoubleClick
        s1(LI5, LI6)
    End Sub
    Private Sub B10_Click(sender As Object, e As EventArgs) Handles B10.Click, LI6.DoubleClick
        s1(LI6, LI5)
    End Sub
    Private Sub B20_Click(sender As Object, e As EventArgs) Handles B20.Click, LI7.DoubleClick
        s1(LI7, LI8)
        RemoveHandler T51.TextChanged, AddressOf T51_TextChanged
        T51.Text = "储槽名称："
        AddHandler T51.TextChanged, AddressOf T51_TextChanged
    End Sub
    Private Sub B21_Click(sender As Object, e As EventArgs) Handles B21.Click, LI8.DoubleClick
        s1(LI8, LI7)
    End Sub
    Private Sub B3_Click(sender As Object, e As EventArgs) Handles B3.Click
        s2(LI1, LI2)
    End Sub
    Private Sub B4_Click(sender As Object, e As EventArgs) Handles B4.Click
        s2(LI2, LI1)
    End Sub
    Private Sub B7_Click(sender As Object, e As EventArgs) Handles B7.Click
        s2(LI3, LI4)
    End Sub
    Private Sub B8_Click(sender As Object, e As EventArgs) Handles B8.Click
        s2(LI4, LI3)
    End Sub
    Private Sub B11_Click(sender As Object, e As EventArgs) Handles B11.Click
        s2(LI5, LI6)
    End Sub
    Private Sub B12_Click(sender As Object, e As EventArgs) Handles B12.Click
        s2(LI6, LI5)
    End Sub
    Private Sub B22_Click(sender As Object, e As EventArgs) Handles B22.Click
        s2(LI7, LI8)
        RemoveHandler T51.TextChanged, AddressOf T51_TextChanged
        T51.Text = "储槽名称："
        AddHandler T51.TextChanged, AddressOf T51_TextChanged
    End Sub
    Private Sub B23_Click(sender As Object, e As EventArgs) Handles B23.Click
        s2(LI8, LI7)
    End Sub
    Private Sub B16_Click(sender As Object, e As EventArgs) Handles B16.Click
        Fcsb.s7(DirectCast(sender, Button), B17, DA1, idt1)
    End Sub
    Private Sub B26_Click(sender As Object, e As EventArgs) Handles B26.Click
        Fcsb.s7(DirectCast(sender, Button), B27, DA2, idt2)
        DA2.Columns(4).ReadOnly = True
        DA2.Columns(5).ReadOnly = True
        DA2.Columns(6).ReadOnly = True
    End Sub
    Private Sub B17_Click(sender As Object, e As EventArgs) Handles B17.Click
        Fcsb.s4(DA1, "物料数量", idt1)
    End Sub
    Private Sub B27_Click(sender As Object, e As EventArgs) Handles B27.Click
        Fcsb.s4(DA2, "储槽液位", idt2)
    End Sub
    Private Sub B15_Click(sender As Object, e As EventArgs) Handles B15.Click
        Fcsb.s8(DA1, True)
    End Sub
    Private Sub B25_Click(sender As Object, e As EventArgs) Handles B25.Click
        Fcsb.s8(DA2, True)
    End Sub
    Private Sub B19_Click(sender As Object, e As EventArgs) Handles B19.Click
        lgc = True
        Close()
    End Sub
    Private Sub CLA_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CL1.ItemCheck, CL2.ItemCheck, CL3.ItemCheck, CL4.ItemCheck
        If sender Is CL2 Then LI15.Items.Clear() : LI15.Hide()
        Dim CL As CheckedListBox = DirectCast(sender, CheckedListBox)
        RemoveHandler CL.ItemCheck, AddressOf CLA_ItemCheck
        If e.Index = 0 Then
            If e.NewValue = CheckState.Checked Then
                For i = 1 To CL.Items.Count - 1
                    CL.SetItemChecked(i, True)
                Next
            Else
                For i = 1 To CL.Items.Count - 1
                    CL.SetItemChecked(i, False)
                Next
            End If
        Else
            Fcsb.s12(CL, e)
        End If
        AddHandler CL.ItemCheck, AddressOf CLA_ItemCheck
    End Sub
    Public Sub CL2_MouseUp(sender As Object, e As EventArgs) Handles CL2.MouseUp, CL2.KeyUp
        If IsNothing(TryCast(e, KeyEventArgs)) OrElse DirectCast(e, KeyEventArgs).KeyCode = Keys.Space Then
            Dim k0 As New List(Of String), CL As CheckedListBox = DirectCast(sender, CheckedListBox)
            If CL.CheckedItems.Count = 0 Then
                For Each r In CL.Items
                    k0.Add(CStr(r))
                Next
            Else
                For Each r In CL.CheckedItems
                    k0.Add(CStr(r))
                Next
            End If
            DA1.Rows.Clear() : DA2.Rows.Clear()
            LI1.Items.Clear() : LI2.Items.Clear()
            LI6.Items.Clear() : LI8.Items.Clear()
            DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Clear()
            k0.Remove("全部")
            cmdstrgx = Fcsb.s2(k0, "操作工序.操作工序")
            Try
                cnct.Open()
                cmdstr = "select 操作工序,id from 操作工序 where " & cmdstrgx & " order by id"
                Fcsb.s3(LI5, cmdstr)
                Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(8), DataGridViewComboBoxColumn))
                tb1.Reset()
                tb1.Columns.Add("物料名称")
                For Each r In LI3.Items
                    tb1.Rows.Add(r)
                Next
                For Each r In LI4.Items
                    tb1.Rows.Add(r)
                Next
                s7(tb1, tb2, clbl AndAlso CL.CheckedItems.Count = 0)
                If suer <> 6 Then
                    If cmdstrgx <> "(" Then
                        If clbl AndAlso CL.CheckedItems.Count = 0 Then
                            cmdstr = "select 储槽名称,位号 from 储槽特性 where 可用性=1 order by 储槽特性.物料名称"
                        Else
                            cmdstr = "select 储槽名称,位号 from 储槽特性,操作工序 where " & cmdstrgx & " and 储槽特性.操作工序=操作工序.操作工序 and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.物料名称"
                        End If
                    Else
                        cmdstr = "select 储槽名称,位号 from 储槽特性,操作工序 where 储槽特性.操作工序=操作工序.操作工序 and 储槽特性.可用性=1 and 操作工序.可用性=1 order by 储槽特性.物料名称"
                    End If
                    Fcsb.s53(dto, New SqlCommand(cmdstr, cnct))
                    Fcsb.s3(LI7, cmdstr)
                    Fcsb.s6(cmdstr, DirectCast(DA2.Columns.Item(2), DataGridViewComboBoxColumn))
                    s15(tb1, tb2)
                    s31(CO8)
                    If cmdstrgx <> "(" Then
                        If clbl AndAlso CL.CheckedItems.Count = 0 Then
                            cmdstr = "select 储槽名称,Id from 储槽特性 where 可用性=1 order by 储槽特性.Id"
                        Else
                            cmdstr = "select 储槽名称,储槽特性.Id from 储槽特性,操作工序 where 储槽特性.操作工序=操作工序.操作工序 and " & cmdstrgx & " and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.Id"
                        End If
                    Else
                        cmdstr = "select 储槽名称,储槽特性.Id from 储槽特性,操作工序 where 储槽特性.操作工序=操作工序.操作工序 and 操作工序.可用性=1 and 储槽特性.可用性=1 order by 储槽特性.Id"
                    End If
                    s14()
                End If
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
            RemoveHandler T50.TextChanged, AddressOf T50_TextChanged
            RemoveHandler T51.TextChanged, AddressOf T51_TextChanged
            T50.Text = "物料名称：" : T51.Text = "储槽名称："
            AddHandler T50.TextChanged, AddressOf T50_TextChanged
            AddHandler T51.TextChanged, AddressOf T51_TextChanged
        End If
    End Sub
    Private Sub B13_Click(sender As Object, e As EventArgs) Handles B13.Click
        Dim n As Integer
        If Integer.TryParse(T1.Text, n) Then
            s8(n)
            DA1.ClearSelection()
        Else
            MsgBox("请正确输入序号！")
        End If
    End Sub
    Public Sub DA1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA1.CellBeginEdit
        s36(DirectCast(sender, DataGridView), e, "select * from 物料数量 where Id=", "yyyy-MM-dd HH:mm", D1, D2)
    End Sub
    Public Sub DA2_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA2.CellBeginEdit
        s36(DirectCast(sender, DataGridView), e, "select 储槽液位.id,日期,储槽液位.储槽名称,储槽液位,dbo.储槽计算(储槽液位.储槽名称,储槽液位,日期),物料名称,操作工序 from 储槽液位,储槽特性 where 储槽液位.储槽名称=储槽特性.储槽名称 and 储槽液位.Id=", "yyyy-MM-dd HH:mm", D3, D4)
    End Sub
    Public Sub DA9_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA9.CellBeginEdit
        s36(DirectCast(sender, DataGridView), e, "select * from 报表备注 where Id=", "yyyy-MM-dd", D5, D6, D7, D8)
    End Sub
    Public Sub DA1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA1.CellEndEdit
        Dim cmdstr, clas As String, dt As Date, num As Decimal, str4 As Object, flag, fg As Boolean, msgBoxResult As MsgBoxResult, DA As DataGridView = DirectCast(sender, DataGridView), str(4) As String, str1() As String = {"", ""}, dac1 As DataGridViewComboBoxColumn = DirectCast(DA.Columns(6), DataGridViewComboBoxColumn), dac2 As DataGridViewComboBoxColumn = DirectCast(DA.Columns(8), DataGridViewComboBoxColumn), dac3 As DataGridViewComboBoxColumn = DirectCast(DA.Columns(7), DataGridViewComboBoxColumn), dac4 As DataGridViewComboBoxColumn = DirectCast(DA.Columns(9), DataGridViewComboBoxColumn)
        If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
            If suer <> 4 Then
                D1.Enabled = True
                D2.Enabled = True
            End If
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
            DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
        End If
        If IsNothing(DA.Rows(e.RowIndex).Cells(0).Value) Then
            Dim str2 As String = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
            Dim str3 As String = Fcsb.s55(str2)
            If e.ColumnIndex = 1 Then
                If DA.NewRowIndex = e.RowIndex Then
                    ni = 1
                    ri = e.RowIndex
                    skip(0) = True
                    TM1.Enabled = True
                ElseIf IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then
                    DA.Rows(e.RowIndex).Cells(1).Value = String.Concat(Format(Now, "yyyy-MM-dd "), "08:00")
                Else
                    DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "08:00")
                End If
                If DA.Rows(e.RowIndex).Cells(8).Value IsNot Nothing AndAlso DA.Rows(e.RowIndex).Cells(1).Value IsNot Nothing Then
                    clas = Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value))
                    DA.Rows(e.RowIndex).Cells(7).Value = IIf(IsNothing(clas) OrElse Not dac3.Items.Contains(clas), Nothing, clas)
                    DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.Black
                End If
            ElseIf e.ColumnIndex = 2 Then
                DA.Rows(e.RowIndex).Cells(2).Value = str2
                If CH33.Checked AndAlso str3 <> "" Then DA.Rows(e.RowIndex).Cells(9).Value = IIf(dac4.Items.Contains(str3), str3, Nothing)
                fg = DA.Rows(e.RowIndex).Cells(3).Value IsNot Nothing
            ElseIf e.ColumnIndex = 3 Then
                Try
                    cnct.Open()
                    str4 = New SqlCommand(String.Concat("select 单釜消耗 from 物料特性 where 物料名称='", Replace(CStr(DA.Rows(e.RowIndex).Cells(3).Value), "'", "''"), "'"), cnct).ExecuteScalar
                    cnct.Close()
                    DA.Rows(e.RowIndex).Cells(4).Value = IIf(str4 Is DBNull.Value, Nothing, str4)
                Catch ex As Exception
                    cnct.Close()
                End Try
                fg = DA.Rows(e.RowIndex).Cells(2).Value IsNot Nothing
            ElseIf e.ColumnIndex = 7 Then
                If CStr(DA.Rows(e.RowIndex).Cells(7).Value) <> Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value)) Then
                    If Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value)) IsNot Nothing Then
                        DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.FromArgb(255, 100, 100)
                    End If
                Else
                    DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.Black
                End If
            ElseIf e.ColumnIndex = 8 Then
                If DA.Rows(e.RowIndex).Cells(8).Value IsNot Nothing AndAlso DA.Rows(e.RowIndex).Cells(1).Value IsNot Nothing Then
                    clas = Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value))
                    DA.Rows(e.RowIndex).Cells(7).Value = IIf(IsNothing(clas) OrElse Not dac3.Items.Contains(clas), Nothing, clas)
                    DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.Black
                End If
            End If
            If fg Then
                Try
                    If CH33.Checked Then
                        Dim n As Integer
                        Dim s As Byte = Fcsb.s10(str2, CH32.Checked)
                        Try
                            cnct.Open()
                            dr = New SqlCommand("select distinct 物料类型,操作工序 from 工序类型 where 批号代码=" & s & " and 物料名称='" & CStr(DA.Rows(e.RowIndex).Cells(3).Value) & "' and 可用性=1", cnct).ExecuteReader
                            While dr.Read
                                n += 1
                                str1(0) = CStr(dr(0))
                                str1(1) = CStr(dr(1))
                            End While
                            cnct.Close()
                            If n > 1 Then
                                str1(0) = Nothing
                                str1(1) = Nothing
                            End If
                        Catch ex As Exception
                            cnct.Close()
                            str1(0) = Nothing
                            str1(1) = Nothing
                        End Try
                        If str1(0) <> Nothing Then
                            DA.Rows(e.RowIndex).Cells(6).Value = IIf(dac1.Items.Contains(str1(0)), str1(0), Nothing)
                            DA.Rows(e.RowIndex).Cells(8).Value = IIf(str1(1) = Nothing OrElse dac2.Items.Contains(str1(1)), str1(1), Nothing)
                        End If
                    End If
                    If DA.Rows(e.RowIndex).Cells(8).Value IsNot Nothing AndAlso DA.Rows(e.RowIndex).Cells(1).Value IsNot Nothing Then
                        clas = Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value))
                        DA.Rows(e.RowIndex).Cells(7).Value = IIf(IsNothing(clas) OrElse Not dac3.Items.Contains(clas), Nothing, clas)
                        DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.Black
                    End If
                Catch ex As Exception
                End Try
            End If
            Return
        ElseIf CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 Then
            CR = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.R
            CG = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.G
            CB = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.B
            skip(1) = False
            skip(0) = False
            str(0) = Replace(CStr(DA.Rows(e.RowIndex).Cells(3).Value), "'", "''")
            str(1) = Replace(CStr(DA.Rows(e.RowIndex).Cells(6).Value), "'", "''")
            str(2) = Replace(CStr(DA.Rows(e.RowIndex).Cells(8).Value), "'", "''")
            str(3) = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
            str(4) = Replace(CStr(DA.Rows(e.RowIndex).Cells(9).Value), "'", "''")
            Do While CH32.Checked AndAlso Fcsb.s10(str(3), e.ColumnIndex = 2) = 0 AndAlso CStr(DA.Rows(e.RowIndex).Cells(2).Value) <> ""
                skip(0) = True
                msgBoxResult = MsgBox("批号格式不正确！", MsgBoxStyle.AbortRetryIgnore)
                If msgBoxResult = MsgBoxResult.Abort Then
                    s11(e)
                    Return
                ElseIf msgBoxResult = MsgBoxResult.Ignore Then
                    skip(0) = False
                    RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    DA.EndEdit()
                    DA.Select()
                    AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    Exit Do
                End If
            Loop
            If CStr(DA.Rows(e.RowIndex).Cells(7).Value) <> Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value),, True) Then
                If Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(7).Value), CStr(DA.Rows(e.RowIndex).Cells(1).Value), CStr(DA.Rows(e.RowIndex).Cells(8).Value), , True) = "" Then
                    GoTo Label1
                End If
                DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.FromArgb(255, 100, 100)
                GoTo Label0
            End If
Label1:
            DA.Rows(e.RowIndex).Cells(7).Style.ForeColor = Color.Black
Label0:
            Do While CH33.Checked AndAlso Fcsb.s15(str, e.ColumnIndex = 2)
                skip(0) = True
                msgBoxResult = MsgBox("输入的条目不匹配", MsgBoxStyle.AbortRetryIgnore)
                If msgBoxResult = MsgBoxResult.Abort Then
                    s11(e)
                    Return
                ElseIf msgBoxResult = MsgBoxResult.Ignore Then
                    skip(0) = False
                    RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    DA.EndEdit()
                    DA.Select()
                    AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    Exit Do
                End If
            Loop
            If e.ColumnIndex = 1 Then
                DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value))
                If Not Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), dt) Then
                    skip(0) = True
                    MsgBox("日期输入有误，请检查后重输！")
                    s11(e)
                    Return
                End If
                If CDate(sv) <> dt Then cmdstr = String.Concat("update 物料数量 set 日期='", CStr(DA.Rows(e.RowIndex).Cells(1).Value), "' where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value))
                DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Second, 30, CDate(DA.Rows(e.RowIndex).Cells(1).Value)), "yyyy-MM-dd HH:mm")
            ElseIf e.ColumnIndex = 2 Then
                DA.Rows(e.RowIndex).Cells(2).Value = str(3)
                If CStr(DA.Rows(e.RowIndex).Cells(2).Value) <> CStr(sv) Then
                    If IsNothing(DA.Rows(e.RowIndex).Cells(2).Value) Then
                        cmdstr = String.Concat("update 物料数量 set ", DA.Columns(2).Name, "=NULL where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value))
                    Else
                        cmdstr = String.Concat(New String() {"update 物料数量 set 批号='", Replace(CStr(DA.Rows(e.RowIndex).Cells(2).Value), "'", "''"), "' where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)})
                    End If
                End If
            ElseIf e.ColumnIndex = 4 Then
                DA.Rows(e.RowIndex).Cells(4).Value = Fcsb.s49(CStr(DA.Rows(e.RowIndex).Cells(4).Value), flag, num)
                If Not flag Then
                    skip(0) = True
                    MsgBox("物料数量输入有误，请检查后重输！")
                    s11(e)
                    Return
                End If
                If num <> CDec(sv) OrElse Not IsNumeric(DA.Rows(e.RowIndex).Cells(4).Value) Then
                    cmdstr = String.Concat(New String() {"update 物料数量 set 物料数量=", CStr(DA.Rows(e.RowIndex).Cells(4).Value), " where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)})
                End If
                DA.Rows(e.RowIndex).Cells(4).Value = CDec(Format(num, "0.000"))
            ElseIf e.ColumnIndex = 5 Then
                Dim cldc As String = Fcsb.s49(CStr(DA.Rows(e.RowIndex).Cells(5).Value), flag, num)
                If Not flag Then
                    If DA.Rows(e.RowIndex).Cells(5).Value IsNot Nothing Then
                        skip(0) = True
                        MsgBox("物料含量输入有误，请检查后重输！")
                        s11(e)
                        Return
                    End If
                Else
                    If num <= 0 Then
                        skip(0) = True
                        MsgBox("物料含量输入有误，请检查后重输！")
                        s11(e)
                        Return
                    End If
                End If
                If num <> CDec(sv) OrElse Not IsNumeric(cldc) AndAlso cldc <> "" Then
                    If IsNothing(DA.Rows(e.RowIndex).Cells(5).Value) Then
                        cmdstr = String.Concat("update 物料数量 set 物料含量=NULL where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value))
                    Else
                        cmdstr = String.Concat(New String() {"update 物料数量 set 物料含量=", cldc, " where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)})
                        DA.Rows(e.RowIndex).Cells(5).Value = CDec(Format(num, "0.00"))
                    End If
                ElseIf DA.Rows(e.RowIndex).Cells(5).Value IsNot Nothing Then
                    DA.Rows(e.RowIndex).Cells(5).Value = CDec(Format(CDec(DA.Rows(e.RowIndex).Cells(5).Value), "0.00"))
                End If
            Else
                If IsNothing(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
                    If e.ColumnIndex = 3 Or e.ColumnIndex = 6 Then
                        skip(0) = True
                        MsgBox(DA.Columns(e.ColumnIndex).HeaderText & "输入有误，请检查后重输！")
                        s11(e)
                        Return
                    End If
                End If
                If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CStr(sv) Then
                    If IsNothing(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
                        cmdstr = String.Concat("update 物料数量 set ", DA.Columns(e.ColumnIndex).Name, "=NULL where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value))
                    Else
                        cmdstr = String.Concat(New String() {"update 物料数量 set ", DA.Columns(e.ColumnIndex).Name, "='", Replace(CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value), "'", "''"), "' where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)})
                    End If
                End If
            End If
            If cmdstr <> "" Then
                Try
                    cmd = New SqlCommand(cmdstr, cnct)
                    cnct.Open()
                    cmd.ExecuteNonQuery()
                    cnct.Close()
                    If CInt(DA.Rows(e.RowIndex).Cells(0).Value) <> 0 Then
                        If CR = 255 AndAlso CG = 255 AndAlso CB = 255 Then
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Red
                        ElseIf CR = 255 Then
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(0, 255, 0)
                        ElseIf CG <> 255 Then
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Red
                            DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.Black
                        Else
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Blue
                            DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.FromArgb(255, 255, 0)
                        End If
                    End If
                    If Not idt1.Contains(CInt(DA.Rows(e.RowIndex).Cells(0).Value)) Then idt1.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
                Catch ex As Exception
                    cnct.Close()
                    MsgBox(String.Concat("数据更新出错了..." & vbCrLf & "", ex.Message))
                    s11(e)
                    skip(0) = True
                    Return
                End Try
            End If
        End If
    End Sub
    Public Sub DA2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA2.CellEndEdit
        Dim an As Date, cmdstr1 As String, xn As Decimal, yn As Boolean = True, DA As DataGridView = DirectCast(sender, DataGridView)
        cmdstr = ""
        If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
            If suer <> 4 Then
                D3.Enabled = True
                D4.Enabled = True
            End If
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
            DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
        End If
        If IsNothing(DA.Rows(e.RowIndex).Cells(0).Value) Then
            If e.ColumnIndex = 1 Then
                If DA.NewRowIndex = e.RowIndex Then
                    ni = 2
                    ri = e.RowIndex
                    skip(0) = True
                    TM1.Enabled = True
                ElseIf IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then
                    DA.Rows(e.RowIndex).Cells(1).Value = Format(Now, "yyyy-MM-dd 07:59")
                Else
                    DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "07:59")
                End If
                Fcsb.s17(e.RowIndex)
            ElseIf e.ColumnIndex = 2 Then
                Fcsb.s17(e.RowIndex)
                Fcsb.s18(e.RowIndex)
            ElseIf e.ColumnIndex = 3 Then
                Fcsb.s17(e.RowIndex)
            End If
        ElseIf CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 Then
            skip(1) = False : skip(0) = False
            If e.ColumnIndex = 1 Then
                DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value))
                If Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), an) Then
                Else
                    skip(0) = True
                    MsgBox("日期输入有误，请检查后重输！")
                    s13(e)
                    Return
                End If
                If an <> CDate(sv) Then cmdstr = "update 储槽液位 set 日期='" & CStr(DA.Rows(e.RowIndex).Cells(1).Value) & "' where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Second, 30, CDate(DA.Rows(e.RowIndex).Cells(1).Value)), "yyyy-MM-dd HH:mm")
            ElseIf e.ColumnIndex = 3 Then
                DA.Rows(e.RowIndex).Cells(3).Value = Fcsb.s49(CStr(DA.Rows(e.RowIndex).Cells(3).Value), yn, xn)
                If IsNothing(DA.Rows(e.RowIndex).Cells(3).Value) Then
                    skip(0) = True
                    MsgBox("液位输入有误，请检查后重输！")
                    s13(e)
                    Return
                Else
                    If Not yn Then
                        skip(0) = True
                        MsgBox("液位输入有误，请检查后重输！")
                        s13(e)
                        Return
                    End If
                End If
                cmdstr = ""
                If xn <> CDec(sv) OrElse Not IsNumeric(DA.Rows(e.RowIndex).Cells(3).Value) Then cmdstr = "update 储槽液位 set 储槽液位=" & CStr(DA.Rows(e.RowIndex).Cells(3).Value) & " where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                DA.Rows(e.RowIndex).Cells(3).Value = CDec(Format(xn, "0.000"))
            ElseIf e.ColumnIndex = 2 Then
                If CStr(DA.Rows(e.RowIndex).Cells(2).Value) <> CStr(sv) Then
                    If IsNothing(DA.Rows(e.RowIndex).Cells(2).Value) Then
                        skip(0) = True
                        MsgBox("储槽名称有误，请检查后重输！")
                        s13(e)
                        Return
                    End If
                    cmdstr = "update 储槽液位 set 储槽名称='" & Replace(CStr(DA.Rows(e.RowIndex).Cells(2).Value), "'", "''") & "' where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                End If
            End If
            If cmdstr = "" Then Return
            Try
                cmd = New SqlCommand(cmdstr, cnct)
                cnct.Open()
                cmd.ExecuteNonQuery()
                cmdstr1 = "select dbo.储槽计算(@储槽名称,@液位,@时间)"
                cmd = New SqlCommand(cmdstr1, cnct)
                cmd.Parameters.Add(New SqlParameter("储槽名称", DA.Rows(e.RowIndex).Cells(2).Value))
                cmd.Parameters.Add(New SqlParameter("液位", DA.Rows(e.RowIndex).Cells(3).Value))
                cmd.Parameters.Add(New SqlParameter("时间", DA.Rows(e.RowIndex).Cells(1).Value))
                dr = cmd.ExecuteReader
                While dr.Read
                    DA.Rows(e.RowIndex).Cells(4).Value = IIf(IsDBNull(dr(0)), Nothing, dr(0))
                End While
                cnct.Close()
                Fcsb.s18(e.RowIndex)
                If CInt(DA.Rows(e.RowIndex).Cells(0).Value) <> 0 Then
                    If DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(200, 200, 200) Then
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(175, 175, 175)
                    ElseIf DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(175, 175, 175) Then
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(150, 150, 150)
                    Else
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(200, 200, 200)
                    End If
                End If
                If Not idt2.Contains(CInt(DA.Rows(e.RowIndex).Cells(0).Value)) Then idt2.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
            Catch ex As Exception
                cnct.Close()
                skip(0) = True
                MsgBox("数据更新出错了..." & vbCrLf & ex.Message)
                s13(e)
            End Try
        End If
    End Sub
    Public Sub DA9_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA9.CellEndEdit
        Dim an As Date, cmdstrb As String, DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
            D5.Enabled = True
            D6.Enabled = True
            D7.Enabled = True
            D8.Enabled = True
            For i = 1 To DA.Columns.Count
                DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
            DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
        End If
        If IsNothing(DA.Rows(e.RowIndex).Cells(0).Value) Then
            If e.ColumnIndex = 1 Then
                If DA.NewRowIndex = e.RowIndex Then
                    ni = 9
                    ri = e.RowIndex
                    skip(0) = True
                    TM1.Enabled = True
                ElseIf IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then
                    DA.Rows(e.RowIndex).Cells(1).Value = Format(Now, "yyyy-MM-dd")
                Else
                    DA.Rows(e.RowIndex).Cells(1).Value = Microsoft.VisualBasic.Left(Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), ""), 10)
                End If
            End If
        ElseIf CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 Then
            CR = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.R
            CG = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.G
            CB = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.B
            If CR < 136 OrElse CR = 255 AndAlso CG = 255 AndAlso CB = 0 Then CR = 255
            If CG < 136 OrElse CR = 255 AndAlso CG = 255 AndAlso CB = 0 Then CG = 255
            If CB < 136 OrElse CR = 255 AndAlso CG = 255 AndAlso CB = 0 Then CB = 255
            CR = CByte(CR - 51) : CG = CByte(CG - 51) : CB = CByte(CB - 51)
            cmdstr = "" : skip(1) = False : skip(0) = False
            If e.ColumnIndex = 1 Then
                DA.Rows(e.RowIndex).Cells(1).Value = Strings.Left(Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), ""), 10)
                If Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), an) Then
                Else
                    skip(0) = True
                    MsgBox("日期输入有误，请检查后重输！")
                    s21(e)
                    Return
                End If
            End If
            If e.ColumnIndex = 1 Then
                If CDate(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CDate(sv) Then
                    cmdstr = "update 报表备注 set "
                    cmdstrb = "备注日期"
                    cmdstr += cmdstrb & "=@" & cmdstrb & " where Id=@Id"
                End If
            Else
                If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CStr(sv) Then
                    cmdstr = "update 报表备注 set "
                    Select Case e.ColumnIndex
                        Case 2
                            cmdstrb = "报表名称"
                        Case 3
                            cmdstrb = "报表备注"
                    End Select
                    cmdstr += cmdstrb & "=@" & cmdstrb & " where Id=@Id"
                End If
            End If
            If cmdstr = "" Then
                Return
            Else
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.Parameters.Add(New SqlParameter("Id", DA.Rows(e.RowIndex).Cells(0).Value))
                cmd.Parameters.Add(New SqlParameter(cmdstrb, DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value))
            End If
            Try
                cnct.Open()
                cmd.ExecuteNonQuery()
                cnct.Close()
                DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(CR, CG, CB)
                If Not idt4.Contains(CInt(DA.Rows(e.RowIndex).Cells(0).Value)) Then idt4.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
            Catch ex As Exception
                cnct.Close()
                skip(0) = True
                MsgBox("备注更改失败，请重试！" & vbCrLf & ex.Message)
                s21(e)
                Return
            End Try
        End If
    End Sub
    Public Sub DA1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseClick
        Dim bn As String, i As Integer, DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 Then
            If e.ColumnIndex = 0 Then
                If e.Button = Windows.Forms.MouseButtons.Left Then
                    If CInt(DA.Rows(e.RowIndex).Cells(0).Value) < 0 Then
                        Fcsb.s45(DirectCast(sender, DataGridView), e.RowIndex, False)
                    ElseIf suer <> 5 AndAlso (CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 OrElse e.RowIndex = DA.NewRowIndex) Then
                        Form2.Show()
                        Form2.WindowState = FormWindowState.Normal
                        If CStr(DA.Rows(e.RowIndex).Cells(6).Value) = "入库" Then
                            Form2.DA1.Rows.Clear()
                            s60()
                            Form2.TC2.SelectedIndex = 0
                            Form2.CB1.Text = ""
                            Form2.T32.Text = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
                            Form2.CB2.Text = CStr(DA.Rows(e.RowIndex).Cells(3).Value)
                            Form2.B78_Click(B78, e)
                        ElseIf CStr(DA.Rows(e.RowIndex).Cells(8).Value) <> "" Then
                            Dim rg(,) As Object = Form2.rg
                            i = Form2.s4(CStr(DA.Rows(e.RowIndex).Cells(8).Value), 9)
                            If i > -1 Then
                                Form2.TC2.SelectedIndex = CInt(rg(0, i)) : bn = DirectCast(rg(2, i), TextBox).Text
                                If CStr(DA.Rows(e.RowIndex).Cells(2).Value) <> "" Then
                                    RemoveHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf Form2.T_TextChanged
                                    DirectCast(rg(2, i), TextBox).Text = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
                                    AddHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf Form2.T_TextChanged
                                    If CStr(rg(8, i)) <> DirectCast(rg(2, i), TextBox).Text Then
                                        Form2.s12(Form2.TC2.TabPages(CInt(rg(0, i))), DirectCast(Form2.rg(2, i), TextBox))
                                        DirectCast(rg(5, i), Button).Text = ""
                                    End If
                                End If
                                If DirectCast(rg(2, i), TextBox).Text <> bn Then
                                    Form2.s1(DirectCast(rg(2, i), TextBox), DirectCast(rg(3, i), Button), rg(10, i), DirectCast(rg(4, i), Button), DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), Form2.TC2.TabPages(CInt(rg(0, i))))
                                Else
                                    Fcsb.s40(DirectCast(rg(7, i), Dictionary(Of Control, String)), DirectCast(rg(2, i), TextBox).Parent, DirectCast(rg(2, i), TextBox))
                                End If
                            Else
                                Form2.TC2.SelectedIndex = 0
                            End If
                        Else
                            Form2.TC2.SelectedIndex = 0
                        End If
                        Form2.Activate()
                    End If
                ElseIf DA.Rows.Count > 1 AndAlso e.RowIndex > -1 AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) IsNot Nothing AndAlso e.Button = Windows.Forms.MouseButtons.Middle AndAlso Not sbl(3) Then
                    If DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                        DA.EndEdit()
                        If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                            Dim en As EventArgs
                            If B16.Text = "解锁表格" Then B16_Click(B16, en)
                            DA.Rows.Add()
                            For x = 1 To DA.Columns.Count - 1
                                DA.Rows(DA.Rows.Count - 2).Cells(x).Value = DA.Rows(e.RowIndex).Cells(x).Value
                            Next
                            RemoveHandler DA.RowValidating, AddressOf DA1_RowValidating
                            RemoveHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                            RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                            DA.Columns(1).Visible = True
                            DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(1)
                            DA.Rows(DA.Rows.Count - 2).ReadOnly = False
                            DA.BeginEdit(True)
                            DA.Rows(e.RowIndex).Cells(0).Selected = True
                            dttm = CStr(DA.Rows(e.RowIndex).Cells(1).Value)
                            AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                            AddHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                            AddHandler DA.RowValidating, AddressOf DA1_RowValidating
                        End If
                    End If
                ElseIf e.RowIndex > -1 AndAlso e.Button = Windows.Forms.MouseButtons.Right Then
                    Dim cmdstrn As Integer
                    If DA.Rows(e.RowIndex).Tag IsNot Nothing Then
                        cmdstrn = DirectCast(DA.Rows(e.RowIndex).Tag, Integer())(1)
                    ElseIf DA.Rows(e.RowIndex).Cells(0).Value IsNot Nothing Then
                        cmdstrn = CInt(DA.Rows(e.RowIndex).Cells(0).Value)
                    Else
                        Return
                    End If
                    Fcsb.s25(DA11, cmdstrn, "物料数量")
                    DA.ClearSelection() : DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True : TC1.SelectedIndex = 4
                End If
            End If
        Else
            If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                If e.Button = MouseButtons.Middle Then
                    Fcsb.s57(DA)
                ElseIf e.Button = MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                    DA.Columns.Item(e.ColumnIndex).Visible = False
                End If
            End If
        End If
    End Sub
    Private Sub DA1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseDoubleClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.ColumnIndex = 2 AndAlso DA.ReadOnly AndAlso e.RowIndex <> -1 AndAlso suer <> 4 AndAlso suer <> 5 Then
            If Not CStr(DA.Rows(e.RowIndex).Cells(2).Value) = "" Then
                For Each key As String In gd.Keys
                    If gd(key).Contains("|" & CStr(DA.Rows(e.RowIndex).Cells(8).Value) & "|") Then
                        If Not bn.ContainsKey(CStr(DA.Rows(e.RowIndex).Cells(2).Value)) Then bn.Add(CStr(DA.Rows(e.RowIndex).Cells(2).Value), key)
                        If Not LI19.Items.Contains(DA.Rows(e.RowIndex).Cells(2).Value) Then
                            LI19.Items.Add(DA.Rows(e.RowIndex).Cells(2).Value)
                            LI19.SetSelected(LI19.Items.Count - 1, True)
                        End If
                        DA.ClearSelection()
                        DA.Rows(e.RowIndex).Cells(0).Selected = True
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub DA_KeyDown(sender As Object, e As KeyEventArgs) Handles DA1.KeyDown, DA2.KeyDown, DA9.KeyDown, DA12.KeyDown
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.KeyCode = Keys.Escape Then
            Try
                If IsNothing(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) Then
                    DA.Rows.RemoveAt(DA.Rows.Count - 2)
                    For Each col As DataGridViewColumn In DA.Columns
                        col.SortMode = DataGridViewColumnSortMode.Automatic
                    Next
                    If suer <> 4 Then
                        If DA Is DA1 Then
                            D1.Enabled = True
                            D2.Enabled = True
                        ElseIf DA Is DA2 Then
                            D3.Enabled = True
                            D4.Enabled = True
                        ElseIf DA Is DA9 Then
                            D5.Enabled = True
                            D6.Enabled = True
                            D7.Enabled = True
                            D8.Enabled = True
                        ElseIf B106.Text = "储槽表格" Then
                            DA.Rows(DA.NewRowIndex).Cells(6).ReadOnly = True
                        End If
                    End If
                    skip(0) = False
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub DA10_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA10.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex = -1 Then
            If e.Button = MouseButtons.Middle Then
                DA.RowHeadersWidth = 41
                For i = 1 To DA.Columns.Count
                    DA.Columns.Item(i - 1).Visible = True
                    DA.Columns.Item(i - 1).Width = CInt(700 / DA.Columns.Count)
                Next
            ElseIf e.Button = MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        ElseIf e.ColumnIndex = -1 AndAlso e.Button = MouseButtons.Middle Then
            Dim rowindexs As New List(Of Integer)
            If DA.Rows(e.RowIndex).Selected Then
                For Each row As DataGridViewRow In DA.SelectedRows
                    rowindexs.Add(row.Index)
                Next
            ElseIf Not DA.Rows(e.RowIndex).IsNewRow Then
                rowindexs.Add(e.RowIndex)
            End If
            rowindexs.Sort()
            For Each index As Integer In rowindexs
                DA.Rows.Add()
                For Each cell As DataGridViewCell In DA.Rows(index).Cells
                    DA.Rows(DA.Rows.Count - 2).Cells(cell.ColumnIndex).Value = cell.Value
                Next
                DA.Rows(DA.Rows.Count - 2).Cells(2).Tag = DA.Rows(index).Cells(2).Tag
            Next
        End If
    End Sub
    Private Sub DA10_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA10.CellMouseDown
        Dim DA As DataGridView = DirectCast(sender, DataGridView), cel As DataGridViewCell
        If e.Button = MouseButtons.Right AndAlso e.ColumnIndex = 2 Then
            If DA.Columns.Count = 4 Then
                For i = 0 To 3
                    DA.Rows(DA.NewRowIndex).Cells(i).Selected = False
                Next
            End If
            For Each cell As DataGridViewCell In DA.SelectedCells
                If cell.ColumnIndex = 2 Then
                    cel = cell
                    Exit For
                End If
            Next
            If DA.Columns.Count = 4 AndAlso e.RowIndex > -1 AndAlso e.RowIndex < DA.NewRowIndex AndAlso e.ColumnIndex > -1 Then
                If DA.SelectedCells.Count >= 2 OrElse cel IsNot Nothing AndAlso ctlbl Then
                    For Each cell As DataGridViewCell In DA.SelectedCells
                        If cell.ColumnIndex = 2 Then
                            If CStr(cel.Value) <> CStr(cell.Value) Then
                                DA.ClearSelection()
                                DA.Rows(e.RowIndex).Cells(2).Selected = True
                                Return
                            End If
                        End If
                    Next
                ElseIf e.ColumnIndex = 2 Then
                    DA.ClearSelection()
                    DA.Rows(e.RowIndex).Cells(2).Selected = True
                    Return
                End If
                For Each cell As DataGridViewCell In DA.SelectedCells
                    cell.Selected = cell.ColumnIndex = 2
                Next
                If DA.SelectedCells.Count = 0 AndAlso TypeOf ActiveControl IsNot DataGridViewTextBoxEditingControl Then DA.Rows(e.RowIndex).Cells(2).Selected = True
            End If
        End If
    End Sub
    Public Sub DAN_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA3.CellMouseUp, DA5.CellMouseUp, DA6.CellMouseUp, DA9.CellMouseUp, DA12.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.SelectedCells.Count = 1 AndAlso e.RowIndex > -1 AndAlso e.ColumnIndex > -1 Then
            DA.Columns(DA.SelectedCells(0).ColumnIndex).Visible = True
            If e.Button = MouseButtons.Left Then
                DA.ClearSelection()
                DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True
                DA.Columns(e.ColumnIndex).Visible = True
                If DA IsNot DA9 AndAlso DA IsNot DA12 Then DA.CurrentCell = DA.Rows(e.RowIndex).Cells(e.ColumnIndex)
                DA.BeginEdit(True)
            End If
        End If
    End Sub
    Public Sub DA12_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA12.CellMouseClick
        Dim fc As Color, i As Integer, em As EventArgs, DA As DataGridView = DirectCast(sender, DataGridView)
        If e.ColumnIndex = 0 AndAlso e.RowIndex > -1 AndAlso B85.Enabled Then
            If e.Button = MouseButtons.Left Then
                If B106.Text = "储槽表格" Then
                    fc = DA.Rows(e.RowIndex).Cells(0).Style.ForeColor
                    Try
                        cnct.Open()
                        cmdstr = String.Concat("update 工序类型 set 可用性=", CStr(IIf(fc = Color.Green, 0, 1)), " where id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value))
                        cmd = New SqlCommand(cmdstr, cnct)
                        cmd.ExecuteNonQuery()
                        cnct.Close()
                        whbl = True
                        If fc = Color.Red Then
                            DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.Green
                        ElseIf fc = Color.Green Then
                            DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.Red
                        End If
                        RemoveHandler DA.SelectionChanged, AddressOf DA12_SelectionChanged
                        RemoveHandler DA.RowValidating, AddressOf DA12_RowValidating
                        DA.ClearSelection()
                        AddHandler DA.RowValidating, AddressOf DA12_RowValidating
                        AddHandler DA.SelectionChanged, AddressOf DA12_SelectionChanged
                    Catch exception As Exception
                        cnct.Close()
                    End Try
                End If
            ElseIf e.Button = MouseButtons.Right Then
                Dim str As String = CStr(DA.Rows(e.RowIndex).Cells(3).Value)
                If str = "消耗" OrElse str = "产出" Then
                    fc = DA.Rows(e.RowIndex).Cells(0).Style.BackColor
                    Try
                        cnct.Open()
                        cmd = New SqlCommand("单耗变更", cnct) With {.CommandType = CommandType.StoredProcedure}
                        cmd.Parameters.Add(New SqlParameter("R", fc.R))
                        cmd.Parameters.Add(New SqlParameter("G", fc.G))
                        cmd.Parameters.Add(New SqlParameter("B", fc.B))
                        cmd.Parameters.Add(New SqlParameter("物料名称", CStr(DA.Rows(e.RowIndex).Cells(1).Value)))
                        cmd.Parameters.Add(New SqlParameter("物料类型", str))
                        cmd.Parameters.Add(New SqlParameter("操作工序", CStr(DA.Rows(e.RowIndex).Cells(2).Value)))
                        dr = cmd.ExecuteReader()
                        While dr.Read()
                            For i = 0 To DA.Rows.Count - 1
                                If CStr(DA.Rows(i).Cells(3).Value) = str AndAlso CStr(DA.Rows(i).Cells(1).Value) = CStr(DA.Rows(e.RowIndex).Cells(1).Value) AndAlso CStr(DA.Rows(i).Cells(2).Value) = CStr(DA.Rows(e.RowIndex).Cells(2).Value) Then
                                    DA.Rows(i).Cells(0).Style.BackColor = Color.FromArgb(CInt(dr(0)), CInt(dr(1)), CInt(dr(2)))
                                End If
                            Next
                        End While
                        cnct.Close()
                        whbl = True
                    Catch ex As Exception
                        cnct.Close()
                    End Try
                    s25(DA, e.RowIndex)
                Else
                    Try
                        cnct.Open()
                        i = New SqlCommand(String.Concat("update 工序类型 set 单耗标记=NULL,单耗预估值=NULL where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)), cnct).ExecuteNonQuery()
                        cnct.Close()
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(255, 255, 255)
                        DA.Rows(e.RowIndex).Cells(6).Value = Nothing
                    Catch ex As Exception
                        cnct.Close()
                    End Try
                End If
                DA.ClearSelection()
            ElseIf DA.Rows.Count > 1 AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) IsNot Nothing AndAlso DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                RemoveHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
                DA.EndEdit()
                If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                    If B106.Text = "解锁表格" Then
                        B106_Click(B106, em)
                    ElseIf B90.Text = "解锁表格" Then
                        B90_Click(B90, em)
                    End If
                    DA.Rows.Add()
                    For i = 1 To DA.Columns.Count - 1
                        DA.Rows(DA.Rows.Count - 2).Cells(i).Value = DA.Rows(e.RowIndex).Cells(i).Value
                    Next
                    RemoveHandler DA.RowValidating, AddressOf DA12_RowValidating
                    RemoveHandler DA.SelectionChanged, AddressOf DA12_SelectionChanged
                    DA.Columns(1).Visible = True
                    DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(1)
                    DA.BeginEdit(True)
                    AddHandler DA.SelectionChanged, AddressOf DA12_SelectionChanged
                    AddHandler DA.RowValidating, AddressOf DA12_RowValidating
                    If B106.Text = "储槽表格" Then DA.Rows(DA.Rows.Count - 2).Cells(6).ReadOnly = True
                End If
                AddHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
            End If
        End If
    End Sub
    Private Sub DA11_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA11.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        Try
            s4(DA, e.ColumnIndex, T54, T55, T56, T57, T58)
        Catch ex As Exception
        End Try
        Dim a(1) As String
        Dim j As Integer
        Dim rst As Date
        For i = 0 To DA.SelectedCells.Count - 1
            If Date.TryParse(CStr(DA.SelectedCells.Item(i).Value), rst) Then
                a(j) = CStr(DA.SelectedCells.Item(i).Value)
                j += 1
                If j = 2 Then Exit For
            End If
        Next
        T61.Text = Fcsb.s43(a(1), a(0))
    End Sub
    Public Sub RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs)
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        TextRenderer.DrawText(e.Graphics, CStr(e.RowIndex + 1), New System.Drawing.Font("Times New Roman", 9), New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y + 4, DA.RowHeadersWidth - 4, e.RowBounds.Height), DA.RowHeadersDefaultCellStyle.ForeColor, Color.Transparent, TextFormatFlags.HorizontalCenter)
        For Each darc As DataGridViewCell In dacl(DA)
            If darc.RowIndex Mod 2 = 0 AndAlso darc.Style.BackColor = DA.AlternatingRowsDefaultCellStyle.BackColor Then
                darc.Style.BackColor = DA.RowsDefaultCellStyle.BackColor
            ElseIf darc.RowIndex Mod 2 = 1 AndAlso darc.Style.BackColor = DA.RowsDefaultCellStyle.BackColor Then
                darc.Style.BackColor = DA.AlternatingRowsDefaultCellStyle.BackColor
            End If
        Next
    End Sub
    Public Sub DA1_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA1.RowValidating
        Dim dt As Date, num, num1 As Decimal, msgr As MsgBoxResult, DA As DataGridView = DirectCast(sender, DataGridView), strm As String = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            If CStr(DA.Rows(e.RowIndex).Cells(1).Value) = "" Then DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Day, -1, Now), "yyyy-MM-dd 08:00")
            DA.EndEdit()
            Dim str(4) As String
            Dim flag As Boolean = True
            Dim flag1 As Boolean = True
            DA.Rows(e.RowIndex).Cells(1).Value = Replace(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "'", "")
            str(3) = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
            str(0) = Replace(CStr(DA.Rows(e.RowIndex).Cells(3).Value), "'", "''")
            str(1) = Replace(CStr(DA.Rows(e.RowIndex).Cells(6).Value), "'", "''")
            str(2) = Replace(CStr(DA.Rows(e.RowIndex).Cells(8).Value), "'", "''")
            str(4) = Replace(CStr(DA.Rows(e.RowIndex).Cells(9).Value), "'", "''")
            DA.Rows(e.RowIndex).Cells(4).Value = Replace(CStr(DA.Rows(e.RowIndex).Cells(4).Value), "'", "")
            DA.Rows(e.RowIndex).Cells(5).Value = Replace(CStr(DA.Rows(e.RowIndex).Cells(5).Value), "'", "")
            DA.Rows(e.RowIndex).Cells(4).Value = Fcsb.s49(CStr(DA.Rows(e.RowIndex).Cells(4).Value), flag, num)
            DA.Rows(e.RowIndex).Cells(5).Value = Fcsb.s49(CStr(DA.Rows(e.RowIndex).Cells(5).Value), flag1, num1)
            Dim str2 As String = "insert into 物料数量 values("
            DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "08:00")
            Dim str5 As String = "'" + str(0) + "'"
            Dim str8 As String = "'" + str(1) + "'"
            Dim str10 As String = "'" + str(2) + "'"
            Dim str11 As String = "'" + str(4) + "'"
            Dim str6 As String = CStr(DA.Rows(e.RowIndex).Cells(4).Value)
            Dim str7 As String = CStr(DA.Rows(e.RowIndex).Cells(5).Value)
            Dim str4 As String = "'" + CStr(DA.Rows(e.RowIndex).Cells(1).Value) + "'"
            Dim str9 As String = "'" + Replace(CStr(DA.Rows(e.RowIndex).Cells(7).Value), "'", "''") + "'"
            If Not Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), dt) Then
                Fcsb.s16(DA, 1, e)
                Return
            End If
            Do While CH32.Checked AndAlso Fcsb.s10(strm, True) = 0 AndAlso CStr(DA.Rows(e.RowIndex).Cells(2).Value) <> ""
                skip(0) = True
                msgr = MsgBox("批号格式不正确！", MsgBoxStyle.AbortRetryIgnore)
                If msgr = MsgBoxResult.Abort Then
                    e.Cancel = True
                    DA.Columns(2).Visible = True
                    DA.CurrentCell = DA.Rows(e.RowIndex).Cells(2)
                    DA.BeginEdit(False)
                    Return
                ElseIf msgr = MsgBoxResult.Ignore Then
                    skip(0) = False
                    RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    DA.EndEdit()
                    DA.Select()
                    AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    Exit Do
                End If
            Loop
            Dim str12 As String = String.Concat("'", Replace(str(3), "'", "''"), "'")
            If CStr(DA.Rows(e.RowIndex).Cells(3).Value) = "" Then
                Fcsb.s16(DA, 3, e)
                Return
            End If
            If Not flag OrElse CStr(DA.Rows(e.RowIndex).Cells(4).Value) = "" Then
                Fcsb.s16(DA, 4, e)
                Return
            End If
            If num1 <= 0 AndAlso CStr(DA.Rows(e.RowIndex).Cells(5).Value) <> "" Then
                Fcsb.s16(DA, 5, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(6).Value) = "" Then
                Fcsb.s16(DA, 6, e)
                Return
            End If
            Do While CH33.Checked AndAlso Fcsb.s15(str)
                skip(0) = True
                msgr = MsgBox("输入的条目不匹配！", MsgBoxStyle.AbortRetryIgnore)
                If msgr = MsgBoxResult.Abort Then
                    e.Cancel = True
                    DA.Columns(2).Visible = True
                    DA.CurrentCell = DA.Rows(e.RowIndex).Cells(2)
                    DA.BeginEdit(False)
                    Return
                ElseIf msgr = MsgBoxResult.Ignore Then
                    skip(0) = False
                    RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    DA.EndEdit()
                    DA.Select()
                    AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                    Exit Do
                End If
            Loop
            If str12 = "''" Then str12 = "NULL"
            If str7 = "" Then str7 = "NULL"
            If str9 = "''" Then str9 = "NULL"
            If str10 = "''" Then str10 = "NULL"
            If str11 = "''" Then str11 = "NULL"
            str2 = String.Concat(str2, str4, ",", str12, ",", str5, ",", str6, ",", str7, ",", str8, ",", str9, ",", str10, ",", str11, ")")
            If Not Fcsb.s9(DA.Rows.Count - 2, DA, "物料数量", str2) Then
                idt1.Add(CInt(DA1.Rows(e.RowIndex).Cells(0).Value))
                DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.DarkViolet
                DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.White
                DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Second, 30, CDate(DA.Rows(e.RowIndex).Cells(1).Value)), "yyyy-MM-dd HH:mm")
            End If
            skip(0) = False : D1.Enabled = True : D2.Enabled = True
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
            DA1.Rows(e.RowIndex).Cells(4).Value = CDec(Format(num, "0.000"))
            If CStr(DA.Rows(e.RowIndex).Cells(5).Value) <> "" Then
                DA.Rows(e.RowIndex).Cells(5).Value = CDec(Format(num1, "0.00"))
            Else
                DA.Rows(e.RowIndex).Cells(5).Value = Nothing
            End If
        End If
    End Sub
    Public Sub DA2_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA2.RowValidating
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            Dim an As Date
            Dim xn As Decimal = -1
            Dim lm(1), cmdstr1 As String
            If CStr(DA.Rows(e.RowIndex).Cells(1).Value) = "" Then DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Day, -1, Now()), "yyyy-MM-dd 07:59")
            DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "07:59")
            DA.EndEdit()
            cmdstr1 = "'" + CStr(DA.Rows(e.RowIndex).Cells(1).Value) + "'"
            Dim cmdstr2 = "'" & Replace(CStr(DA.Rows(e.RowIndex).Cells(2).Value), "'", "''") & "'"
            If Not Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), an) Then
                Fcsb.s16(DA, 1, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(2).Value) = "" Then
                Fcsb.s16(DA, 2, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(3).Value) <> "" Then
                DA.Rows(e.RowIndex).Cells(3).Value = Fcsb.s49(CStr(DA.Rows(e.RowIndex).Cells(3).Value), True, xn)
                If xn < 0 Then
                    Fcsb.s16(DA, 3, e)
                    Return
                End If
            Else
                Fcsb.s16(DA, 3, e)
                Return
            End If
            If cmdstr2 = "''" Then cmdstr2 = "NULL"
            cmdstr = "insert into 储槽液位 values("
            cmdstr += cmdstr1 & "," : cmdstr += cmdstr2 & ","
            cmdstr += CStr(DA.Rows(e.RowIndex).Cells(3).Value) & ")"
            If Fcsb.s9(DA.Rows.Count - 2, DA, "储槽液位", cmdstr) Then
                DA.Rows(e.RowIndex).Cells(3).Value = CDec(Format(xn, "0.000"))
                DA.Rows(e.RowIndex).Cells(4).Value = Nothing
                DA.Rows(e.RowIndex).Cells(5).Value = Nothing
                DA.Rows(e.RowIndex).Cells(6).Value = Nothing
            Else
                DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Pink
                idt2.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
                DA.Rows(e.RowIndex).Cells(3).Value = CDec(Format(xn, "0.000"))
                DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Second, 30, CDate(DA.Rows(e.RowIndex).Cells(1).Value)), "yyyy-MM-dd HH:mm")
                Fcsb.s17(e.RowIndex) : Fcsb.s18(e.RowIndex)
            End If
            skip(0) = False : D3.Enabled = True : D4.Enabled = True
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
    End Sub
    Public Sub DA9_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA9.RowValidating
        Dim an As Date, DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            DA.EndEdit()
            If IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Day, -1, Now), "yyyy-MM-dd")
            cmdstr = "Insert into 报表备注(备注日期,报表名称,报表备注) values(@备注日期,@报表名称,@报表备注)"
            If Not Date.TryParse(Replace(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "：", ":"), an) Then
                Fcsb.s16(DA, 1, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(2).Value) = "" Then
                Fcsb.s16(DA, 2, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(3).Value) = "" Then
                Fcsb.s16(DA, 3, e)
                Return
            End If
            Dim sqlprmt(2) As SqlParameter
            Try
                sqlprmt(0) = New SqlParameter("备注日期", CDate(DA.Rows(e.RowIndex).Cells(1).Value))
            Catch ex As Exception
            End Try
            sqlprmt(1) = New SqlParameter("报表名称", CStr(DA.Rows(e.RowIndex).Cells(2).Value))
            sqlprmt(2) = New SqlParameter("报表备注", CStr(DA.Rows(e.RowIndex).Cells(3).Value))
            If Not Fcsb.s9(DA.Rows.Count - 2, DA, "报表备注", cmdstr, sqlprmt) Then
                idt4.Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
                DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(255, 204, 204, 204)
                DA.Rows(e.RowIndex).Cells(1).Value = Format(an, "yyyy-MM-dd")
            End If
            skip(0) = False : D5.Enabled = True : D6.Enabled = True : D7.Enabled = True : D8.Enabled = True
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
    End Sub
    Public Sub TM1_Tick(sender As Object, e As EventArgs) Handles TM1.Tick
        DirectCast(sender, Timer).Enabled = False
        DirectCast(sender, Timer).Interval = 1
        If ctbl Then
            RemoveHandler B14.Click, AddressOf B14_Click
            Fcsb.s39(suer = 4)
            AddHandler B14.Click, AddressOf B14_Click
        ElseIf ckbl Then
            RemoveHandler B45.Click, AddressOf B45_Click
            s37(False)
            AddHandler B45.Click, AddressOf B45_Click
        ElseIf ttbl Then
            RemoveHandler B24.Click, AddressOf B24_Click
            s38(suer = 4)
            AddHandler B24.Click, AddressOf B24_Click
        ElseIf ccbl Then
            RemoveHandler B52.Click, AddressOf B52_Click
            s39(sbl(3))
            AddHandler B52.Click, AddressOf B52_Click
        ElseIf bcbl Then
            RemoveHandler B103.Click, AddressOf B103_Click
            s40(True)
            AddHandler B103.Click, AddressOf B103_Click
        ElseIf ccbl2 Then
            RemoveHandler B78.Click, AddressOf B78_Click
            Fcsb.s42(suer = 4)
            AddHandler B78.Click, AddressOf B78_Click
        ElseIf b105bl Then
            RemoveHandler B105.Click, AddressOf B105_Click
            s46(False)
            AddHandler B105.Click, AddressOf B105_Click
        ElseIf ni <> 0 Then
            Select Case ni
                Case 1
                    RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
                    RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
                    DA1.Rows.Add()
                    AddHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
                    DA1.Rows(ri).Cells(1).Value = Format(Now, "yyyy-MM-dd") + " 08:00"
                    DA1.Columns(2).Visible = True
                    DA1.CurrentCell = DA1.Rows(ri).Cells(2)
                    AddHandler DA1.RowValidating, AddressOf DA1_RowValidating
                    If DA1.Rows.Count > 2 AndAlso IsNothing(DA1.Rows(DA1.Rows.Count - 3).Cells(0).Value) Then
                        DA1.Rows.RemoveAt(DA1.Rows.Count - 2)
                        DA1.Rows(DA1.Rows.Count - 2).Cells(1).Value = dttm
                        DA1.CurrentCell = DA1.Rows(ri).Cells(1)
                    End If
                    DA1.BeginEdit(True)
                Case 2
                    RemoveHandler DA2.RowValidating, AddressOf DA2_RowValidating
                    RemoveHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
                    DA2.Rows.Add()
                    AddHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
                    DA2.Rows(ri).Cells(1).Value = Format(Now, "yyyy-MM-dd") + " 07:59"
                    DA2.Columns(2).Visible = True
                    DA2.CurrentCell = DA2.Rows(ri).Cells(2)
                    AddHandler DA2.RowValidating, AddressOf DA2_RowValidating
                    If DA2.Rows.Count > 2 AndAlso IsNothing(DA2.Rows(DA2.Rows.Count - 3).Cells(0).Value) Then
                        DA2.Rows.RemoveAt(DA2.Rows.Count - 2)
                        DA2.Rows(DA2.Rows.Count - 2).Cells(1).Value = dttm
                        DA2.CurrentCell = DA2.Rows(ri).Cells(1)
                    End If
                    DA2.BeginEdit(True)
                Case 3
                    DA3.Rows.Add()
                    DA3.Rows(ri).Cells(0).Value = Format(Now, "yyyy-MM-dd") + " 07:59"
                    DA3.ClearSelection()
                Case 5
                    DA5.Rows.Add()
                    DA5.Rows(ri).Cells(0).Value = Format(Now, "yyyy-MM-dd") + " 08:00"
                    DA5.ClearSelection()
                Case 6
                    For i = 6 To DA6.Columns.Count - 1
                        DA6.Columns.RemoveAt(6)
                    Next
                    cmdstr = "select distinct 工序类型.物料名称,物料特性.Id from 物料特性,工序类型 where 操作工序 is NULL and 物料特性.可用性=1 and 工序类型.可用性=1 and 物料类型='" & CStr(DA6.Rows(0).Cells(5).Value) & "' and 物料特性.物料名称=工序类型.物料名称 order by 物料特性.Id"
                    Try
                        cnct.Open()
                        cmd = New SqlCommand(cmdstr, cnct)
                        dr = cmd.ExecuteReader
                        While dr.Read
                            DA6.Columns.Add(CStr(dr(0)), CStr(dr(0)))
                        End While
                        cnct.Close()
                    Catch ex As Exception
                        cnct.Close()
                        MsgBox("载入物料出错！" & vbCrLf & ex.Message)
                    End Try
                    AddHandler DA6.CellEndEdit, AddressOf DA_CellEndEdit
                Case 9
                    RemoveHandler DA9.RowValidating, AddressOf DA9_RowValidating
                    RemoveHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
                    DA9.Rows.Add()
                    AddHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
                    DA9.Rows(ri).Cells(1).Value = Format(Now, "yyyy-MM-dd")
                    DA9.Columns(2).Visible = True
                    DA9.CurrentCell = DA9.Rows(ri).Cells(2)
                    AddHandler DA9.RowValidating, AddressOf DA9_RowValidating
                    DA9.BeginEdit(True)
                Case 10
                    RemoveHandler DA10.CellEndEdit, AddressOf DA10_CellEndEdit
                    DA10.Rows.Add()
                    AddHandler DA10.CellEndEdit, AddressOf DA10_CellEndEdit
                    DA10.Rows(ri).Cells(0).Value = Format(Now, "yyyy-MM-dd")
                    DA10.Columns(1).Visible = True
                    DA10.CurrentCell = DA10.Rows(ri).Cells(1)
                    DA10.BeginEdit(True)
                Case 12
                    RemoveHandler DA12.RowValidating, AddressOf DA12_RowValidating
                    RemoveHandler DA12.CellEndEdit, AddressOf DA12_CellEndEdit
                    DA12.Rows.Add()
                    AddHandler DA12.CellEndEdit, AddressOf DA12_CellEndEdit
                    DA12.Rows(ri).Cells(1).Value = Format(Now, "yyyy-MM-dd 07:59")
                    DA12.Columns(2).Visible = True
                    DA12.CurrentCell = DA12.Rows(ri).Cells(2)
                    AddHandler DA12.RowValidating, AddressOf DA12_RowValidating
                    If DA12.Rows.Count > 2 AndAlso IsNothing(DA12.Rows(DA12.Rows.Count - 3).Cells(0).Value) Then
                        DA12.Rows.RemoveAt(DA12.Rows.Count - 2)
                        DA12.Rows(DA12.Rows.Count - 2).Cells(1).Value = dttm
                        DA12.CurrentCell = DA12.Rows(ri).Cells(1)
                    End If
                    DA12.BeginEdit(True)
            End Select
            ni = 0
        ElseIf CBool(lbl(L125)(0)) Then
            RemoveHandler L125.Click, AddressOf LB_Click
            s24(False)
            AddHandler L125.Click, AddressOf LB_Click
        ElseIf CBool(lbl(L126)(0)) Then
            RemoveHandler L126.Click, AddressOf L_Click
            Fcsb.s30(False, DirectCast(L126, Object))
            AddHandler L126.Click, AddressOf L_Click
        ElseIf CBool(lbl(L127)(0)) Then
            RemoveHandler L127.Click, AddressOf L_Click
            Fcsb.s30(False, DirectCast(L127, Object))
            AddHandler L127.Click, AddressOf L_Click
        ElseIf CBool(lbl(L128)(0)) Then
            RemoveHandler L128.Click, AddressOf L_Click
            Fcsb.s30(False, DirectCast(L128, Object))
            AddHandler L128.Click, AddressOf L_Click
        ElseIf lbl126 Then
            RemoveHandler Form2.L126.Click, AddressOf Form2.L126_Click
            Fcsb.s30(False, DirectCast(Form2.L126, Object), False)
            AddHandler Form2.L126.Click, AddressOf Form2.L126_Click
        ElseIf CBool(lbl(L130)(0)) Then
            RemoveHandler L130.Click, AddressOf LB_Click
            s27(False)
            AddHandler L130.Click, AddressOf LB_Click
        ElseIf L124bl Then
            RemoveHandler L124.MouseClick, AddressOf L124_MouseClick
            s22(False, ex)
            AddHandler L124.MouseClick, AddressOf L124_MouseClick
        End If
    End Sub
    Private Sub B24_Click(sender As Object, e As EventArgs) Handles B24.Click
        If Not ttbl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            ttbl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        s38(True)
    End Sub
    Private Sub B28_Click(sender As Object, e As EventArgs) Handles B28.Click
        Dim n As Integer
        If Integer.TryParse(T2.Text, n) Then
            s9(n)
            DA2.ClearSelection()
        Else
            MsgBox("请正确输入序号！")
        End If
    End Sub
    Private Sub B30_Click(sender As Object, e As EventArgs) Handles B30.Click
        Fcsb.s8(DA3, False)
    End Sub
    Private Sub B34_Click(sender As Object, e As EventArgs) Handles B34.Click
        Fcsb.s8(DA5, False)
    End Sub
    Private Sub B18_Click(sender As Object, e As EventArgs) Handles B18.Click
        DA6.Rows.Clear()
        For i = 6 To DA6.Columns.Count - 1
            DA6.Columns.RemoveAt(6)
        Next
        DA6.Rows.Add()
    End Sub
    Private Sub B31_Click(sender As Object, e As EventArgs) Handles B31.Click
        Dim bl As Boolean
        s42(bl)
        RemoveHandler DA3.SelectionChanged, AddressOf DA3_SelectionChanged
        If bl Then DA3.Rows.Clear()
        AddHandler DA3.SelectionChanged, AddressOf DA3_SelectionChanged
    End Sub
    Private Sub B35_Click(sender As Object, e As EventArgs) Handles B35.Click
        Dim bl As Boolean
        s43(bl)
        If bl Then DA5.Rows.Clear()
    End Sub
    Private Sub B29_Click(sender As Object, e As EventArgs) Handles B29.Click
        Dim bl As Boolean
        s44(bl)
        If bl Then
            DA6.Rows.Clear()
            For i = 6 To DA6.Columns.Count - 1
                DA6.Columns.RemoveAt(6)
            Next
            DA6.Rows.Add()
        End If
    End Sub
    Private Sub B48_Click(sender As Object, e As EventArgs) Handles B48.Click
        Dim a, b, c, d, f As New DataTable, k, k1, k2 As Integer, bl As Boolean = True
        If CL1.Items.Count = 1 OrElse CL3.Items.Count = 1 Then Return
        If CL1.CheckedItems.Count = 0 AndAlso CL3.CheckedItems.Count = 0 Then Return
        c.Columns.Add("物料名称")
        For j = 1 To CL1.Items.Count - 1
            If CL1.GetItemChecked(j) Then bl = False : Exit For
        Next
        If bl Then
            For i = 0 To CL1.Items.Count - 1
                c.Rows.Add(CL1.Items.Item(i))
            Next
        Else
            For i = 0 To CL1.CheckedItems.Count - 1
                c.Rows.Add(CL1.CheckedItems.Item(i))
            Next
        End If
        bl = True : If CStr(c.Rows(0).Item(0)) = "全部" Then c.Rows.RemoveAt(0)
        d.Columns.Add("物料名称")
        For j = 1 To CL3.Items.Count - 1
            If CL3.GetItemChecked(j) Then bl = False : Exit For
        Next
        If bl Then
            For i = 0 To CL3.Items.Count - 1
                d.Rows.Add(CL3.Items.Item(i))
            Next
        Else
            For i = 0 To CL3.CheckedItems.Count - 1
                d.Rows.Add(CL3.CheckedItems.Item(i))
            Next
        End If
        If CStr(d.Rows(0)(0)).Contains("全部") Then d.Rows.RemoveAt(0)
        f.Columns.Add("时间", Type.GetType("System.DateTime"))
        k1 = CInt(Math.Min(CDate(D7.Text).ToOADate, CDate(D8.Text).ToOADate))
        k2 = CInt(Math.Max(CDate(D7.Text).ToOADate, CDate(D8.Text).ToOADate))
        If D8.Checked Then
            Dim stp As Integer
            If L42.Text = "~" Then
                stp = 1
            Else
                stp = k2 - k1
            End If
            For k = k1 To k2 Step stp
                f.Rows.Add(Date.FromOADate(k))
                If k1 = k2 Then
                    If rb Then s19(Fcsb.s23(), c, d, f)
                    Exit For
                End If
            Next
        Else
            f.Rows.Add(CDate(D7.Text))
        End If
        If rb Then
            s19(Fcsb.s23(), c, d, f)
        Else
            a.Columns.Add("物料名称")
            For x = 1 To CO10.Items.Count
                If UCase(CO10.Text).Contains(CStr(CO10.Items(x - 1))) Then
                    If CO10.Items(x - 1) Is " " Then
                        a.Rows.Add(DBNull.Value)
                    Else
                        a.Rows.Add(CStr(CO10.Items(x - 1)))
                    End If
                End If
            Next
            b.Columns.Add("物料名称")
            For x = 1 To CO11.Items.Count
                If UCase(CO11.Text).Contains(CStr(CO11.Items(x - 1))) Then
                    If CO11.Items(x - 1) Is " " Then
                        b.Rows.Add(DBNull.Value)
                    Else
                        b.Rows.Add(CStr(CO11.Items(x - 1)))
                    End If
                End If
            Next
            Try
                cnct.Open()
                cmd.CommandTimeout = 0
                cmd = New SqlCommand("时期统计", cnct)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("物料名称", c))
                cmd.Parameters.Add(New SqlParameter("物料类型", d))
                cmd.Parameters.Add(New SqlParameter("班别班组", a))
                cmd.Parameters.Add(New SqlParameter("反应釜号", b))
                cmd.Parameters.Add(New SqlParameter("起始时间", Date.FromOADate(k1)))
                cmd.Parameters.Add(New SqlParameter("结束时间", Date.FromOADate(k2)))
                dr = cmd.ExecuteReader
                While dr.Read
                    DA10.Rows.Add()
                    RemoveHandler DA10.CellValueChanged, AddressOf DA10_CellValueChanged
                    DA10.Rows(DA10.Rows.Count - 2).Cells(0).Value = Format(CDate(dr(0)), "yyyy-MM-dd")
                    DA10.Rows(DA10.Rows.Count - 2).Cells(1).Value = Format(CDate(dr(1)), "yyyy-MM-dd")
                    DA10.Rows(DA10.Rows.Count - 2).Cells(2).Value = CStr(dr(2))
                    DA10.Rows(DA10.Rows.Count - 2).Cells(4).Value = IIf(IsDBNull(dr(4)), "NULL", dr(4))
                    DA10.Rows(DA10.Rows.Count - 2).Cells(5).Value = IIf(IsDBNull(dr(5)), "NULL", dr(5))
                    DA10.Rows(DA10.Rows.Count - 2).Cells(6).Value = IIf(IsDBNull(dr(6)), "NULL", dr(6))
                    AddHandler DA10.CellValueChanged, AddressOf DA10_CellValueChanged
                    If IsDBNull(dr(3)) Then
                        DA10.Rows(DA10.Rows.Count - 2).Cells(3).Value = Nothing
                    Else
                        DA10.Rows(DA10.Rows.Count - 2).Cells(3).Value = dr(3)
                    End If
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
                MsgBox("时期统计未能完成统计" & vbCrLf & ex.Message)
                Return
            End Try
        End If
        CL1.SetItemChecked(0, False)
        CL3.SetItemChecked(0, False)
        CO10.Text = "" : CO11.Text = "" : DA10.ClearSelection()
    End Sub
    Private Sub R16_CheckedChanged(sender As Object, e As EventArgs) Handles R16.CheckedChanged
        D8.Checked = Not DirectCast(sender, RadioButton).Checked
        If DirectCast(sender, RadioButton).Checked Then
            CL3.Items.Clear() : CL3.Items.Add("全部选择")
            sbl(0) = suer = 0 OrElse suer = 1
            sbl(1) = suer = 2 OrElse suer = 4
            sbl(2) = suer = 4 OrElse suer = 6
            sbl(3) = suer = 3 OrElse suer = 5 OrElse suer = 6
            If Not sbl(2) Then
                cnct.Open()
                cmd = New SqlCommand("select 统计类型 from 统计类型 where 可用性=1 order by id", cnct)
                dr = cmd.ExecuteReader
                While dr.Read
                    CL3.Items.Add(dr(0))
                End While
                cnct.Close()
            End If
            CO10.Enabled = False : CO11.Enabled = False : rb = True
            DA10.Columns.Clear()
            DA10.Columns.Add("日期", "日期")
            DA10.Columns.Add("物料名称", "物料名称")
            DA10.Columns.Add("物料类型", "物料类型")
            DA10.Columns.Add("物料数量", "物料数量")
            For Each C As DataGridViewColumn In DA10.Columns
                C.Width = 175
            Next
            DA10.Columns(2).ContextMenuStrip = CMS2
        Else
            CL3.Items.Clear() : CL3.Items.Add("全部")
            CL3.Items.Add("消耗") : CL3.Items.Add("产出")
            CL3.Items.Add("回收") : CL3.Items.Add("入库")
            CO10.Enabled = True : CO11.Enabled = True
            rb = False : L42.Text = "~"
            DA10.Columns.Clear()
            DA10.Columns.Add("起始时间", "起始时间")
            DA10.Columns.Add("结束时间", "结束时间")
            DA10.Columns.Add("物料名称", "物料名称")
            DA10.Columns.Add("物料数量", "物料数量")
            DA10.Columns.Add("物料类型", "物料类型")
            DA10.Columns.Add("班别班组", "班别班组")
            DA10.Columns.Add("反应釜号", "反应釜号")
        End If
        If Not ary.ContainsKey(L40) Then ary.Add(L40, New DataTable)
        If Not ary.ContainsKey(L41) Then ary.Add(L41, New DataTable)
        ary(L41).Reset()
        ary(L41).Columns.Add("Id", Type.GetType("System.Int32"))
        ary(L41).Columns.Add("项目")
        For i = 0 To CL3.Items.Count - 1
            ary(L41).Rows.Add(i, CL3.Items(i))
        Next
    End Sub
    Private Sub B49_Click(sender As Object, e As EventArgs) Handles B49.Click
        Fcsb.s8(DA10, True)
    End Sub
    Private Sub B45_Click(sender As Object, e As EventArgs) Handles B45.Click
        Dim blct As Boolean
        If Not ckbl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            ckbl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
            blct = True
        End If
        s37(blct)
    End Sub
    Private Sub B46_Click(sender As Object, e As EventArgs) Handles B46.Click
        Fcsb.s8(DA9, True)
    End Sub
    Private Sub B47_Click(sender As Object, e As EventArgs) Handles B47.Click
        Fcsb.s4(DA9, "报表备注", idt4)
    End Sub
    Private Sub B80_Click(sender As Object, e As EventArgs) Handles B80.Click
        Fcsb.s7(DirectCast(sender, Button), B47, DA9, idt4)
    End Sub
    Private Sub T1_GotFocus(sender As Object, e As EventArgs) Handles T1.GotFocus
        AcceptButton = B13
    End Sub
    Private Sub T2_GotFocus(sender As Object, e As EventArgs) Handles T2.GotFocus
        AcceptButton = B28
    End Sub
    Private Sub T2_LostFocus(sender As Object, e As EventArgs) Handles T2.LostFocus
        If TC1.SelectedIndex = 1 Then AcceptButton = B24
    End Sub
    Private Sub T28_GotFocus(sender As Object, e As EventArgs) Handles T28.GotFocus
        AcceptButton = Nothing
    End Sub
    Private Sub TC1_KeyDown(sender As Object, e As KeyEventArgs) Handles TC1.KeyDown
        If e.KeyCode = Keys.Escape Then LI15.Hide()
    End Sub
    Private Sub TC1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TC1.SelectedIndexChanged
        Dim tb1, tb2 As New DataTable, k0 As New List(Of String), TC As TabControl = DirectCast(sender, TabControl)
        If TC.SelectedIndex = 5 Then
            If Not suer = 0 Then
                whbl = True
                TC.SelectedIndex = 0
                T50.Focus()
            End If
            If CB1.Text = "" Then
                G16.Enabled = True
                cmd = New SqlCommand("select 盘存类型 from 系统配置", cnct)
                cnct.Open()
                If IsDBNull(cmd.ExecuteScalar) Then
                    CH1.Checked = False : CH2.Checked = False : CH3.Checked = False
                ElseIf CByte(cmd.ExecuteScalar) = 0 Then
                    CH1.Checked = False : CH2.Checked = True : CH3.Checked = False
                ElseIf CByte(cmd.ExecuteScalar) = 1 Then
                    CH1.Checked = True : CH2.Checked = False : CH3.Checked = False
                Else
                    CH1.Checked = False : CH2.Checked = False : CH3.Checked = True
                End If
                cnct.Close()
            Else
                G16.Enabled = False
            End If
            WindowState = FormWindowState.Normal
            Size = New Size(1114, 663)
            CenterToScreen()
        ElseIf TC.SelectedIndex = 0 Then
            AcceptButton = B14
        ElseIf TC.SelectedIndex = 1 Then
            AcceptButton = B24
        End If
        If whbl Then
            Try
                LI1.Items.Clear()
                DA1.Rows.Clear() : DA2.Rows.Clear() : LI2.Items.Clear()
                LI4.Items.Clear() : LI6.Items.Clear() : LI8.Items.Clear()
                cnct.Open()
                cmdstr = "select 操作工序,id from 操作工序 where 可用性=1"
                cmdstr += " and not exists(select 1 from  用户工序 where 操作人员='" & usr & "') union select 用户工序.操作工序,id from 用户工序,操作工序 where 操作人员='" & usr & "' and exists(select 1 from 用户工序 where 操作人员='" & usr & "') and 操作工序.操作工序=用户工序.操作工序 and 操作工序.可用性=1"
                cmdstr += " order by id"
                Fcsb.s3(CL2, cmdstr)
                CL2.SetItemChecked(0, True)
                dr = New SqlCommand(cmdstr, cnct).ExecuteReader
                CB6.Items.Clear()
                While dr.Read()
                    CB6.Items.Add(dr(0))
                End While
                CB6.Items.Add("")
                dr.Close()
                If suer <> 4 AndAlso suer <> 5 Then Fcsb.s3(CL4, cmdstr)
                If suer <> 4 AndAlso suer <> 5 Then CL4.SetItemChecked(0, True)
                Fcsb.s3(LI5, cmdstr)
                Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(8), DataGridViewComboBoxColumn))
                cmdstr = "select 物料类型 from 物料类型 where 可用性=1 order by id"
                Fcsb.s3(LI3, cmdstr)
                Fcsb.s6(cmdstr, DirectCast(DA1.Columns.Item(6), DataGridViewComboBoxColumn))
                tb1.Reset()
                tb2.Reset()
                tb1.Columns.Add("物料名称")
                tb2.Columns.Add("物料名称")
                For Each r In LI3.Items
                    tb1.Rows.Add(r)
                Next
                For Each r In LI5.Items
                    tb2.Rows.Add(r)
                Next
                CL1.Items.Clear()
                DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Clear()
                If LI3.Items.Count > 0 Then
                    cmd = New SqlCommand("消耗产量", cnct)
                    cmd.Parameters.Add(New SqlParameter("操作工序", tb2))
                    cmd.Parameters.Add(New SqlParameter("物料类型", tb1))
                    cmd.Parameters.Add(New SqlParameter("类型", 1))
                    cmd.CommandType = CommandType.StoredProcedure
                    Fcsb.s53(dtn, cmd)
                    dr = cmd.ExecuteReader
                    While dr.Read
                        LI1.Items.Add(dr(0))
                        DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Add(dr(0))
                    End While
                    DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Add("")
                    dr.Close()
                End If
                CL1.Items.Add("全部")
                cmd = New SqlCommand("消耗产量", cnct)
                cmd.Parameters.Add(New SqlParameter("操作工序", tb2))
                cmd.Parameters.Add(New SqlParameter("物料类型", tb1))
                cmd.Parameters.Add(New SqlParameter("类型", 2))
                cmd.CommandType = CommandType.StoredProcedure
                dr = cmd.ExecuteReader
                While dr.Read
                    CL1.Items.Add(dr(0))
                End While
                dr.Close()
                cmdstr = "select 储槽名称,位号 from 储槽特性,物料特性,操作工序 where 储槽特性.可用性=1 and 物料特性.物料名称=储槽特性.物料名称 and 操作工序.可用性=1 and 储槽特性.操作工序=操作工序.操作工序 order by 储槽特性.物料名称"
                Fcsb.s53(dto, New SqlCommand(cmdstr, cnct))
                Fcsb.s3(LI7, cmdstr)
                Fcsb.s6(cmdstr, DirectCast(DA2.Columns.Item(2), DataGridViewComboBoxColumn))
                s15(tb1, tb2)
                s31(CO8)
                cmdstr = "select 储槽名称 from 储槽特性,操作工序 where 储槽特性.可用性=1 and 操作工序.可用性=1 and 储槽特性.操作工序=操作工序.操作工序 order by 储槽特性.id"
                s14()
                cmdstrgx = "("
                Fcsb.s29()
                DA6.Rows(0).Cells(5).Value = ""
                s41()
                s59(Form2 IsNot Nothing, False)
                whbl = False
            Catch ex As Exception
                cnct.Close()
            End Try
        End If
    End Sub
    Private Sub B50_Click(sender As Object, e As EventArgs) Handles B50.Click
        DA11.Columns.Clear()
        nn = False
        cmd = New SqlCommand(T28.Text, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader()
            If dr.FieldCount = 0 Then
                MsgBox("操作成功！")
            Else
                For i = 1 To dr.FieldCount
                    DA11.Columns.Add(dr.GetName(i - 1), dr.GetName(i - 1))
                    DA11.Columns(i - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Next
                Fcsb.s56(DA11)
                RemoveHandler DA11.CellValueChanged, AddressOf DA11_CellValueChanged
                While dr.Read()
                    DA11.Rows.Add()
                    For i = 1 To dr.FieldCount
                        DA11.Rows(DA11.Rows.Count - 2).Cells(i - 1).Value = IIf(IsDBNull(dr(i - 1)), Nothing, dr(i - 1))
                    Next
                End While
                AddHandler DA11.CellValueChanged, AddressOf DA11_CellValueChanged
            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox("语句查询失败！" & vbCrLf & ex.Message)
        End Try
        DA11.ClearSelection()
    End Sub
    Private Sub B52_Click(sender As Object, e As EventArgs) Handles B52.Click
        If Not ccbl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            ccbl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        s39(True)
    End Sub
    Private Sub B51_Click(sender As Object, e As EventArgs) Handles B51.Click
        Fcsb.s8(DA11, True)
    End Sub
    Private Sub B53_Click(sender As Object, e As EventArgs) Handles B53.Click, LI9.DoubleClick
        s20(LI9, LI12, "物料", True)
        whbl = True
    End Sub
    Private Sub B54_Click(sender As Object, e As EventArgs) Handles B54.Click, LI12.DoubleClick
        s20(LI12, LI9, "物料", False)
        whbl = True
    End Sub
    Private Sub B60_Click(sender As Object, e As EventArgs) Handles B60.Click, LI10.DoubleClick
        s20(LI10, LI13, "储槽", True)
        whbl = True
    End Sub
    Private Sub B61_Click(sender As Object, e As EventArgs) Handles B61.Click, LI13.DoubleClick
        s20(LI13, LI10, "储槽", False)
        whbl = True
    End Sub
    Private Sub B55_Click(sender As Object, e As EventArgs) Handles B55.Click
        Fcsb.s20(LI9, LI12, "物料", True)
        whbl = True
    End Sub
    Private Sub B56_Click(sender As Object, e As EventArgs) Handles B56.Click
        Fcsb.s20(LI12, LI9, "物料", False)
        whbl = True
    End Sub
    Private Sub B62_Click(sender As Object, e As EventArgs) Handles B62.Click
        Fcsb.s20(LI10, LI13, "储槽", True)
        whbl = True
    End Sub
    Private Sub B63_Click(sender As Object, e As EventArgs) Handles B63.Click
        Fcsb.s20(LI13, LI10, "储槽", False)
        whbl = True
    End Sub
    Private Sub B57_Click(sender As Object, e As EventArgs) Handles B57.Click
        Fcsb.s21(LI12, "物料")
        whbl = True
    End Sub
    Private Sub B64_Click(sender As Object, e As EventArgs) Handles B64.Click
        Fcsb.s21(LI13, "储槽")
        whbl = True
    End Sub
    Private Sub B58_Click(sender As Object, e As EventArgs) Handles B58.Click
        Fcsb.s22(LI12, -1)
    End Sub
    Private Sub B65_Click(sender As Object, e As EventArgs) Handles B65.Click
        Fcsb.s22(LI13, -1)
    End Sub
    Private Sub B59_Click(sender As Object, e As EventArgs) Handles B59.Click
        Fcsb.s22(LI12, 1)
    End Sub
    Private Sub B66_Click(sender As Object, e As EventArgs) Handles B66.Click
        Fcsb.s22(LI13, 1)
    End Sub
    Private Sub CB1_TextChanged(sender As Object, e As EventArgs) Handles CB1.TextChanged
        Dim CB As ComboBox = DirectCast(sender, ComboBox), cnctn As SqlConnection = New SqlConnection(String.Concat(New String() {"data source=", st(3), ";initial catalog=", st(2), ";user id=", usr, ";password=", pswd}))
        RemoveHandler CB.TextChanged, AddressOf CB1_TextChanged
        For Each ct As Control In G9.Controls
            If CStr(ct.Tag) <> "" Then
                If ct.Controls.Count = 0 Then
                    If TypeOf ct Is CheckBox Then
                        DirectCast(ct, CheckBox).Checked = False
                    ElseIf ct IsNot CB1 Then
                        ct.Text = ""
                    End If
                ElseIf ct IsNot G16 Then
                    For Each bt As Control In ct.Controls
                        DirectCast(bt, RadioButton).Checked = False
                    Next
                End If
            End If
        Next
        If CB.Text <> "" Then
            cmd = New SqlCommand("select * from 物料特性 where 物料名称=@物料名称", cnct)
            cmd.Parameters.Add(New SqlParameter("物料名称", CB.Text))
            G16.Enabled = False
            Try
                cnct.Open()
                dr = cmd.ExecuteReader()
                While dr.Read
                    For Each ct As Control In G9.Controls
                        If CStr(ct.Tag) <> "" Then
                            If ct.Controls.Count = 0 Then
                                If Not IsDBNull(dr(CStr(ct.Tag))) Then
                                    If TypeOf ct Is CheckBox Then
                                        DirectCast(ct, CheckBox).Checked = CBool(dr(CStr(ct.Tag)))
                                    Else
                                        ct.Text = CStr(dr(CStr(ct.Tag)))
                                    End If
                                End If
                            ElseIf Not IsDBNull(dr(CStr(ct.Tag))) Then
                                For Each bt As RadioButton In ct.Controls
                                    If CStr(bt.Tag) = CStr(dr(CStr(ct.Tag))) Then
                                        bt.Checked = True
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Next
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
                MsgBox("无法获取物料信息" & vbCrLf & ex.Message)
            End Try
        Else
            cnct.Open()
            dr = New SqlCommand("select 盘存类型 from 系统配置", cnct).ExecuteReader
            While dr.Read
                For Each ct As Control In G16.Controls
                    If CByte(ct.Tag) = CByte(dr(0)) Then DirectCast(ct, RadioButton).Checked = True
                Next
            End While
            cnct.Close()
            G16.Enabled = True
        End If
        If DA12.Columns.Count <> 0 AndAlso B106.Text = "储槽表格" Then
            DA12.Rows.Clear()
            If CB.Items.Contains(CB.Text) Then
                Try
                    cnct.Open()
                    cmdstr = "select * from 工序类型 where 物料名称=@物料名称"
                    cmd = New SqlCommand(cmdstr, cnct)
                    cmd.Parameters.Add(New SqlParameter("物料名称", CB.Text))
                    dr = cmd.ExecuteReader()
                    While dr.Read()
                        DA12.Rows.Add()
                        For i = 0 To 6
                            DA12.Rows(DA12.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                        Next
                        If CBool(dr(7)) Then
                            DA12.Rows(DA12.Rows.Count - 2).Cells(0).Style.ForeColor = Color.Green
                        Else
                            DA12.Rows(DA12.Rows.Count - 2).Cells(0).Style.ForeColor = Color.Red
                        End If
                        cnctn.Open()
                        Dim cmdn As SqlCommand = New SqlCommand("select R,G,B from 单耗类别 where Id=@Id", cnctn)
                        cmdn.Parameters.AddWithValue("Id", dr(8))
                        Dim drn As SqlDataReader = cmdn.ExecuteReader()
                        If drn.HasRows Then
                            While drn.Read()
                                DA12.Rows(DA12.Rows.Count - 2).Cells(0).Style.BackColor = Color.FromArgb(CByte(drn(0)), CByte(drn(1)), CByte(drn(2)))
                            End While
                        Else
                            DA12.Rows(DA12.Rows.Count - 2).Cells(0).Style.BackColor = DA12.RowsDefaultCellStyle.BackColor
                        End If
                        cnctn.Close()
                    End While
                    cnct.Close()
                    DA12.Columns(0).Visible = True
                    DA12.CurrentCell = DA12.Rows(DA12.Rows.Count - 1).Cells(0)
                    DA12.Rows(DA12.NewRowIndex).Cells(6).ReadOnly = True
                    DA12.FirstDisplayedScrollingRowIndex = 0
                Catch ex As Exception
                    cnct.Close()
                    cnctn.Close()
                Finally
                    cnctn.Dispose()
                End Try
            End If
            DA12.ClearSelection()
            If CB.Text = "" Then DA12.Rows(DA12.NewRowIndex).Cells(6).ReadOnly = True
        End If
        AddHandler CB.TextChanged, AddressOf CB1_TextChanged
    End Sub
    Private Sub CB2_TextChanged(sender As Object, e As EventArgs) Handles CB2.TextChanged
        Dim CB As ComboBox = DirectCast(sender, ComboBox)
        RemoveHandler CB.TextChanged, AddressOf CB2_TextChanged
        CH15.Checked = False
        For Each ct As Control In G14.Controls
            If CStr(ct.Tag) <> "" Then
                If ct Is CH15 Then
                    CH15.Checked = False
                ElseIf ct IsNot CB Then
                    ct.Text = ""
                End If
            End If
        Next
        If CB.Text <> "" Then
            cmdstr = "select * from 储槽特性 where 储槽名称=@储槽名称"
            cmd = New SqlCommand(cmdstr, cnct)
            cmd.Parameters.Add(New SqlParameter("储槽名称", CB.Text))
            Try
                cnct.Open()
                dr = cmd.ExecuteReader
                While dr.Read
                    For Each ct As Control In G14.Controls
                        If ct Is CH15 Then
                            CH15.Checked = CBool(dr(CStr(CH15.Tag)))
                        ElseIf CStr(ct.Tag) <> "" Then
                            If Not IsDBNull(dr(CStr(ct.Tag))) Then ct.Text = CStr(dr(CStr(ct.Tag)))
                        End If
                    Next
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
                MsgBox("无法获取储槽信息" & vbCrLf & ex.Message)
            End Try
        End If
        If DA12.Columns.Count <> 0 AndAlso B90.Text = "物料表格" Then
            DA12.Rows.Clear()
            If CB.Items.Contains(CB.Text) Then
                Try
                    cnct.Open()
                    cmdstr = "select * from 储槽物料 where 储槽名称=@储槽名称"
                    cmd = New SqlCommand(cmdstr, cnct)
                    cmd.Parameters.Add(New SqlParameter("储槽名称", CB.Text))
                    dr = cmd.ExecuteReader
                    While dr.Read
                        DA12.Rows.Add()
                        DA12.Rows(DA12.Rows.Count - 2).Cells(0).Value = dr(0)
                        DA12.Rows(DA12.Rows.Count - 2).Cells(1).Value = Format(dr(1), "yyyy-MM-dd HH:mm")
                        DA12.Rows(DA12.Rows.Count - 2).Cells(2).Value = dr(2)
                        DA12.Rows(DA12.Rows.Count - 2).Cells(3).Value = IIf(IsDBNull(dr(3)), Nothing, dr(3))
                        DA12.Rows(DA12.Rows.Count - 2).Cells(4).Value = IIf(IsDBNull(dr(4)), Nothing, dr(4))
                    End While
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                End Try
            End If
            DA12.ClearSelection()
        End If
        AddHandler CB.TextChanged, AddressOf CB2_TextChanged
    End Sub
    Private Sub B78_Click(sender As Object, e As EventArgs) Handles B78.Click
        Dim str As String = CStr(IIf(CB2.Items.Contains(CB2.Text), "更改", "添加"))
        If CB2.Items.Contains(CB2.Text) Then
            cmdstr = "update 储槽特性 set "
            For Each ct As Control In G14.Controls
                If CStr(ct.Tag) <> "" Then
                    cmdstr += "[" & CStr(ct.Tag) & "]" & "=" & "@" & Replace(Replace(CStr(ct.Tag), "(", ""), ")", "") & ","
                End If
            Next
            cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & " where 储槽名称=@储槽名称"
        Else
            cmdstr = "insert into 储槽特性(Id,"
            For Each ct As Control In G14.Controls
                If CStr(ct.Tag) <> "" Then cmdstr += "[" & CStr(ct.Tag) & "],"
            Next
            cmdstr += "储槽类型) values(NULL,"
            For Each ct As Control In G14.Controls
                If CStr(ct.Tag) <> "" Then cmdstr += "@" & Replace(Replace(CStr(ct.Tag), "(", ""), ")", "") & ","
            Next
            cmdstr += "0)"
        End If
        cmd = New SqlCommand(cmdstr, cnct)
        For Each ct As Control In G14.Controls
            If CStr(ct.Tag) <> "" Then
                If ct.Controls.Count = 0 Then
                    If TypeOf ct Is CheckBox Then
                        cmd.Parameters.AddWithValue(CStr(ct.Tag), DirectCast(ct, CheckBox).Checked)
                    Else
                        cmd.Parameters.AddWithValue(Replace(Replace(CStr(ct.Tag), "(", ""), ")", ""), IIf(ct.Text = "", DBNull.Value, ct.Text))
                    End If
                End If
            End If
        Next
        Try
            cnct.Open()
            cmd.ExecuteNonQuery()
            cnct.Close()
            MsgBox(str & "储槽成功！")
            If str = "添加" Then
                CB2.Items.Add(CB2.Text)
                If DA12.Columns.Count > 0 Then
                    If B90.Text = "物料表格" Then
                        DirectCast(DA12.Columns(2), DataGridViewComboBoxColumn).Items.Remove("")
                        DirectCast(DA12.Columns(2), DataGridViewComboBoxColumn).Items.Add(CB2.Text)
                        DirectCast(DA12.Columns(2), DataGridViewComboBoxColumn).Items.Add("")
                    End If
                End If
            End If
            whbl = True
            If CH15.Checked Then
                LI10.Items.Remove(CB2.Text)
                If Not LI13.Items.Contains(CB2.Text) Then LI13.Items.Add(CB2.Text)
            Else
                LI13.Items.Remove(CB2.Text)
                If Not LI10.Items.Contains(CB2.Text) Then LI10.Items.Add(CB2.Text)
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox(str & "储槽时发生错误！" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub B76_Click(sender As Object, e As EventArgs) Handles B76.Click
        Dim bl As Boolean, str As String = CStr(IIf(CB1.Items.Contains(CB1.Text), "更改", "添加"))
        If CB1.Text = "" Then
            For Each ct As Control In G16.Controls
                If DirectCast(ct, RadioButton).Checked Then
                    cmdstr = "update 系统配置 set 盘存类型=@盘存类型"
                    cmd = New SqlCommand(cmdstr, cnct)
                    cmd.Parameters.AddWithValue("盘存类型", ct.Tag)
                    Try
                        cnct.Open()
                        cmd.ExecuteNonQuery()
                        cnct.Close()
                        whbl = True
                        MsgBox("当前盘存模式为 " & DirectCast(ct, RadioButton).Text)
                        Return
                    Catch ex As Exception
                        cnct.Close()
                        Return
                    End Try
                End If
            Next
        End If
        If CB1.Items.Contains(CB1.Text) Then
            cmdstr = "update 物料特性 set "
            For Each ct As Control In G9.Controls
                If CStr(ct.Tag) <> "" Then
                    cmdstr += CStr(ct.Tag) & "=" & "@" & CStr(ct.Tag) & ","
                End If
            Next
            cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & " where 物料名称=@物料名称"
        Else
            cmdstr = "insert into 物料特性(Id,"
            For Each ct As Control In G9.Controls
                If CStr(ct.Tag) <> "" Then cmdstr += "[" & CStr(ct.Tag) & "],"
            Next
            cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & ") values(NULL,"
            For Each ct As Control In G9.Controls
                If CStr(ct.Tag) <> "" Then cmdstr += "@" & CStr(ct.Tag) & ","
            Next
            cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & ")"
        End If
        cmd = New SqlCommand(cmdstr, cnct)
        For Each ct As Control In G9.Controls
            If CStr(ct.Tag) <> "" Then
                If ct.Controls.Count = 0 Then
                    If TypeOf ct Is CheckBox Then
                        cmd.Parameters.AddWithValue(CStr(ct.Tag), DirectCast(ct, CheckBox).Checked)
                    Else
                        cmd.Parameters.AddWithValue(CStr(ct.Tag), IIf(ct.Text = "", DBNull.Value, ct.Text))
                    End If
                ElseIf ct IsNot G16 Then
                    For Each bt As Control In ct.Controls
                        If DirectCast(bt, RadioButton).Checked Then
                            cmd.Parameters.AddWithValue(CStr(ct.Tag), CInt(DirectCast(bt, RadioButton).Tag))
                            bl = True
                        End If
                    Next
                    If Not bl Then
                        cmd.Parameters.AddWithValue(CStr(ct.Tag), DBNull.Value)
                    Else
                        bl = False
                    End If
                End If
            End If
        Next
        Try
            cnct.Open()
            cmd.ExecuteNonQuery()
            cnct.Close()
            MsgBox(str & "物料成功！")
            whbl = True
            If str = "添加" Then
                CB1.Items.Add(CB1.Text)
                If CH14.Checked Then
                    LI12.Items.Add(CB1.Text)
                Else
                    LI9.Items.Add(CB1.Text)
                End If
                If DA12.Columns.Count > 0 Then
                    If B106.Text = "储槽表格" Then
                        DirectCast(DA12.Columns(1), DataGridViewComboBoxColumn).Items.Remove("")
                        DirectCast(DA12.Columns(1), DataGridViewComboBoxColumn).Items.Add(CB1.Text)
                        DirectCast(DA12.Columns(1), DataGridViewComboBoxColumn).Items.Add("")
                    End If
                End If
                CB3.Items.Add(CB1.Text)
            Else
                If CH14.Checked Then
                    LI9.Items.Remove(CB1.Text)
                    If Not LI12.Items.Contains(CB1.Text) Then LI12.Items.Add(CB1.Text)
                Else
                    LI12.Items.Remove(CB1.Text)
                    If Not LI9.Items.Contains(CB1.Text) Then LI9.Items.Add(CB1.Text)
                End If
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox(str & "物料时发生错误！" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub B77_Click(sender As Object, e As EventArgs) Handles B77.Click
        If Not CB1.Items.Contains(CB1.Text) Then MsgBox("没有该物料！") : Return
        cmdstr = "delete from 物料特性 where 物料名称=@物料名称"
        cmd = New SqlCommand(cmdstr, cnct)
        cmd.Parameters.Add(New SqlParameter("物料名称", CB1.Text))
        Try
            cnct.Open()
            cmd.ExecuteNonQuery()
            cnct.Close()
            whbl = True
            MsgBox("删除物料成功！")
        Catch ex As Exception
            cnct.Close()
            MsgBox("删除物料失败！" & vbCrLf & ex.Message)
        End Try
        LI9.Items.Remove(CB1.Text)
        LI12.Items.Remove(CB1.Text)
        CB3.Items.Remove(CB1.Text)
        Dim str As String = CB1.Text
        CB1.Items.Remove(CB1.Text)
        RemoveHandler CB1.TextChanged, AddressOf CB1_TextChanged
        CB1.Text = str
        AddHandler CB1.TextChanged, AddressOf CB1_TextChanged
    End Sub
    Private Sub B79_Click(sender As Object, e As EventArgs) Handles B79.Click
        If Not CB2.Items.Contains(CB2.Text) Then MsgBox("没有该储槽！") : Return
        cmdstr = "delete from 储槽特性 where 储槽名称=@储槽名称"
        cmd = New SqlCommand(cmdstr, cnct)
        cmd.Parameters.Add(New SqlParameter("储槽名称", CB2.Text))
        Try
            cnct.Open()
            cmd.ExecuteNonQuery()
            cnct.Close()
            whbl = True
            MsgBox("删除储槽成功！")
        Catch ex As Exception
            cnct.Close()
            MsgBox("删除储槽失败！" & vbCrLf & ex.Message)
            Return
        End Try
        LI10.Items.Remove(CB2.Text)
        LI13.Items.Remove(CB2.Text)
        Dim str As String = CB2.Text
        CB2.Items.Remove(CB2.Text)
        RemoveHandler CB2.TextChanged, AddressOf CB2_TextChanged
        CB2.Text = str
        AddHandler CB2.TextChanged, AddressOf CB2_TextChanged
    End Sub
    Private Sub DA10_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA10.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.SelectedCells.Count = 1 AndAlso e.RowIndex > -1 AndAlso e.ColumnIndex = 0 AndAlso R16.Checked Then
            DA.Columns(DA.SelectedCells(0).ColumnIndex).Visible = True
            If e.Button = MouseButtons.Left Then
                DA.ClearSelection()
                DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True
                DA.Columns(e.ColumnIndex).Visible = True
                DA.CurrentCell = DA.Rows(e.RowIndex).Cells(e.ColumnIndex)
                DA.BeginEdit(False)
            End If
        End If
        Try
            If e.ColumnIndex <> 2 OrElse DA.Columns.Count <> 4 Then s4(DA, e.ColumnIndex, L82, L84, L80, L96, L98)
        Catch exception As Exception
        End Try
    End Sub
    Public Sub DA_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA3.CellEndEdit, DA5.CellEndEdit, DA6.CellEndEdit
        Dim bl As Boolean, DA As DataGridView = DirectCast(sender, DataGridView)
        Dim str1(1) As String
        ni = CInt(Strings.Right(DA.Name, 1))
        Dim str As String = CStr(IIf(ni = 3, " 07:59", " 08:00"))
        If e.ColumnIndex = 0 Then
            If CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
                bl = True
            Else
                DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value), str)
            End If
            If bl Then
                If DA.Rows(e.RowIndex).IsNewRow Then
                    ri = e.RowIndex
                    TM1.Enabled = True
                Else
                    DA.Rows(e.RowIndex).Cells(0).Value = Format(Now, "yyyy-MM-dd") + str
                End If
            End If
        End If
        If ni = 5 OrElse ni = 6 Then
            Dim d As String = CStr(DA.Rows(e.RowIndex).Cells(1).Value)
            Dim str2 As String
            Fcsb.s10(d, CH32.Checked)
            If CH32.Checked Then DA.Rows(e.RowIndex).Cells(1).Value = d
            If CH33.Checked Then DA.Rows(e.RowIndex).Cells(3).Value = Fcsb.s55(d)
            Dim s As Byte = Fcsb.s10(d, CH32.Checked)
            Try
                cnct.Open()
                dr = New SqlCommand("select distinct 物料类型,操作工序 from 工序类型 where 批号代码=" & s & " and 可用性=1", cnct).ExecuteReader
                While dr.Read
                    str1(1) = CStr(dr(1))
                    Exit While
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
            Try
                str2 = Fcsb.s24(CStr(DA.Rows(e.RowIndex).Cells(2).Value), CStr(DA.Rows(e.RowIndex).Cells(0).Value), str1(1))
            Catch ex As Exception
                cnct.Close()
                str2 = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
            End Try
            If e.ColumnIndex > -1 AndAlso e.ColumnIndex < 3 Then
                If e.ColumnIndex <> 2 Then DA.Rows(e.RowIndex).Cells(2).Value = str2
                If str2 <> CStr(DA.Rows(e.RowIndex).Cells(2).Value) Then
                    DA.Rows(e.RowIndex).Cells(2).Style.ForeColor = Color.Red
                Else
                    DA.Rows(e.RowIndex).Cells(2).Style.ForeColor = Color.Black
                End If
            ElseIf e.ColumnIndex = 5 AndAlso ni = 6 Then
                RemoveHandler DA6.CellEndEdit, AddressOf DA_CellEndEdit
                TM1.Enabled = True
            End If
        End If
    End Sub
    Public Sub DA5_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA5.CellMouseClick, DA6.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        Try
            If e.RowIndex = -1 Then
                If e.Button = Windows.Forms.MouseButtons.Middle Then
                    s30(DA, e, True)
                ElseIf e.Button = Windows.Forms.MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                    s30(DA, e, False)
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub T38_KeyDown(sender As Object, e As KeyEventArgs) Handles T38.KeyDown
        If DirectCast(sender, TextBox).Text = "" AndAlso suer <> 4 Then D1.Checked = False
    End Sub
    Private Sub T38_KeyUp(sender As Object, e As KeyEventArgs) Handles T38.KeyUp
        If DirectCast(sender, TextBox).Text = "" AndAlso suer <> 4 Then D1.Checked = True
    End Sub
    Public Sub T38_TextChanged(sender As Object, e As EventArgs) Handles T38.TextChanged
        Dim T As TextBox = DirectCast(sender, TextBox)
        LI15.SetBounds(T.Left + 16, T.Top + 61, T.Size.Width, 124)
        LI15.Font = New System.Drawing.Font("Times New Roman", 10)
        If suer <> 4 Then
            If T.Text <> "" Then
                Dim k0 As New List(Of String), cmdstr1 As String = ""
                Dim cmdstr2 As String = Fcsb.s5(D1, D2, "日期")
                If CL2.CheckedItems.Count > 0 Then
                    For Each r In CL2.CheckedItems
                        k0.Add(CStr(r))
                    Next
                    cmdstr1 = Fcsb.s2(k0, "操作工序")
                End If
                LI15.Items.Clear()
                Try
                    cnct.Open()
                    cmdstr = "select distinct top 10 批号 from 物料数量 where 批号 COLLATE Chinese_PRC_CI_AS like '%" & Replace(T.Text, "'", "''") & "%'"
                    If cmdstr1 <> "" Then cmdstr += " and " & cmdstr1
                    If cmdstr2 <> "(" Then
                        cmdstr += " and " & cmdstr2 & " order by 批号 desc"
                    Else
                        cmdstr += " order by 批号 desc"
                    End If
                    cmd = New SqlCommand(cmdstr, cnct)
                    dr = cmd.ExecuteReader
                    While dr.Read
                        LI15.Items.Add(dr(0))
                    End While
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                End Try
                If LI15.Items.Count > 0 Then
                    LI15.Show()
                Else
                    LI15.Hide()
                End If
            Else
                LI15.Hide()
            End If
        End If
    End Sub
    Private Sub LI15_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles LI15.MouseDoubleClick
        RemoveHandler T38.TextChanged, AddressOf T38_TextChanged
        T38.Text = CStr(DirectCast(sender, ListBox).SelectedItem)
        LI15.Hide()
        AddHandler T38.TextChanged, AddressOf T38_TextChanged
    End Sub
    Private Sub B82_Click(sender As Object, e As EventArgs) Handles B82.Click
        Try
            cnct.Open()
            cmdstr = "insert into 操作工序 values(0,'" & Replace(CB5.Text, "'", "''") & "',1)"
            cmd = New SqlCommand(cmdstr, cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
            MsgBox("增加操作工序成功！")
        Catch ex As Exception
            cnct.Close()
            MsgBox("新增操作工序有错误发生！" & vbCrLf & ex.Message)
        End Try
        cnct.Open()
        cmdstr = "select 操作工序 from 操作工序 order by id"
        DA12.Rows.Clear()
        If DA12.Columns.Count = 7 Then Fcsb.s6(cmdstr, DirectCast(DA12.Columns.Item(2), DataGridViewComboBoxColumn))
        CB5.Items.Clear()
        s31(CB5)
        cnct.Close()
        CB4.Items.Add(CB5.Text)
        s26(CL5, "操作工序")
        whbl = True
    End Sub
    Private Sub B81_Click(sender As Object, e As EventArgs) Handles B81.Click
        If CB5.Text = "" Then Return
        If Not CB5.Items.Contains(CB5.Text) Then
            MsgBox("不存在该操作工序")
            CB5.Text = "" : Return
        End If
        Try
            cnct.Open()
            cmdstr = "delete from 操作工序 where 操作工序='" & Replace(CB5.Text, "'", "''") & "'"
            cmd = New SqlCommand(cmdstr, cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
            CB5.Items.Remove(CB5.Text)
            MsgBox("删除指定的操作工序成功！")
        Catch ex As Exception
            cnct.Close()
            MsgBox("删除指定的操作工序失败！" & vbCrLf & ex.Message)
        End Try
        cnct.Open()
        cmdstr = "select 操作工序 from 操作工序 order by id"
        DA12.Rows.Clear()
        If DA12.Columns.Count = 7 Then Fcsb.s6(cmdstr, DirectCast(DA12.Columns.Item(2), DataGridViewComboBoxColumn))
        cnct.Close()
        CB4.Items.Remove(CB5.Text)
        s26(CL5, "操作工序")
        whbl = True
    End Sub
    Private Sub B87_Click(sender As Object, e As EventArgs) Handles B87.Click
        s29(CL5, -1) : s28(CL5, "操作工序") : whbl = True
    End Sub
    Private Sub B88_Click(sender As Object, e As EventArgs) Handles B88.Click
        s29(CL6, -1) : s28(CL6, "物料类型") : whbl = True
    End Sub
    Private Sub B86_Click(sender As Object, e As EventArgs) Handles B86.Click
        s29(CL5, 1) : s28(CL5, "操作工序") : whbl = True
    End Sub
    Private Sub B89_Click(sender As Object, e As EventArgs) Handles B89.Click
        s29(CL6, 1) : s28(CL6, "物料类型") : whbl = True
    End Sub
    Private Sub B90_Click(sender As Object, e As EventArgs) Handles B90.Click
        Dim dgv As DataGridViewComboBoxColumn, dgv0, dgv6 As New DataGridViewTextBoxColumn, dgv1, dgv2, dgv3, dgv4, dgv5 As New DataGridViewComboBoxColumn
        If DirectCast(sender, Button).Text = "物料表格" Then
            DirectCast(sender, Button).Text = "解锁表格"
            DA12.Columns.Clear()
            dgv0.HeaderText = "Id" : dgv0.Width = 58 : dgv0.DefaultCellStyle.BackColor = Color.White : DA12.Columns.Add(dgv0) : dgv0.SortMode = DataGridViewColumnSortMode.Automatic
            dgv1.HeaderText = "物料名称" : dgv1.Width = 128 : DA12.Columns.Add(dgv1) : dgv1.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing : dgv1.SortMode = DataGridViewColumnSortMode.Automatic
            dgv2.HeaderText = "工序" : dgv2.Width = 66 : DA12.Columns.Add(dgv2) : dgv2.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing : dgv2.SortMode = DataGridViewColumnSortMode.Automatic
            dgv3.HeaderText = "类型" : dgv3.Width = 60 : DA12.Columns.Add(dgv3) : dgv3.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing : dgv3.SortMode = DataGridViewColumnSortMode.Automatic
            dgv4.HeaderText = "批号" : dgv4.Width = 60 : DA12.Columns.Add(dgv4) : dgv4.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing : dgv4.SortMode = DataGridViewColumnSortMode.Automatic
            dgv5.HeaderText = "釜号" : dgv5.Width = 60 : DA12.Columns.Add(dgv5) : dgv5.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing : dgv5.SortMode = DataGridViewColumnSortMode.Automatic
            dgv6.HeaderText = "单耗" : dgv6.Width = 60 : DA12.Columns.Add(dgv6) : dgv6.SortMode = DataGridViewColumnSortMode.Automatic
            DA12.Columns(1).Tag = "物料名称" : DA12.Columns(2).Tag = "操作工序" : DA12.Columns(3).Tag = "物料类型"
            DA12.Columns(4).Tag = "批号代码" : DA12.Columns(5).Tag = "可用釜号" : DA12.Columns(6).Tag = "单耗预估值"
            cnct.Open()
            Fcsb.s6("select 物料名称 from 物料特性 order by id", DirectCast(DA12.Columns(1), DataGridViewComboBoxColumn))
            Fcsb.s6("select 操作工序 from 操作工序 order by id", DirectCast(DA12.Columns(2), DataGridViewComboBoxColumn))
            Fcsb.s6("select 物料类型 from 物料类型 order by id", DirectCast(DA12.Columns(3), DataGridViewComboBoxColumn))
            cnct.Close()
            dgv = DirectCast(DA12.Columns(4), DataGridViewComboBoxColumn)
            cnct.Open()
            cmd = New SqlCommand("select 批号代码 from 批号代码", cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                dgv.Items.Add(dr(0))
            End While
            cnct.Close()
            dgv.Items.Add("")
            dgv = DirectCast(DA12.Columns(5), DataGridViewComboBoxColumn)
            cnct.Open()
            cmd = New SqlCommand("select 反应釜号 from 反应釜号", cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                dgv.Items.Add(dr(0))
            End While
            cnct.Close()
            dgv.Items.Add("")
            B106.Text = "储槽表格"
            B107.Enabled = False
            DA12.ReadOnly = True
            DA12.Rows(DA12.NewRowIndex).Cells(6).ReadOnly = True
        Else
            Fcsb.s7(DirectCast(sender, Button), B85, DA12)
            If B90.Text = "锁定表格" Then DA12.Rows(DA12.NewRowIndex).Cells(6).ReadOnly = True
        End If
    End Sub
    Private Sub B85_Click(sender As Object, e As EventArgs) Handles B85.Click
        Fcsb.s4(DA12, "工序类型") : whbl = True
    End Sub
    Public Sub DA12_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA12.CellBeginEdit
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        sv = Nothing
        For Each col As DataGridViewColumn In DA.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        If IsNothing(DA.Rows(e.RowIndex).Cells(0).Value) Then Return
        Try
            If B106.Text = "储槽表格" Then
                cmdstr = "select * from 工序类型 where Id="
            Else
                cmdstr = "select * from 储槽物料 where Id="
            End If
            cnct.Open()
            dr = New SqlCommand(cmdstr & CStr(DA.Rows(e.RowIndex).Cells(0).Value), cnct).ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    sv = IIf(IsDBNull(dr(e.ColumnIndex)), Nothing, dr(e.ColumnIndex))
                    If Not skip(1) Then
                        For Each col As DataGridViewColumn In DA.Columns
                            If TypeOf col Is DataGridViewComboBoxColumn AndAlso Not DirectCast(col, DataGridViewComboBoxColumn).Items.Contains(IIf(IsDBNull(dr(col.Index)), "", dr(col.Index))) Then
                                MsgBox("数据库值:" & CStr(IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))) & " 不在" & col.Name & "列表中！")
                            ElseIf col.Index = 1 AndAlso B90.Text = "物料表格" Then
                                DA.Rows(e.RowIndex).Cells(1).Value = Format(CDate(dr(1)), "yyyy-MM-dd HH:mm")
                            Else
                                DA.Rows(e.RowIndex).Cells(col.Index).Value = IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))
                            End If
                        Next
                    End If
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
    Public Sub DA12_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA12.CellEndEdit
        Dim str() As String, rst As Date, srt As Decimal, DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
            DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
        End If
        If B106.Text = "储槽表格" AndAlso e.ColumnIndex = 4 Then
            If DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value IsNot Nothing Then
                DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = CByte(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
            End If
        End If
        If IsNothing(DA.Rows(e.RowIndex).Cells(0).Value) Then
            If B90.Text = "物料表格" Then
                If e.ColumnIndex = 1 Then
                    If DA.NewRowIndex = e.RowIndex Then
                        ni = 12
                        ri = e.RowIndex
                        skip(0) = True
                        TM1.Enabled = True
                    ElseIf IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then
                        DA.Rows(e.RowIndex).Cells(1).Value = Format(Now, "yyyy-MM-dd 07:59")
                    Else
                        DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "07:59")
                    End If
                ElseIf e.ColumnIndex > 2 Then
                    If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
                    ElseIf Decimal.TryParse(CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value), srt) Then
                        DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = CDec(Format(srt, "0.0000"))
                    End If
                End If
            End If
        ElseIf CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 Then
            cmdstr = "" : skip(1) = False : skip(0) = False
            If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CStr(sv) Then
                If B106.Text = "储槽表格" Then
                    ReDim str(0)
                    For i = 1 To DA.Columns.Count - 2
                        str(UBound(str)) = "'" + Replace(CStr(DA.Rows(e.RowIndex).Cells(i).Value), "'", "''") + "'"
                        ReDim Preserve str(i)
                    Next
                    ReDim Preserve str(UBound(str) - 1)
                    For i = 0 To UBound(str)
                        If str(i) = "''" Then str(i) = "NULL"
                    Next
                    If e.ColumnIndex = 6 Then
                        If scm.ContainsKey(DA.Rows(e.RowIndex).Cells(0).Style.BackColor) Then
                            If CStr(DA.Rows(e.RowIndex).Cells(6).Value) = "" Then
                                cmdstr = "update 工序类型 set 单耗预估值=NULL where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                            ElseIf IsNumeric(DA.Rows(e.RowIndex).Cells(6).Value) Then
                                cmdstr = "update 工序类型 set 单耗预估值=" & CStr(DA.Rows(e.RowIndex).Cells(6).Value) & " where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                                DA.Rows(e.RowIndex).Cells(6).Value = CDec(Format(CDec(DA.Rows(e.RowIndex).Cells(6).Value), "0.000"))
                            Else
                                s33("单耗预估值", DA, e)
                                Return
                            End If
                        Else
                            cmdstr = "update 工序类型 set 单耗预估值=NULL where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                            DA.Rows(e.RowIndex).Cells(6).Value = Nothing
                        End If
                    Else
                        If Fcsb.s26(str(0), str(1), str(2), str(3), str(4)) Then
                            s33("工序类型有重复！", DA, e)
                            Return
                        ElseIf CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
                            cmdstr = "update 工序类型 set " & CStr(DA.Columns(e.ColumnIndex).Tag) & "=NULL where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                        Else
                            cmdstr = "update 工序类型 set " & CStr(DA.Columns(e.ColumnIndex).Tag) & "=" & str(e.ColumnIndex - 1) & " where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                        End If
                    End If
                Else
                    If e.ColumnIndex = 1 Then
                        DA.Rows(e.RowIndex).Cells(1).Value = Fcsb.s48(CStr(DA.Rows(e.RowIndex).Cells(1).Value))
                        If Not Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), rst) Then
                            s33(DA.Columns(1).HeaderText, DA, e)
                            Return
                        Else
                            DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Second, 30, rst), "yyyy-MM-dd HH:mm")
                        End If
                    ElseIf e.ColumnIndex <> 2 Then
                        If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = "" Then
                        ElseIf Not Decimal.TryParse(CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value), srt) Then
                            s33(DA.Columns(e.ColumnIndex).HeaderText, DA, e)
                            Return
                        Else
                            DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = CDec(Format(srt, "0.0000"))
                        End If
                    End If
                    If IsNothing(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
                        cmdstr = "update 储槽物料 set " & DA.Columns(e.ColumnIndex).HeaderText & "=NULL where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                    Else
                        cmdstr = "update 储槽物料 set " & DA.Columns(e.ColumnIndex).HeaderText & "='" & Replace(CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value), "'", "''") & "' where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                    End If
                End If
                Try
                    If cmdstr <> "" Then
                        cnct.Open()
                        cmd = New SqlCommand(cmdstr, cnct)
                        cmd.ExecuteNonQuery()
                        cnct.Close()
                        whbl = True
                        If e.ColumnIndex = 6 Then s25(DA, e.RowIndex)
                    End If
                Catch ex As Exception
                    cnct.Close()
                    s33(If(IsNothing(DA.Columns(e.ColumnIndex).Tag), DA.Columns(e.ColumnIndex).HeaderText, CStr(DA.Columns(e.ColumnIndex).Tag)), DA, e)
                End Try
            End If
        End If
    End Sub
    Public Sub DA12_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA12.RowValidating
        Dim rst As Date, srt As Decimal, DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            DA.EndEdit()
            If B90.Text = "物料表格" AndAlso IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then
                DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Day, -1, Now), "yyyy-MM-dd 07:59")
            End If
            Dim str() As String
            ReDim str(0)
            For i = 1 To DA.Columns.Count - 1
                str(UBound(str)) = String.Concat("'", Replace(Replace(CStr(DA.Rows(e.RowIndex).Cells(i).Value), "：", ":"), "'", "''"), "'")
                ReDim Preserve str(i)
            Next
            ReDim Preserve str(UBound(str) - 1)
            For i = 0 To UBound(str)
                If str(i) = "''" Then str(i) = "NULL"
            Next
            If B90.Text = "物料表格" Then
                If Not Date.TryParse(CStr(DA.Rows(e.RowIndex).Cells(1).Value), rst) Then
                    Fcsb.s16(DA, 1, e)
                    Return
                End If
                If IsNothing(DA.Rows(e.RowIndex).Cells(2).Value) Then
                    Fcsb.s16(DA, 2, e)
                    Return
                End If
                If IsNothing(DA.Rows(e.RowIndex).Cells(3).Value) Then
                ElseIf Not Decimal.TryParse(CStr(DA.Rows(e.RowIndex).Cells(3).Value), srt) Then
                    Fcsb.s16(DA, 3, e)
                    Return
                ElseIf srt < 0 OrElse srt > 1 Then
                    Fcsb.s16(DA, 3, e)
                    Return
                Else
                    DA.Rows(e.RowIndex).Cells(3).Value = srt
                End If
                If IsNothing(DA.Rows(e.RowIndex).Cells(4).Value) Then
                ElseIf Not Decimal.TryParse(CStr(DA.Rows(e.RowIndex).Cells(4).Value), srt) Then
                    Fcsb.s16(DA, 4, e)
                    Return
                ElseIf srt <= 0 Then
                    Fcsb.s16(DA, 4, e)
                    Return
                Else
                    DA.Rows(e.RowIndex).Cells(4).Value = srt
                End If
                cmdstr = "insert into 储槽物料 values("
                For i = 1 To DA.Columns.Count - 1
                    cmdstr += str(i - 1) & ","
                Next
                cmdstr = Strings.Left(cmdstr, cmdstr.LastIndexOf(",") - 1) + Replace(cmdstr, ",", ")", cmdstr.LastIndexOf(","))
                cmdstr += "select max(Id) from 储槽物料"
            Else
                If IsNothing(DA.Rows(e.RowIndex).Cells(1).Value) Then
                    Fcsb.s16(DA, 1, e)
                    Return
                ElseIf IsNothing(DA.Rows(e.RowIndex).Cells(3).Value) Then
                    Fcsb.s16(DA, 3, e)
                    Return
                ElseIf Fcsb.s26(str(0), str(1), str(2), str(3), str(4)) Then
                    skip(0) = True
                    e.Cancel = True
                    DA.Columns(1).Visible = True
                    DA.CurrentCell = DA.Rows(e.RowIndex).Cells(1)
                    MsgBox("工序类型有重复！")
                    DA.BeginEdit(False)
                    Return
                End If
                cmdstr = "insert into 工序类型 values("
                For i = 1 To DA.Columns.Count - 2
                    cmdstr += str(i - 1) & ","
                Next
                cmdstr += "NULL,1,NULL)select max(Id) from 工序类型"
            End If
            Try
                cnct.Open()
                DA.Rows(e.RowIndex).Cells(0).Value = New SqlCommand(cmdstr, cnct).ExecuteScalar()
                If B106.Text = "储槽表格" Then
                    DA.Rows(e.RowIndex).Cells(6).Value = Nothing
                    DA.Rows(e.RowIndex).Cells(6).ReadOnly = False
                    DA.Rows(DA.NewRowIndex).Cells(6).ReadOnly = True
                    DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.Green
                Else
                    DA.Rows(e.RowIndex).Cells(1).Value = Format(DateAdd(DateInterval.Second, 30, CDate(DA.Rows(e.RowIndex).Cells(1).Value)), "yyyy-MM-dd HH:mm")
                End If
                whbl = True
            Catch ex As Exception
                whbl = False
                DA.Rows(e.RowIndex).ReadOnly = True
                DA.Rows(e.RowIndex).Cells(0).Value = IIf(B90.Text = "物料表格", CDec(0), CShort(0))
                MsgBox(String.Concat("记录提交未完全成功！" & vbCrLf & "", ex.Message))
            Finally
                cnct.Close()
                skip(0) = False
                For Each col As DataGridViewColumn In DA.Columns
                    col.SortMode = DataGridViewColumnSortMode.Automatic
                Next
            End Try
        End If
    End Sub
    Private Sub DA11_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DA11.CellContentClick
        Dim bl As Boolean, str2 As Integer, str1, ph As String, DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex < 0 OrElse DA.Columns.Count < 4 Then Return
        str1 = CStr(DA.Rows(e.RowIndex).Cells(3).Value)
        If nn Then
            If e.ColumnIndex = 4 Then
                str2 = CInt(DA.Rows(e.RowIndex).Cells(4).Value)
                If str1 = "物料数量" Then
                    s8(str2)
                    DA1.ClearSelection()
                ElseIf str1 = "储槽液位" Then
                    s9(str2)
                    DA2.ClearSelection()
                ElseIf str1 IsNot Nothing AndAlso suer <> 5 Then
                    Form2.Show()
                    cmdstr = "select [" & Strings.Left(str1, Len(str1) - 2) & "批号] from [" & str1 & "] where id=" & str2
                    Try
                        cnct.Open()
                        cmd = New SqlCommand(cmdstr, cnct)
                        ph = CStr(IIf(IsNothing(cmd.ExecuteScalar), "", cmd.ExecuteScalar))
                        cnct.Close()
                    Catch ex As Exception
                        cnct.Close()
                    End Try
                    If ph = "" Then
                        bl = True
                        cmdstr = "select TOP 1 SQL语句 from 操作记录 where 记录Id=" & str2 & " and 记录表='" & Replace(str1, "'", "''") & "' order by Id Desc"
                        cmd = New SqlCommand(cmdstr, cnctm)
                        Try
                            cnctm.Open()
                            ph = CStr(IIf(IsNothing(cmd.ExecuteScalar), "", cmd.ExecuteScalar))
                            cnctm.Close()
                            If ph.Contains("'NULL'") Then
                                ph = "NULL"
                            Else
                                ph = Replace(ph, "'''", "''")
                                ph = Replace(ph, "''", vbCrLf)
                                ph = Replace(ph, "'", "")
                                ph = Strings.Right(ph, Len(ph) - ph.IndexOf("=") - 1)
                                ph = Replace(ph, vbCrLf, "'")
                            End If
                        Catch ex As Exception
                            cnctm.Close()
                        End Try
                    End If
                    Try
                        cnctm.Open()
                        cmd = New SqlCommand("select 1 from [" & str1 & "] where Id=" & CStr(DA.Rows(e.RowIndex).Cells(4).Value), cnctm)
                        cmdstr = CStr(IIf(IsNothing(cmd.ExecuteScalar), "", cmd.ExecuteScalar))
                        cnctm.Close()
                    Catch ex As Exception
                        cnctm.Close()
                    End Try
                    Dim i As Integer = Form2.s4(Strings.Left(str1, Len(str1) - 2), 9)
                    Dim rg(,) As Object = Form2.rg
                    If i > -1 Then
                        rg(8, i) = ph
                        Form2.TC2.SelectedIndex = CInt(rg(0, i))
                        If bl Then
                            Dim med, key, value As String
                            Dim dt As New Dictionary(Of String, String)
                            DirectCast(rg(5, i), Button).Text = cmdstr
                            cnctm.Open()
                            cmd = New SqlCommand("select SQL语句 from 操作记录 where 记录Id=" & str2 & " and 记录表='" & Replace(str1, "'", "''") & "' order by Id", cnctm)
                            dr = cmd.ExecuteReader
                            While dr.Read
                                If Strings.Left(CStr(dr(0)), 1) = "u" Then
                                    med = Fcsb.s19(CStr(dr(0)), str1)
                                    key = Mid(med, 5, med.IndexOf("=") - 4)
                                    If med.LastIndexOf(" where " & Strings.Left(str1, Len(str1) - 2) & "批号=") > 0 Then
                                        value = Mid(med, med.IndexOf("=") + 2, med.LastIndexOf(" where " & Strings.Left(str1, Len(str1) - 2) & "批号=") - med.IndexOf("=") - 1)
                                        If value = "NULL" Then value = ""
                                        If Mid(CStr(dr(0)), CStr(dr(0)).IndexOf("=") + 2, CStr(dr(0)).LastIndexOf(" where [" & Strings.Left(str1, Len(str1) - 2) & "批号]=") - CStr(dr(0)).IndexOf("=") - 1) = "'NULL'" Then value = "NULL"
                                        If dt.ContainsKey(key) Then
                                            If value = "" Then
                                                dt.Remove(key)
                                            Else
                                                dt(key) = value
                                            End If
                                        ElseIf value <> "" Then
                                            dt.Add(key, value)
                                        End If
                                    End If
                                End If
                            End While
                            cnctm.Close()
                            bl = True
                            RemoveHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf Form2.TB_TextChanged
                            Fcsb.s32(Form2.TC2.TabPages(CInt(rg(0, i))), dt, DirectCast(rg(7, i), Dictionary(Of Control, String)), ph)
                            RemoveHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf Form2.T_TextChanged
                            DirectCast(rg(2, i), TextBox).Text = ph
                            AddHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf Form2.T_TextChanged
                            Fcsb.s34(DirectCast(rg(2, i), TextBox), CStr(rg(9, i)), rg(10, i), bl)
                            DirectCast(rg(3, i), Button).Enabled = bl
                            DirectCast(rg(4, i), Button).Enabled = bl
                            Fcsb.s40(DirectCast(rg(7, i), Dictionary(Of Control, String)), DirectCast(rg(2, i), Control).Parent, DirectCast(rg(2, i), TextBox))
                            If bl Then AddHandler DirectCast(rg(6, i), TextBox).TextChanged, AddressOf Form2.TB_TextChanged
                        Else
                            RemoveHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf Form2.T_TextChanged
                            DirectCast(rg(2, i), TextBox).Text = ph
                            AddHandler DirectCast(rg(2, i), TextBox).TextChanged, AddressOf Form2.T_TextChanged
                            Form2.s1(DirectCast(rg(2, i), TextBox), DirectCast(rg(3, i), Button), rg(10, i), DirectCast(rg(4, i), Button), DirectCast(rg(5, i), Button), DirectCast(rg(7, i), Dictionary(Of Control, String)), rg(8, i), Form2.TC2.TabPages(CInt(rg(0, i))))
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub B92_Click(sender As Object, e As EventArgs) Handles B92.Click
        For i = 1 To LI15.SelectedItems.Count
            Dim j As String = CStr(LI15.SelectedItems.Item(i - 1))
            For Each key As String In gd.Keys
                If gd(key).Contains("|" & Fcsb.s13(j) & "|") Then
                    If Not bn.ContainsKey(j) Then
                        bn.Add(j, key)
                        LI19.Items.Add(j)
                        LI19.SetSelected(LI19.Items.Count - 1, True)
                    End If
                    Exit For
                End If
            Next
        Next
        LI15.Hide()
    End Sub
    Private Sub B97_Click(sender As Object, e As EventArgs) Handles B97.Click
        If LI19.SelectedItems.Count = 0 Then bn.Clear() : LI19.Items.Clear()
        For i = 0 To LI19.SelectedItems.Count - 1
            bn.Remove(LI19.Items(LI19.SelectedIndex).ToString)
            LI19.Items.RemoveAt(LI19.SelectedIndex)
        Next
        LI19.ClearSelected()
    End Sub
    Private Sub T45_TextChanged(sender As Object, e As EventArgs) Handles T45.TextChanged
        Dim dt As New DataTable, T As TextBox = DirectCast(sender, TextBox)
        If suer = 4 OrElse suer = 5 Then Return
        LI15.SetBounds(T.Left + 22, T.Top + 66, T.Size.Width, 349)
        LI15.Font = New System.Drawing.Font("Times New Roman", 10)
        If T.Text <> "" Then
            LI15.Items.Clear()
            LI15.Show()
        Else
            LI15.Hide()
        End If
        dt.Columns.Add("物料名称")
        If CL4.CheckedItems.Count > 0 Then
            For i = 1 To CL4.CheckedItems.Count
                If CStr(CL4.CheckedItems.Item(i - 1)) <> "全部" Then dt.Rows.Add(CL4.CheckedItems.Item(i - 1))
            Next
        Else
            For i = 1 To CL4.Items.Count - 1
                dt.Rows.Add(CL4.Items.Item(i))
            Next
        End If
        Try
            cnct.Open()
            cmd = New SqlCommand("批号筛选", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("工段", dt))
            cmd.Parameters.Add(New SqlParameter("批号", T45.Text))
            dr = cmd.ExecuteReader
            While dr.Read
                LI15.Items.Add(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Public Sub DA11_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DA11.CellValueChanged
        DirectCast(sender, DataGridView).Columns(e.ColumnIndex).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub
    Private Sub DA11_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA11.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex = -1 Then
            If e.Button = Windows.Forms.MouseButtons.Middle Then
                Fcsb.s57(DA)
            ElseIf e.Button = Windows.Forms.MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        ElseIf e.ColumnIndex > -1 Then
            If nn Then
                If e.ColumnIndex = 4 Then
                    If e.Button = Windows.Forms.MouseButtons.Right Then
                        Fcsb.s25(DA, CInt(DA.Rows(e.RowIndex).Cells(4).Value), CStr(DA.Rows(e.RowIndex).Cells(3).Value))
                    End If
                End If
            End If
        End If
    End Sub
    Public Sub DA3_SelectionChanged(sender As Object, e As EventArgs) Handles DA3.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView), td As Date
        T48.Text = ""
        If DA.SelectedCells.Count <> 1 OrElse DA.SelectedCells.Item(0).ColumnIndex = 0 Then Return
        If Date.TryParse(CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex).Cells(0).Value), td) Then
            Try
                cnct.Open()
                cmd = New SqlCommand("库存回溯", cnct)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("储槽名称", DA.Columns(DA.SelectedCells.Item(0).ColumnIndex).HeaderText))
                cmd.Parameters.Add(New SqlParameter("日期", td))
                dr = cmd.ExecuteReader
                While dr.Read
                    If IsDBNull(dr(0)) Then
                        T48.Text = CStr(dr(1)) & vbCrLf & Format(dr(2), "yyMMdd HH:mm")
                    Else
                        T48.Text = CStr(dr(0)) & "  " & CStr(dr(1)) & vbCrLf & Format(dr(2), "yyMMdd HH:mm")
                    End If
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        End If
    End Sub
    Private Sub B99_Click(sender As Object, e As EventArgs) Handles B99.Click
        Dim k0 As New ArrayList, dt As New DataTable
        If LI7.Items.Count = 0 AndAlso LI8.Items.Count = 0 Then Return
        dt.Columns.Add("物料名称")
        If LI8.Items.Count = 0 Then
            For i = 1 To LI7.Items.Count
                dt.Rows.Add(LI7.Items.Item(i - 1))
            Next
        Else
            For i = 1 To LI8.Items.Count
                dt.Rows.Add(LI8.Items.Item(i - 1))
            Next
        End If
        DA2.Rows.Clear()
        D3.Checked = False : D4.Checked = True
        Dim dtm As Date = CDate(D4.Text)
        Try
            cnct.Open()
            cmd = New SqlCommand("回溯表格", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("日期", dtm))
            cmd.Parameters.Add(New SqlParameter("储槽名称", dt))
            dr = cmd.ExecuteReader
            While dr.Read
                DA2.Rows.Add()
                For i = 0 To 6
                    DA2.Rows(DA2.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                Next i
                DA2.Rows(DA2.Rows.Count - 2).Cells(1).Value = Format(DA2.Rows(DA2.Rows.Count - 2).Cells(1).Value, "yyyy-MM-dd HH:mm")
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox(ex.Message)
        End Try
        Fcsb.s14(B26, 0, DA2, idt2, Color.Pink)
        DA2.ClearSelection()
    End Sub
    Protected Overloads Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Dim DA As DataGridView
        If TypeOf ActiveControl Is DataGridViewTextBoxEditingControl Then
            DA = DirectCast(ActiveControl, DataGridViewTextBoxEditingControl).EditingControlDataGridView
        ElseIf TypeOf ActiveControl Is DataGridViewComboBoxEditingControl Then
            DA = DirectCast(ActiveControl, DataGridViewComboBoxEditingControl).EditingControlDataGridView
        Else
            Exit Function
        End If
        If keyData = Keys.Escape Then
            If DA Is DA1 Then
                RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                s45(DA, sv, "yyyy-MM-dd HH:mm", {D1, D2})
                AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
            ElseIf DA Is DA2 Then
                RemoveHandler DA.CellEndEdit, AddressOf DA2_CellEndEdit
                s45(DA, sv, "yyyy-MM-dd HH:mm", {D3, D4})
                AddHandler DA.CellEndEdit, AddressOf DA2_CellEndEdit
            ElseIf DA Is DA9 Then
                RemoveHandler DA.CellEndEdit, AddressOf DA9_CellEndEdit
                s45(DA, sv, "yyyy-MM-dd", {D5, D6, D7, D8})
                AddHandler DA.CellEndEdit, AddressOf DA9_CellEndEdit
            ElseIf DA Is DA12 Then
                RemoveHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
                s45(DA, sv)
                AddHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
            ElseIf DA Is DA11 Then
                If DA.CurrentCell.Tag IsNot Nothing Then DA.CurrentCell.Value = sv
            ElseIf DA Is DA10 Then
                RemoveHandler DA.CellEndEdit, AddressOf DA10_CellEndEdit
                DA.CancelEdit() : DA.EndEdit()
                AddHandler DA.CellEndEdit, AddressOf DA10_CellEndEdit
            Else
                RemoveHandler DA.CellEndEdit, AddressOf DA_CellEndEdit
                DA.CancelEdit() : DA.EndEdit()
                AddHandler DA.CellEndEdit, AddressOf DA_CellEndEdit
            End If
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
                If keyData = 131085 AndAlso DA Is DA10 AndAlso DA.Columns.Count = 4 Then
                    For Each cell As DataGridViewCell In DA.SelectedCells
                        If cell.ColumnIndex <> 3 Then cell.Value = ActiveControl.Text
                    Next
                End If
                If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                    DA.EndEdit()
                    Return True
                End If
            End If
        End If
    End Function
    Private Sub DA2_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA2.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 Then
            If CStr(DA.Rows(e.RowIndex).Cells(0).Value) <> "" Then
                If e.ColumnIndex = 0 Then
                    If e.Button = Windows.Forms.MouseButtons.Left AndAlso CInt(DA.Rows(e.RowIndex).Cells(0).Value) < 0 Then
                        If CStr(DA.Rows(e.RowIndex).Cells(0).Value) <> "" Then
                            Fcsb.s46(DirectCast(sender, DataGridView), e.RowIndex, False)
                        End If
                    ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
                        If e.ColumnIndex = 0 Then DA.ClearSelection() : DA.Rows(e.RowIndex).Cells(0).Selected = True
                        Dim ct As Integer = DA11.Rows.Count
                        Dim cmdstrn As Integer
                        If DA.Rows(e.RowIndex).Tag IsNot Nothing Then
                            cmdstrn = DirectCast(DA.Rows(e.RowIndex).Tag, Integer())(1)
                        ElseIf DA.Rows(e.RowIndex).Cells(0).Value IsNot Nothing Then
                            cmdstrn = CInt(DA.Rows(e.RowIndex).Cells(0).Value)
                        Else
                            Return
                        End If
                        Fcsb.s25(DA11, cmdstrn, "储槽液位")
                        TC1.SelectedIndex = 4
                    ElseIf DA.Rows.Count > 1 Then
                        If e.Button = Windows.Forms.MouseButtons.Middle AndAlso CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) <> "" AndAlso Not sbl(3) Then
                            DA.EndEdit()
                            If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                                Dim en As EventArgs
                                If B26.Text = "解锁表格" Then B26_Click(B26, en)
                                DA.Rows.Add()
                                For i = 1 To DA.Columns.Count - 1
                                    DA.Rows(DA.Rows.Count - 2).Cells(i).Value = DA.Rows(e.RowIndex).Cells(i).Value
                                Next
                                RemoveHandler DA.RowValidating, AddressOf DA2_RowValidating
                                RemoveHandler DA.SelectionChanged, AddressOf DA2_SelectionChanged
                                DA.Columns(1).Visible = True
                                DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(1)
                                DA.Rows(DA.Rows.Count - 2).ReadOnly = False
                                DA.BeginEdit(True)
                                DA.Rows(e.RowIndex).Cells(0).Selected = True
                                AddHandler DA.SelectionChanged, AddressOf DA2_SelectionChanged
                                AddHandler DA.RowValidating, AddressOf DA2_RowValidating
                                dttm = CStr(DA.Rows(e.RowIndex).Cells(1).Value)
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
                If e.Button = Windows.Forms.MouseButtons.Middle Then
                    Fcsb.s57(DA)
                ElseIf e.Button = Windows.Forms.MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                    DA.Columns.Item(e.ColumnIndex).Visible = False
                End If
            End If
        End If
    End Sub
    Public Sub DA1_SelectionChanged(sender As Object, e As EventArgs) Handles DA1.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If skip(1) Then
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
            End If
            AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
        End If
        Dim a(1, 1) As String
        Fcsb.s52(DA, a, T59)
    End Sub
    Public Sub DA2_SelectionChanged(sender As Object, e As EventArgs) Handles DA2.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If skip(1) Then
            RemoveHandler DA.CellEndEdit, AddressOf DA2_CellEndEdit
            If Not DA.CurrentCell Is dgvcell Then
                RemoveHandler DA.SelectionChanged, AddressOf DA2_SelectionChanged
                RemoveHandler DA.RowValidating, AddressOf DA2_RowValidating
                RemoveHandler DA.CellBeginEdit, AddressOf DA2_CellBeginEdit
                DA.Columns(dgvcell.ColumnIndex).Visible = True
                DA.CurrentCell = dgvcell
                DA.BeginEdit(False)
                AddHandler DA.CellBeginEdit, AddressOf DA2_CellBeginEdit
                AddHandler DA.RowValidating, AddressOf DA2_RowValidating
                AddHandler DA.SelectionChanged, AddressOf DA2_SelectionChanged
            End If
            AddHandler DA.CellEndEdit, AddressOf DA2_CellEndEdit
        End If
        Dim a(1, 1) As String
        Fcsb.s52(DA, a, T60)
    End Sub
    Public Sub DA12_SelectionChanged(sender As Object, e As EventArgs) Handles DA12.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If skip(1) Then
            RemoveHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
            If Not DA.CurrentCell Is dgvcell Then
                RemoveHandler DA.SelectionChanged, AddressOf DA12_SelectionChanged
                RemoveHandler DA.RowValidating, AddressOf DA12_RowValidating
                DA.Columns(dgvcell.ColumnIndex).Visible = True
                DA.CurrentCell = dgvcell
                RemoveHandler DA.CellBeginEdit, AddressOf DA12_CellBeginEdit
                DA.BeginEdit(False)
                AddHandler DA.CellBeginEdit, AddressOf DA12_CellBeginEdit
                AddHandler DA.RowValidating, AddressOf DA12_RowValidating
                AddHandler DA.SelectionChanged, AddressOf DA12_SelectionChanged
            End If
            AddHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
        End If
    End Sub
    Public Sub DA9_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA9.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If Not DA.IsCurrentCellInEditMode AndAlso (IsNothing(DA.CurrentCell) OrElse DA.Rows(DA.CurrentCell.RowIndex).Cells(0).Value IsNot Nothing OrElse DA.Rows(DA.CurrentCell.RowIndex).IsNewRow) Then
            If e.Button = Windows.Forms.MouseButtons.Middle Then
                Fcsb.s57(DA)
            ElseIf e.Button = Windows.Forms.MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        End If
    End Sub
    Public Sub DA9_SelectionChanged(sender As Object, e As EventArgs) Handles DA9.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If skip(1) Then
            RemoveHandler DA.CellEndEdit, AddressOf DA9_CellEndEdit
            If DA.CurrentCell IsNot dgvcell Then
                RemoveHandler DA.SelectionChanged, AddressOf DA9_SelectionChanged
                RemoveHandler DA.RowValidating, AddressOf DA9_RowValidating
                RemoveHandler DA.CellBeginEdit, AddressOf DA9_CellBeginEdit
                DA.Columns(dgvcell.ColumnIndex).Visible = True
                DA.CurrentCell = dgvcell
                DA.BeginEdit(False)
                AddHandler DA.CellBeginEdit, AddressOf DA9_CellBeginEdit
                AddHandler DA.RowValidating, AddressOf DA9_RowValidating
                AddHandler DA.SelectionChanged, AddressOf DA9_SelectionChanged
            End If
            AddHandler DA.CellEndEdit, AddressOf DA9_CellEndEdit
        End If
    End Sub
    Private Sub LB_Click(sender As Object, e As EventArgs) Handles L125.Click, L130.Click
        If Not CBool(lbl(sender)(0)) Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            lbl(sender)(0) = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        If sender Is L125 Then
            s24(True)
        Else
            s27(True)
        End If
    End Sub
    Private Sub L124_MouseClick(sender As Object, e As MouseEventArgs) Handles L124.MouseClick
        ex = e.X
        If Not L124bl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            L124bl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        s22(True, ex)
    End Sub
    Private Sub L_Click(sender As Object, e As EventArgs) Handles L126.Click, L127.Click, L128.Click
        If Not CBool(lbl(sender)(0)) Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            lbl(sender)(0) = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        Fcsb.s30(True, sender)
    End Sub
    Private Sub L39_Click(sender As Object, e As EventArgs) Handles L39.Click
        Dim L As Label = DirectCast(sender, Label)
        If L.Text = "~" Then
            L.Text = "、"
        Else
            L.Text = "~"
        End If
    End Sub
    Private Sub B103_Click(sender As Object, e As EventArgs) Handles B103.Click
        Dim blct As Boolean
        If Not bcbl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            bcbl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
            blct = True
        End If
        s40(Not blct)
    End Sub
    Private Sub B104_Click(sender As Object, e As EventArgs) Handles B104.Click
        Try
            cnct.Open()
            cmd = New SqlCommand(My.Settings.T49T, cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
            MsgBox("数据整理成功！")
        Catch ex As Exception
            cnct.Close()
            MsgBox("数据整理失败！")
        End Try
    End Sub
    Private Sub CH4_MouseUp(sender As Object, e As EventArgs) Handles CH4.MouseUp
        Dim CH As CheckBox = DirectCast(sender, CheckBox)
        If CH.Text = "数据库值" Then
            CH.Text = "月初转存"
        ElseIf CH.Text = "月初转存" Then
            CH.Text = "用户输入"
        ElseIf CH.Text = "用户输入" Then
            CH.Text = "智能盘存"
        ElseIf CH.Text = "智能盘存" Then
            CH.Text = "数据库值"
        End If
        CH.Checked = True
    End Sub
    Private Sub L_TextChanged(sender As Object, e As EventArgs) Handles L40.TextChanged, L41.TextChanged
        If flg Then
            Dim TB As TextBox = DirectCast(sender, TextBox), str As String = CStr(IIf(TB Is L40, "物料名称", "项目")), CL As CheckedListBox = DirectCast(IIf(TB Is L40, CL1, CL3), CheckedListBox)
            CL.Items.Clear()
            CL.Items.Add(ary(sender).Rows(0)(1))
            For Each row As DataRow In ary(sender).Select(str & " like '%" & Replace(TB.Text, "'", "''") & "%'", "Id Asc")
                If CInt(row(0)) > 0 Then CL.Items.Add(row(1))
            Next
            If CL.Items.Count = 2 Then CL.SetItemChecked(0, True)
        End If
    End Sub
    Private Sub LM_LostFocus(sender As Object, e As EventArgs) Handles L40.LostFocus, L41.LostFocus
        Dim TB As TextBox = DirectCast(sender, TextBox), str As String = CStr(IIf(TB Is L40, "物料名称", "项目"))
        If TB.Text = "" Then
            RemoveHandler TB.TextChanged, AddressOf L_TextChanged
            TB.Text = str
            AddHandler TB.TextChanged, AddressOf L_TextChanged
        End If
    End Sub
    Private Sub LM_GotFocus(sender As Object, e As EventArgs) Handles L40.GotFocus, L41.GotFocus
        Dim TB As TextBox = DirectCast(sender, TextBox), str As String = CStr(IIf(TB Is L40, "物料名称", "项目")), CL As CheckedListBox = DirectCast(IIf(TB Is L40, CL1, CL3), CheckedListBox)
        If TB.Text = str Then
            ary(sender).Reset()
            ary(sender).Columns.Add("Id", Type.GetType("System.Int32"))
            ary(sender).Columns.Add(str)
            For i = 0 To CL.Items.Count - 1
                ary(sender).Rows.Add(i, CL.Items(i))
            Next
            RemoveHandler TB.TextChanged, AddressOf L_TextChanged
            TB.Text = ""
            AddHandler TB.TextChanged, AddressOf L_TextChanged
        End If
    End Sub
    Private Sub B109_Click(sender As Object, e As EventArgs) Handles B109.Click
        bbbl = True
        PB.Visible = False
        DirectCast(sender, Button).Visible = False
    End Sub
    Private Sub T50_GotFocus(sender As Object, e As EventArgs) Handles T50.GotFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        T.SelectionLength = 0
        If skip(0) Then
            RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
            RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
            If skip(1) Then DA1.Columns(dgvcell.ColumnIndex).Visible = True
            DA1.Select()
            If skip(1) Then DA1.CurrentCell = dgvcell
            RemoveHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
            DA1.BeginEdit(False)
            AddHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
            AddHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
            AddHandler DA1.RowValidating, AddressOf DA1_RowValidating
        ElseIf T.Text = "物料名称：" Then
            T.Text = ""
        End If
    End Sub
    Private Sub T50_LostFocus(sender As Object, e As EventArgs) Handles T50.LostFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        RemoveHandler T.TextChanged, AddressOf T50_TextChanged
        If T.Text = "" Then T.Text = "物料名称："
        AddHandler T.TextChanged, AddressOf T50_TextChanged
    End Sub
    Private Sub T50_TextChanged(sender As Object, e As EventArgs) Handles T50.TextChanged
        If dtn.Columns.Count = 0 Then Return
        LI1.Items.Clear()
        Dim dtr() As DataRow
        dtr = dtn.Select("物料名称 like '%" & Replace(T50.Text, "'", "''") & "%'", "Id")
        For i = 0 To dtr.Count - 1
            LI1.Items.Add(dtr(i)(0))
        Next
    End Sub
    Private Sub LI19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LI19.SelectedIndexChanged
        L134.Text = "高亮显示 " & DirectCast(sender, ListBox).SelectedItems.Count & " 条"
    End Sub
    Private Sub T51_GotFocus(sender As Object, e As EventArgs) Handles T51.GotFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        T.SelectionLength = 0
        If T.Text = "储槽名称：" AndAlso Not skip(0) Then T.Text = ""
    End Sub
    Private Sub T51_LostFocus(sender As Object, e As EventArgs) Handles T51.LostFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        RemoveHandler T.TextChanged, AddressOf T51_TextChanged
        If T.Text = "" Then T.Text = "储槽名称："
        AddHandler T.TextChanged, AddressOf T51_TextChanged
    End Sub
    Private Sub T51_TextChanged(sender As Object, e As EventArgs) Handles T51.TextChanged
        Dim T As TextBox = DirectCast(sender, TextBox)
        If dto.Columns.Count = 0 Then Return
        LI7.Items.Clear()
        Dim dtr() As DataRow
        dtr = dto.Select("储槽名称 like '%" & Replace(T.Text, "'", "''") & "%' or 位号 like '%" & Replace(T.Text, "'", "''") & "%'", "Id")
        For i = 0 To dtr.Count - 1
            LI7.Items.Add(dtr(i)(0))
        Next
    End Sub
    Public Sub B105_Click(sender As Object, e As EventArgs) Handles B105.Click
        If Not b105bl Then
            ni = 0
            TM1.Interval = SystemInformation.DoubleClickTime
            TM1.Enabled = True
            b105bl = True
            Return
        Else
            TM1.Interval = 1
            TM1.Enabled = False
        End If
        s46(True)
    End Sub
    Private Sub B106_Click(sender As Object, e As EventArgs) Handles B106.Click
        Dim dgv As New DataGridViewComboBoxColumn
        If DirectCast(sender, Button).Text = "储槽表格" Then
            DirectCast(sender, Button).Text = "解锁表格"
            DA12.Columns.Clear()
            DA12.Columns.Add("", "Id") : DA12.Columns(0).Width = 45
            DA12.Columns.Add("", "日期") : DA12.Columns(1).Width = 130
            dgv.HeaderText = "储槽名称" : dgv.Width = 177 : DA12.Columns.Add(dgv) : dgv.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing : dgv.SortMode = DataGridViewColumnSortMode.Automatic
            DA12.Columns.Add("", "含量") : DA12.Columns(3).Width = 70
            DA12.Columns.Add("", "比重") : DA12.Columns(4).Width = 70
            cnct.Open()
            Fcsb.s6("select 储槽名称 from 储槽特性 order by id", DirectCast(DA12.Columns(2), DataGridViewComboBoxColumn))
            cnct.Close()
            B90.Text = "物料表格"
            B85.Enabled = False
            DA12.ReadOnly = True
        Else
            Fcsb.s7(DirectCast(sender, Button), B107, DA12)
        End If
    End Sub
    Private Sub B107_Click(sender As Object, e As EventArgs) Handles B107.Click
        Fcsb.s4(DA12, "储槽物料")
    End Sub
    Private Sub TSMI1_Click(sender As Object, e As EventArgs) Handles TSMI1.Click
        Dim i As Integer = 1, bl As Boolean
        If CMS1.SourceControl Is DA1 Then
            If Not DA1.ReadOnly OrElse suer = 4 OrElse suer = 5 Then Return
            Try
                Do Until i > DA1.SelectedCells.Count
                    bl = False
                    If CStr(DA1.SelectedCells.Item(i - 1).Value) <> "" Then
                        If DA1.SelectedCells.Item(i - 1).ColumnIndex = 2 Then
                            For Each key As String In gd.Keys
                                If gd(key).Contains("|" & CStr(DA1.Rows(DA1.SelectedCells.Item(i - 1).RowIndex).Cells(8).Value) & "|") Then
                                    If Not LI19.Items.Contains(DA1.SelectedCells.Item(i - 1).Value) Then
                                        LI19.Items.Add(DA1.SelectedCells.Item(i - 1).Value)
                                        bn.Add(CStr(DA1.SelectedCells.Item(i - 1).Value), key)
                                        LI19.SetSelected(LI19.Items.Count - 1, True)
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                    i += 1
                Loop
            Catch ex As Exception
            End Try
        Else
            For Each key As String In gd.Keys
                If gd(key).Contains("|" & Fcsb.s13(DirectCast(CMS1.SourceControl, TextBox).Text) & "|") Then
                    If Not LI19.Items.Contains(DirectCast(CMS1.SourceControl, TextBox).Text) Then
                        LI19.Items.Add(DirectCast(CMS1.SourceControl, TextBox).Text)
                        bn.Add(DirectCast(CMS1.SourceControl, TextBox).Text, key)
                        LI19.SetSelected(LI19.Items.Count - 1, True)
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub
    Private Sub CMS2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles CMS2.Opening
        Dim str() As String = {"盘存日起", "最近30天"}, str1() As String = {"绝对值", "相对值"}, str2() As String = {"实际值", "设置值", "备忘值"}
        Dim fg As Byte, num As Integer, i(2) As Integer, dt As DataTable, da As SqlDataAdapter, ary As New List(Of Integer), CMS As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        Try
            DA10.EndEdit()
            CMS.Items.Clear()
            If DA10.SelectedCells(0).ColumnIndex = 2 Then
                For Each item As DataGridViewCell In DA10.SelectedCells
                    ary.Add(item.RowIndex)
                Next
                cnct.Open()
                cmdstr = "select 类型标记 from 统计类型 where 统计类型=@统计类型"
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.Parameters.AddWithValue("统计类型", CStr(DA10.SelectedCells(0).Value))
                fg = CByte(cmd.ExecuteScalar)
                cnct.Close()
                If fg = 0 OrElse Not CL1.Items.Contains(CStr(DA10.Rows(DA10.SelectedCells(0).RowIndex).Cells(1).Value)) OrElse CStr(DA10.Rows(DA10.SelectedCells(0).RowIndex).Cells(1).Value) = "全部" Then
                    e.Cancel = True
                    DA10.Rows(DA10.SelectedCells(0).RowIndex).Cells(3).Value = Nothing
                Else
                    e.Cancel = False
                    Select Case fg
                        Case 1
                            s56(e)
                        Case 2
                            s56(e, False)
                        Case 3
                            s57(e, True)
                        Case 4
                            s56(e, True)
                        Case 5
                            s49(CMS, ary, num)
                        Case 7
                            s56(e, True)
                        Case 8
                            s50(CMS, ary, str)
                        Case 9
                            s51(CMS, ary, num)
                        Case 10
                            s52(CMS, ary, num, dt, i)
                        Case 11
                            s51(CMS, ary, num)
                        Case 12
                            s54(CMS, ary, num, dt, i, "年消耗标记")
                        Case 13
                            s54(CMS, ary, num, dt, i, "日消耗标记")
                        Case 14
                            s54(CMS, ary, num, dt, i, "累计消耗标记")
                        Case 15
                            s52(CMS, ary, num, dt, i)
                        Case 16
                            s52(CMS, ary, num, dt, i)
                        Case 17
                            s57(e, False)
                        Case 18
                            s57(e, False)
                        Case 19
                            s57(e, True)
                        Case 20
                            s50(CMS, ary, str1)
                        Case 21
                            s50(CMS, ary, str1)
                        Case 22
                            s53(CMS, ary, str2)
                        Case 23
                            s53(CMS, ary, str2)
                        Case 24
                            s50(CMS, ary, str1)
                        Case 25
                            s57(e, True)
                        Case 26
                            s51(CMS, ary, num)
                    End Select
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            cnct.Close()
            e.Cancel = True
            For Each cell As DataGridViewCell In DA10.SelectedCells
                DA10.Rows(cell.RowIndex).Cells(3).Value = Nothing
            Next
        End Try
    End Sub
    Private Sub TSMI_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles TSMIA.DropDownItemClicked, TSMIB.DropDownItemClicked, TSMIC.DropDownItemClicked, TSMID.DropDownItemClicked, CMS2.ItemClicked
        If Not DirectCast(e.ClickedItem, ToolStripMenuItem).HasDropDownItems Then TSMI = DirectCast(e.ClickedItem, ToolStripMenuItem)
    End Sub
    Private Sub TSMI_Click(sender As Object, e As EventArgs) Handles TSMI.Click
        Dim fg As Byte, TSMIM As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Try
            If DA10.SelectedCells.Count > 0 Then
                For Each current As DataGridViewCell In DA10.SelectedCells
                    If current.ColumnIndex = 2 Then
                        Dim str As String = CStr(DA10.Rows(current.RowIndex).Cells(0).Value)
                        Dim str1 As String = CStr(DA10.Rows(current.RowIndex).Cells(1).Value)
                        Dim str2 As String = CStr(DA10.Rows(current.RowIndex).Cells(2).Value)
                        If CL1.Items.Contains(CStr(DA10.Rows(current.RowIndex).Cells(1).Value)) AndAlso CStr(DA10.Rows(current.RowIndex).Cells(1).Value) <> "全部" Then
                            cnct.Open()
                            cmdstr = "select 类型标记 from 统计类型 where 统计类型=@统计类型"
                            cmd = New SqlCommand(cmdstr, cnct)
                            cmd.Parameters.AddWithValue("统计类型", CStr(DA10.SelectedCells(0).Value))
                            fg = CByte(cmd.ExecuteScalar)
                            Select Case fg
                                Case 5
                                    s17(current, TSMIM, str2, str1, str, ",@盘存模式", ",@消耗标记")
                                Case 8
                                    s23(current, TSMIM, str2, "@单耗标记")
                                Case 9
                                    s10(str, str1, str2, TSMIM, current)
                                Case 10
                                    s3(str, str1, str2, TSMIM, current)
                                Case 11
                                    s48(str, str1, str2, TSMIM, current)
                                Case 12
                                    s32(str2, str1, str, TSMIM, current, "@年消耗标记")
                                Case 13
                                    s32(str2, str1, str, TSMIM, current, "@日消耗标记")
                                Case 14
                                    s32(str2, str1, str, TSMIM, current, "@累计消耗标记")
                                Case 15
                                    s3(str, str1, str2, TSMIM, current)
                                Case 16
                                    s3(str, str1, str2, TSMIM, current)
                                Case 20
                                    s17(current, TSMIM, str2, str1, str,, ",@模式")
                                Case 19
                                    s17(current, TSMIM, str2, str1, str)
                                Case 21
                                    s17(current, TSMIM, str2, str1, str, ",@盘存模式", ",@模式")
                                Case 22
                                    s17(current, TSMIM, str2, str1, str,, ",@" & str2 & "标记")
                                Case 23
                                    s17(current, TSMIM, str2, str1, str, ",@盘存模式", ",@" & str2 & "标记")
                                Case 24
                                    s17(current, TSMIM, str2, str1, str, ",@盘存模式", ",@" & str2 & "标记")
                                Case 26
                                    s10(str, str1, str2, TSMIM, current)
                            End Select
                            cnct.Close()
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub CMS4_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles CMS4.ItemClicked
        TSMIN = DirectCast(e.ClickedItem, ToolStripMenuItem)
    End Sub
    Private Sub CMS4_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles CMS4.Opening
        Dim CMS As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        CMS.Items.Clear()
        cnct.Open()
        cmd = New SqlCommand("select 物料类型 from 物料类型 order by id", cnct)
        dr = cmd.ExecuteReader
        While dr.Read
            CMS.Items.Add(CStr(dr(0)))
            If LI3.Items.Contains(dr(0)) OrElse LI4.Items.Contains(dr(0)) Then
                TSMIN = DirectCast(CMS.Items(CMS.Items.Count - 1), ToolStripMenuItem)
                TSMIN.CheckState = CheckState.Checked
            End If
        End While
        cnct.Close()
        LI3.Items.Clear()
        LI4.Items.Clear()
        For Each item As ToolStripMenuItem In CMS.Items
            If item.CheckState = CheckState.Checked Then
                LI3.Items.Add(item.ToString)
            End If
        Next
        If Not b105bl Then e.Cancel = Not sbl(0)
    End Sub
    Private Sub TSMIN_Click(sender As Object, e As EventArgs) Handles TSMIN.Click
        If TSMIN.CheckState = CheckState.Unchecked Then
            LI3.Items.Add(TSMIN.Text)
        Else
            LI3.Items.Remove(TSMIN.Text)
        End If
        LI1.Items.Clear()
        LI2.Items.Clear()
        tb1.Reset()
        tb1.Columns.Add("物料名称")
        For Each r In LI3.Items
            tb1.Rows.Add(r)
        Next
        Try
            DA1.Rows.Clear()
            DirectCast(DA1.Columns(6), DataGridViewComboBoxColumn).Items.Clear()
            cnct.Open()
            dr = New SqlCommand("select 物料类型 from 物料类型 order by id", cnct).ExecuteReader
            While dr.Read
                If LI3.Items.Contains(dr(0)) Then DirectCast(DA1.Columns.Item(6), DataGridViewComboBoxColumn).Items.Add(dr(0))
            End While
            dr.Close()
            Fcsb.s29()
            s7(tb1, tb2, clbl AndAlso CL2.CheckedItems.Count = 0)
            s15(tb1, tb2)
            cnct.Close()
            DirectCast(DA1.Columns.Item(6), DataGridViewComboBoxColumn).Items.Add("")
        Catch ex As Exception
            cnct.Close()
        End Try
        s59(Form2 IsNot Nothing, False)
    End Sub
    Private Sub TC1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TC1.Selecting
        LI15.Hide()
        If skip(0) Then
            Select Case tci
                Case 0
                    RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
                    RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
                Case 1
                    RemoveHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
                    RemoveHandler DA2.RowValidating, AddressOf DA2_RowValidating
                Case 3
                    RemoveHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
                    RemoveHandler DA9.RowValidating, AddressOf DA9_RowValidating
                Case 5
                    AddHandler DA12.GotFocus, AddressOf DA12_GotFocus
            End Select
        End If
        e.Cancel = skip(0) OrElse suer = 4 AndAlso e.TabPageIndex = 3 OrElse suer = 6 AndAlso e.TabPageIndex > 0 AndAlso e.TabPageIndex < 4
        If Not e.Cancel Then tci = e.TabPageIndex
    End Sub
    Public Sub DAM_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseUp, DA2.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.SelectedCells.Count = 0 OrElse DA.SelectedRows.Count > 0 Then Return
        If DA.SelectedCells.Count = 1 AndAlso e.RowIndex > -1 AndAlso e.ColumnIndex > -1 Then
            DA.Columns(DA.SelectedCells(0).ColumnIndex).Visible = True
            If e.Button = Windows.Forms.MouseButtons.Left Then
                DA.BeginEdit(True)
            End If
        End If
        Dim T(5) As TextBox
        If sender Is DA1 Then
            T(1) = L10 : T(2) = L12
            T(3) = L8 : T(4) = L27
            T(5) = L48
        ElseIf sender Is DA2 Then
            T(1) = L18 : T(2) = L20
            T(3) = L16 : T(4) = L85
            T(5) = L87
        End If
        Try
            s4(DA, e.ColumnIndex, T(1), T(2), T(3), T(4), T(5))
        Catch ex As Exception
        End Try
    End Sub
    Private Sub L137_Click(sender As Object, e As EventArgs) Handles L137.Click
        s12({T53, T52}, "0.000")
    End Sub
    Private Sub L108_Click(sender As Object, e As EventArgs) Handles L108.Click
        s12({T40, T39}, "0.00")
    End Sub
    Public Sub D_Change(sender As Object, e As EventArgs) Handles D1.MouseUp, D2.MouseUp, D1.ValueChanged, D2.ValueChanged
        T38_TextChanged(T38, e)
    End Sub
    Private Sub T28_KeyDown(sender As Object, e As KeyEventArgs) Handles T28.KeyDown
        Dim T As TextBox = DirectCast(sender, TextBox)
        If sbl(0) Then
            Dim i, j As Integer
            If e.Control AndAlso e.KeyCode = Keys.Right Then
                i = 1
            ElseIf e.Control AndAlso e.KeyCode = Keys.Left Then
                i = -1
            End If
            If Math.Abs(i) = 1 Then
                Dim ds As New DataSet
                cmd = New SqlCommand("语句查询", cnct)
                cmd.CommandType = CommandType.StoredProcedure
                da = New SqlDataAdapter(cmd)
                da.Fill(ds)
                Do
                    If T.Text = "" Then
                        T.Text = ds.Tables(CInt((ds.Tables.Count - 1) / 2 - i / 2 * (ds.Tables.Count - 1))).Rows(0)(0).ToString
                        Return
                    ElseIf T.Text = ds.Tables(j).Rows(0)(0).ToString Then
                        If j = CInt((ds.Tables.Count - 1) / 2 + i / 2 * (ds.Tables.Count - 1)) Then
                            T.Text = ""
                        Else
                            T.Text = ds.Tables(j + i).Rows(0)(0).ToString
                        End If
                        Return
                    ElseIf j < ds.Tables.Count - 1 Then
                        j += 1
                    Else
                        T.Text = ""
                        Return
                    End If
                Loop
            End If
        End If
    End Sub
    Private Sub T1_TextChanged(sender As Object, e As EventArgs) Handles T1.TextChanged
        If (IsNumeric(T1.Text) OrElse T1.Text = "") AndAlso suer <> 4 Then
            B13.Enabled = True : B13.Text = "序号查询" : AcceptButton = B13
        Else
            B13.Enabled = False : B13.Text = "班别班组" : AcceptButton = B14
        End If
        If T1.Text = "" Then AcceptButton = B14
    End Sub
    Private Sub L42_Click(sender As Object, e As EventArgs) Handles L42.Click
        Dim L As Label = DirectCast(sender, Label)
        If R16.Checked Then
            If L.Text = "~" Then
                L.Text = "、"
            Else
                L.Text = "~"
            End If
        End If
    End Sub
    Private Sub L_LostFocus(sender As Object, e As EventArgs) Handles L8.LostFocus, L10.LostFocus, L12.LostFocus, L27.LostFocus, L48.LostFocus, L16.LostFocus, L18.LostFocus, L85.LostFocus, L87.LostFocus, L20.LostFocus, L80.LostFocus, L82.LostFocus, L96.LostFocus, L98.LostFocus, L84.LostFocus, T59.LostFocus, T1.LostFocus
        AcceptButton = B14
    End Sub
    Private Sub CMS1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles CMS1.Opening
        Dim CMS As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        If CMS.SourceControl Is DA1 Then e.Cancel = B17.Enabled OrElse suer = 4 OrElse suer = 5
    End Sub
    Private Sub T_KeyUp(sender As Object, e As KeyEventArgs) Handles T46.KeyUp, T47.KeyUp
        Dim T As TextBox = DirectCast(sender, TextBox)
        RemoveHandler T.TextChanged, AddressOf T_TextChanged
        Fcsb.s51(T, e, False)
        AddHandler T.TextChanged, AddressOf T_TextChanged
        s16(DirectCast(sender, TextBox))
    End Sub
    Private Sub LT_Keyup(sender As Object, e As KeyEventArgs) Handles L8.KeyUp, L10.KeyUp, L12.KeyUp, L27.KeyUp, L48.KeyUp, L16.KeyUp, L18.KeyUp, L85.KeyUp, L87.KeyUp, L20.KeyUp, L80.KeyUp, L82.KeyUp, L96.KeyUp, L98.KeyUp, L84.KeyUp, T54.KeyUp, T55.KeyUp, T56.KeyUp, T57.KeyUp, T58.KeyUp, T59.KeyUp, T39.KeyUp, T40.KeyUp, T52.KeyUp, T53.KeyUp
        RemoveHandler DirectCast(sender, TextBox).TextChanged, AddressOf LT_TextChanged
        Fcsb.s51(DirectCast(sender, TextBox), e, True)
        AddHandler DirectCast(sender, TextBox).TextChanged, AddressOf LT_TextChanged
    End Sub
    Public Sub LT_TextChanged(sender As Object, e As EventArgs) Handles L8.TextChanged, L10.TextChanged, L12.TextChanged, L27.TextChanged, L48.TextChanged, L16.TextChanged, L18.TextChanged, L85.TextChanged, L87.TextChanged, L20.TextChanged, L80.TextChanged, L82.TextChanged, L96.TextChanged, L98.TextChanged, L84.TextChanged, T54.TextChanged, T55.TextChanged, T56.TextChanged, T57.TextChanged, T58.TextChanged, T59.TextChanged, T39.TextChanged, T40.TextChanged, T52.TextChanged, T53.TextChanged
        DirectCast(sender, Control).Tag = DirectCast(sender, Control).Text
    End Sub
    Private Sub LT_GotFocus(sender As Object, e As EventArgs) Handles L8.GotFocus, L10.GotFocus, L12.GotFocus, L27.GotFocus, L48.GotFocus, L16.GotFocus, L18.GotFocus, L85.GotFocus, L87.GotFocus, L20.GotFocus, L80.GotFocus, L82.GotFocus, L96.GotFocus, L98.GotFocus, L84.GotFocus, T54.GotFocus, T55.GotFocus, T56.GotFocus, T57.GotFocus, T58.GotFocus, T59.GotFocus, T39.GotFocus, T40.GotFocus, T52.GotFocus, T53.GotFocus
        AcceptButton = Nothing
    End Sub
    Public Sub T_TextChanged(sender As Object, e As EventArgs) Handles T46.TextChanged, T47.TextChanged
        If CO8.Text = "" Then Return
        s16(DirectCast(sender, TextBox))
    End Sub
    Private Sub CT_TextChanged(sender As Object, e As EventArgs) Handles CO8.TextChanged, D11.ValueChanged, CB6.SelectedIndexChanged
        If sender Is D11 OrElse sender Is CB6 Then
            If CB6.Text <> "" Then
                T62.Text = "班别(以" & CB6.Text & "为准):" & Fcsb.s24(,, CB6.Text, D11)
            ElseIf UBound(st) = 4 Then
                T62.Text = "班别(以" & st(4) & "为准):" & Fcsb.s24(,, st(4), D11)
            Else
                T62.Text = ""
            End If
        End If
        If CO8.Text <> "" AndAlso cbl IsNot Nothing Then s16(cbl)
    End Sub
    Private Sub T_GotFocus(sender As Object, e As EventArgs) Handles T46.GotFocus, T47.GotFocus
        cbl = DirectCast(sender, TextBox)
    End Sub
    Public Sub 原料动态表(ByRef dt As Date, xlsheet As Worksheet, ByRef dat As DataTable, ByRef tad As DataTable)
        Dim dtm As New DataTable
        Try
            cmd = New SqlCommand("原料动态", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("日期", dt))
            cmd.Parameters.Add(New SqlParameter("物料名称", dat))
            cmd.Parameters.Add(New SqlParameter("盘存类型", Fcsb.s23()))
            da = New SqlDataAdapter(cmd)
            da.Fill(dtm)
            xlsheet.Cells.ImportDataTable(dtm, True, "A3")
            cmd = New SqlCommand("select dbo.盘存日期(DEFAULT,'" & dt & "', @盘存类型)", cnct)
            cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
            cnct.Open()
            xlsheet.Cells(1, 12).Value = Format(cmd.ExecuteScalar, "盘存日:yyyy-MM-dd HH:mm")
            cnct.Close()
            xlsheet.Cells(1, 15).Value = Format(dt, "日期:yyyy-MM-dd 周ddd")
            Dim k As Integer
            Dim i As Integer = dtm.Rows.Count
            Dim xlcell As Cells
            Dim st1 As New Style
            Dim st2 As New Style
            xlcell = xlsheet.Cells
            st1.VerticalAlignment = TextAlignmentType.Center
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st1.Font.Name = "Times New Roman"
            st1.Font.Size = 11
            st1.Custom = "0.000"
            st1.HorizontalAlignment = TextAlignmentType.Right
            xlcell.CreateRange("C4:M" & i + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("O4:O" & i + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.HorizontalAlignment = TextAlignmentType.Center
            For j = 4 To i + 3
                If CSng(xlcell.CreateRange("N" & j).Value) < 10 Then
                    st1.Custom = "0.0"
                ElseIf CSng(xlcell.CreateRange("N" & j).Value) < 100 Then
                    st1.Custom = "0."
                Else
                    st1.Custom = "0"
                End If
                xlcell.CreateRange("N" & j).ApplyStyle(st1, New StyleFlag With {.All = True})
            Next
            st1.Custom = "0"
            xlcell.CreateRange("P4:Q" & i + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("A4:A" & i + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "楷体"
            xlcell.CreateRange("B4:B" & i + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "微软雅黑"
            st1.Font.Size = 11
            xlcell.CreateRange("A3:Q3").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.Merge(0, 0, 1, 18)
            st1.Font.Name = "宋体"
            st1.Font.Size = 16
            st1.Font.IsBold = True
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
            xlcell(0, 0).SetStyle(st1)
            st1.Font.IsBold = False
            st1.Font.Name = "仿宋"
            st1.Font.Size = 11
            st1.IsTextWrapped = True
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Double
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.None
            st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            xlcell.Merge(i + 3, 0, 1, 18)
            st1.HorizontalAlignment = TextAlignmentType.Left
            st1.VerticalAlignment = TextAlignmentType.Top
            xlcell.CreateRange("A" & i + 4 & ":Q" & i + 4).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.Merge(1, 0, 1, 11)
            xlcell.Merge(1, 11, 1, 3)
            xlcell.Merge(1, 14, 1, 3)
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.None
            st1.IsTextWrapped = False
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.VerticalAlignment = TextAlignmentType.Center
            xlcell.CreateRange("A2:Q2").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.SetRowHeight(0, 40.56)
            xlcell.SetRowHeight(1, 21.67)
            xlcell.SetRowHeight(2, 25)
            xlcell.CreateRange("C3:D3").ColumnWidth = 8.57
            xlcell.CreateRange("E3:Q3").ColumnWidth = 9
            xlcell.SetColumnWidth(0, 5)
            xlcell.SetColumnWidth(1, 9)
            If xlcell.MaxDataRow < 23 Then
                With xlsheet.PageSetup
                    .TopMargin = 0.8
                    .RightMargin = 0.8
                    .LeftMargin = 0.8
                    .BottomMargin = 0.8
                    .HeaderMargin = 0
                    .FooterMargin = 0
                    .CenterHorizontally = True
                    .CenterVertically = False
                    .Orientation = PageOrientationType.Landscape
                    .PaperSize = PaperSizeType.PaperA4
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With
                xlcell.SetRowHeight(i + 3, 55.56)
                xlcell.CreateRange("A4:Q" & i + 3).RowHeight = 25
            Else
                With xlsheet.PageSetup
                    .TopMargin = 1.5
                    .RightMargin = 1.5
                    .LeftMargin = 1.5
                    .BottomMargin = 1.5
                    .HeaderMargin = 0
                    .FooterMargin = 0
                    .CenterHorizontally = True
                    .CenterVertically = False
                    .Orientation = PageOrientationType.Portrait
                    .PaperSize = PaperSizeType.PaperA3
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With
                xlcell.SetRowHeight(i + 3, 90)
                xlcell.CreateRange("A4:Q" & i + 3).RowHeight = Math.Max(Math.Min(1150 / i, 25), 20)
            End If
            st2.ForegroundColor = Color.FromArgb(255, 128, 128)
            st2.Pattern = BackgroundType.Solid
            st2.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            st2.VerticalAlignment = TextAlignmentType.Center
            st2.Font.Name = "Times New Roman"
            st2.Font.Size = 11
            st2.Custom = "0.000"
            For k = 3 To i + 2
                If CInt(xlcell(k, 17).Value) < 0 Then
                    xlcell(k, 10).SetStyle(st2)
                End If
            Next
            st2.ForegroundColor = xlcell(i + 2, 10).GetStyle.ForegroundColor
            st2.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Double
            xlcell(i + 2, 10).SetStyle(st2)
            st2.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            xlcell.DeleteColumn(17)
            s18(xlsheet, dtm.Rows.Count + 3, 0, "原料动态表", dt)
        Catch ex As Exception
            cnct.Close()
            MsgBox(Format(dt, "yyyy-MM-dd") & "原料动态表生成发生错误！" & vbCrLf & ex.Message)
            Return
        End Try
    End Sub
    Public Sub 平均核算表(ByRef dt As Date, xlsheet As Worksheet, ByRef dat As DataTable, ByRef tad As DataTable)
        Dim dtm As New DataTable
        Try
            cmd = New SqlCommand("平均核算", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("日期", dt))
            cmd.Parameters.Add(New SqlParameter("物料名称", dat))
            cmd.Parameters.Add(New SqlParameter("操作工序", tad))
            da = New SqlDataAdapter(cmd)
            da.Fill(dtm)
            xlsheet.Cells.ImportDataTable(dtm, True, "A3")
            Dim l As Integer = 3
            Dim k As Integer = dtm.Rows.Count
            Dim a() As Integer
            ReDim a(0)
            Dim st1 As New Style
            Dim st2 As New Style
            Dim xlcell As Cells
            xlcell = xlsheet.Cells
            If k > 1 Then
                Do
                    If CDec(xlcell(l, 13).Value) <> CDec(xlcell(l + 1, 13).Value) Then
                        a(UBound(a)) = l
                        ReDim Preserve a(UBound(a) + 1)
                    End If
                    l = l + 1
                    If l = k + 2 Then
                        a(UBound(a)) = l
                        Exit Do
                    End If
                Loop
            End If
            If a(0) > 3 Then
                For j = 5 To 7
                    xlcell.Merge(3, j, a(0) - 2, 1)
                Next
                For i = 1 To UBound(a)
                    If a(i) - a(i - 1) > 1 Then
                        For j = 5 To 7
                            xlcell.Merge(a(i - 1) + 1, j, a(i) - a(i - 1), 1)
                        Next
                    End If
                Next
            End If
            xlcell.DeleteColumn(13)
            xlcell.Merge(0, 0, 1, 15)
            xlcell.Merge(1, 0, 1, 12)
            st2.Font.Name = "仿宋"
            st2.Font.Size = 11
            st2.HorizontalAlignment = TextAlignmentType.Left
            st2.VerticalAlignment = TextAlignmentType.Center
            xlcell.CreateRange(1, 0, 1, 9).ApplyStyle(st2, New StyleFlag With {.All = True})
            st1.Font.Name = "宋体"
            st1.Font.Size = 14
            st1.Font.IsBold = True
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.VerticalAlignment = TextAlignmentType.Center
            st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Double
            xlcell.CreateRange("A1:O1").ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "Times New Roman"
            st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st1.Font.Size = 12
            st1.Custom = "0.000"
            st1.Font.IsBold = False
            st1.HorizontalAlignment = TextAlignmentType.Right
            st1.VerticalAlignment = TextAlignmentType.Center
            xlcell.CreateRange("C4:M" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.Custom = "0"
            xlcell.CreateRange("A4:A" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "楷体"
            xlcell.CreateRange("B4:B" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "微软雅黑"
            st1.Font.Size = 12
            xlcell.CreateRange("A3:O3").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("N4:O" & k + 3).Merge()
            st1.Font.Name = "仿宋"
            xlcell.CreateRange("N4:O" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("N2:O2").Merge()
            xlcell("M2").Value = "日期"
            xlcell("N3").Value = "备注"
            xlcell.CreateRange("N3:O3").Merge()
            st1.Font.Name = "仿宋"
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.None
            st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.None
            xlcell.CreateRange("M2:O2").ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.HorizontalAlignment = TextAlignmentType.Left
            st1.VerticalAlignment = TextAlignmentType.Top
            st1.IsTextWrapped = True
            xlcell("N4").SetStyle(st1)
            xlcell.SetColumnWidth(0, 5.42)
            xlcell.SetColumnWidth(1, 9.87)
            xlcell.SetColumnWidth(10, 9.87)
            xlcell.CreateRange("C:C").ColumnWidth = 10.86
            xlcell.CreateRange("D:D").ColumnWidth = 12.14
            xlcell.CreateRange("E:E").ColumnWidth = 13.43
            xlcell.CreateRange("I:I").ColumnWidth = 12.14
            xlcell.CreateRange("J:J").ColumnWidth = 10.86
            xlcell.CreateRange("K:K").ColumnWidth = 12.14
            xlcell.CreateRange("F:H").ColumnWidth = 8.75
            xlcell.CreateRange("N:O").ColumnWidth = 9.31
            xlcell.CreateRange("L:M").ColumnWidth = 12.22
            With xlsheet.PageSetup
                .LeftMargin = 0.8
                .RightMargin = 0.8
                .TopMargin = 0.8
                .BottomMargin = 0.8
                .HeaderMargin = 0
                .FooterMargin = 0
                .CenterHorizontally = True
                .CenterVertically = False
                .Orientation = PageOrientationType.Landscape
                .PaperSize = PaperSizeType.PaperA4
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
            xlcell.CreateRange("4:" & k + 3).RowHeight = Math.Max(Math.Min(625 / k, 30), 20)
            xlcell.CreateRange("1:1").RowHeight = 21.5
            xlcell.CreateRange("2:2").RowHeight = 14.25
            xlcell.CreateRange("3:3").RowHeight = 17.25
            xlsheet.Cells(1, 13).Value = Format(dt, "yyyy-MM-dd")
            s18(xlsheet, 3, 13, "平均核算表", dt)
        Catch ex As Exception
            MsgBox(Format(dt, "yyyy-MM-dd") & "平均核算表生成发生错误！" & vbCrLf & ex.Message)
            Return
        End Try
    End Sub
    Public Sub 阶段核算表(ByRef dt As Date, xlsheet As Worksheet, ByRef dat As DataTable, ByRef tad As DataTable)
        Dim dtm As New DataTable, dtt As Date
        Try
            cnct.Open()
            cmd = New SqlCommand("select dbo.盘存日期(NULL,'" & Format(DateAdd(DateInterval.Day, -1, dt), "yyyy-MM-dd") & "',@盘存类型)", cnct)
            cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
            dtt = CDate(cmd.ExecuteScalar())
            xlsheet.Cells(1, 11).Value = Format(dtt, "yy/MM/dd") & "～" & Format(dt, "yy/MM/dd")
            cnct.Close()
            cmd = New SqlCommand("阶段核算", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("日期", dt))
            cmd.Parameters.Add(New SqlParameter("物料名称", dat))
            cmd.Parameters.Add(New SqlParameter("操作工序", tad))
            cmd.Parameters.Add(New SqlParameter("盘存类型", Fcsb.s23()))
            da = New SqlDataAdapter(cmd)
            da.Fill(dtm)
            xlsheet.Cells.ImportDataTable(dtm, True, "A3")
            Dim l As Integer = 3
            Dim k As Integer = dtm.Rows.Count
            Dim a() As Integer
            ReDim a(0)
            Dim st1 As New Style
            Dim st2 As New Style
            Dim xlcell As Cells = xlsheet.Cells
            If k > 1 Then
                Do
                    If CDec(xlcell(l, 10).Value) <> CDec(xlcell(l + 1, 10).Value) Then
                        a(UBound(a)) = l
                        ReDim Preserve a(UBound(a) + 1)
                    End If
                    l = l + 1
                    If l = k + 2 Then
                        a(UBound(a)) = l
                        Exit Do
                    End If
                Loop
                If a(0) > 3 Then
                    xlcell.Merge(3, 6, a(0) - 2, 1)
                    For i = 1 To UBound(a)
                        If a(i) - a(i - 1) > 1 Then
                            xlcell.Merge(a(i - 1) + 1, 6, a(i) - a(i - 1), 1)
                        End If
                    Next
                End If
            End If
            xlcell.DeleteColumn(10)
            xlcell.Merge(0, 0, 1, 12)
            xlcell.Merge(1, 0, 1, 9)
            For i = 3 To xlcell.MaxDataRow
                For j = 2 To 4 Step 2
                    xlcell(i, j).Formula = xlcell(i, j).Value.ToString
                Next
            Next
            st2.Font.Name = "仿宋"
            st2.Font.Size = 11
            st2.HorizontalAlignment = TextAlignmentType.Left
            st2.VerticalAlignment = TextAlignmentType.Center
            xlcell.CreateRange(1, 0, 1, 9).ApplyStyle(st2, New StyleFlag With {.All = True})
            st1.Font.Name = "宋体"
            st1.Font.Size = 14
            st1.Font.IsBold = True
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.VerticalAlignment = TextAlignmentType.Center
            st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Double
            xlcell.CreateRange("A1:L1").ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "Times New Roman"
            st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st1.Font.Size = 12
            st1.Custom = "0.000"
            st1.Font.IsBold = False
            st1.HorizontalAlignment = TextAlignmentType.Right
            st1.VerticalAlignment = TextAlignmentType.Center
            xlcell.CreateRange("C4:J" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.Custom = "0"
            xlcell.CreateRange("A4:A" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "楷体"
            xlcell.CreateRange("B4:B" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "微软雅黑"
            st1.Font.Size = 12
            xlcell.CreateRange("A3:L3").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("K4:L" & k + 3).Merge()
            st1.Font.Name = "仿宋"
            xlcell.CreateRange("K4:L" & k + 3).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("K2:L2").Merge()
            xlcell("J2").Value = "日期"
            xlcell.CreateRange("K3:L3").Merge()
            xlcell("K3").Value = "备注"
            xlcell.CreateRange("K3:K3").Merge()
            st1.Font.Name = "仿宋"
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.None
            st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.None
            xlcell.CreateRange("J2:L2").ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.HorizontalAlignment = TextAlignmentType.Left
            st1.VerticalAlignment = TextAlignmentType.Top
            st1.IsTextWrapped = True
            xlcell("K4").SetStyle(st1)
            xlcell.SetColumnWidth(0, 5)
            xlcell.CreateRange("B:H").ColumnWidth = 12
            xlcell.CreateRange("I:J").ColumnWidth = 15
            xlcell.CreateRange("K:L").ColumnWidth = 10
            With xlsheet.PageSetup
                .LeftMargin = 0.9
                .RightMargin = 0.9
                .TopMargin = 0.5
                .BottomMargin = 0.5
                .HeaderMargin = 0
                .FooterMargin = 0
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = PageOrientationType.Landscape
                .PaperSize = PaperSizeType.PaperA4
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
            xlcell.CreateRange("4:" & k + 3).RowHeight = Math.Max(Math.Min(625 / k, 30), 20)
            xlcell.CreateRange("1:1").RowHeight = 21.5
            xlcell.CreateRange("2:2").RowHeight = 14.25
            xlcell.CreateRange("3:3").RowHeight = 17.25
            s18(xlsheet, 3, 10, "阶段核算表", dt)
            xlsheet.Workbook.CalculateFormula(True)
        Catch ex As Exception
            cnct.Close()
            MsgBox(Format(dt, "yyyy-MM-dd") & "阶段核算表生成发生错误！" & vbCrLf & ex.Message)
        End Try
    End Sub
    Public Sub 消耗产量表(ByRef dt As Date, xlsheet As Worksheet, ByRef dat As DataTable, ByRef tad As DataTable)
        Dim dtm As New DataTable
        Try
            cmd = New SqlCommand("工序报表", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("日期", dt))
            cmd.Parameters.Add(New SqlParameter("物料名称", dat))
            cmd.Parameters.Add(New SqlParameter("操作工序", tad))
            cmd.Parameters.Add(New SqlParameter("盘存类型", Fcsb.s23()))
            da = New SqlDataAdapter(cmd)
            da.Fill(dtm)
            xlsheet.Cells.ImportDataTable(dtm, True, "A4")
            Dim str As String
            Dim xlcell As Cells
            Dim st1 As New Style
            Dim st2 As New Style
            xlcell = xlsheet.Cells
            Dim i, l As Integer
            Dim k As New List(Of Integer)
            xlcell.CreateRange("1:1").RowHeight = 23.33
            xlcell.CreateRange("A1:I1").Merge()
            xlcell.CreateRange("A2:G2").Merge()
            xlcell.CreateRange("A3:A4").Merge()
            xlcell.CreateRange("B3:B4").Merge()
            xlcell.CreateRange("C3:E3").Merge()
            xlcell.CreateRange("F3:H3").Merge()
            xlcell.CreateRange("I3:I4").Merge()
            i = 5
            k.Add(5)
            Do
                str = CStr(xlcell(i - 1, 0).Value)
                If str = "" Then Exit Do
                Do
                    i = i + 1
                    If CStr(xlcell(i - 1, 0).Value) <> str Then
                        k.Add(i) : Exit Do
                    End If
                Loop
            Loop
            For l = 0 To k.Count - 2
                If k.Item(l) + 1 <= k.Item(l + 1) - 1 Then
                    xlcell.CreateRange("A" & k.Item(l) & ":A" & k.Item(l + 1) - 1).Merge()
                End If
            Next
            xlcell.CreateRange("2:" & i - 1).RowHeight = Math.Max(Math.Min(889 / (i - 2), 30), 18)
            xlcell.CreateRange("I5:I" & i - 1).Merge()
            xlcell.CreateRange("A1:I1").Merge()
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.VerticalAlignment = TextAlignmentType.Center
            st1.Font.Name = "宋体"
            st1.Font.Size = 12
            st1.Font.IsBold = True
            xlcell.CreateRange("A1:I1").ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            xlcell.CreateRange("A3:I" & i - 1).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell(1, 7).Value = "日期"
            st1.HorizontalAlignment = TextAlignmentType.Right
            xlcell.CreateRange("H2").ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.HorizontalAlignment = TextAlignmentType.Left
            xlcell.CreateRange("A2:G2").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("I2").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("A:A").ColumnWidth = 6.67
            xlcell.CreateRange("B:H").ColumnWidth = 17.86
            xlcell.CreateRange("C:H").ColumnWidth = 11.11
            xlcell.CreateRange("E:E").ColumnWidth = 13.33
            xlcell.CreateRange("I:I").ColumnWidth = 13.33
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.Font.Name = "微软雅黑"
            st1.Font.Size = 12
            st1.Font.IsBold = False
            xlcell.CreateRange("A3:I4").ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell.CreateRange("A5:A" & i - 1).ApplyStyle(st1, New StyleFlag With {.All = True})
            st1.Font.Name = "楷体"
            xlcell.CreateRange("B5:B" & i - 1).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell(2, 8).Value = "备注"
            st2.VerticalAlignment = TextAlignmentType.Top
            st2.HorizontalAlignment = TextAlignmentType.Left
            st2.Font.Name = "仿宋"
            st2.Font.Size = 12
            st2.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st2.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            st2.IsTextWrapped = True
            xlcell.CreateRange("I5:I" & i - 1).ApplyStyle(st2, New StyleFlag With {.All = True})
            st1.Custom = "0.000"
            st1.HorizontalAlignment = TextAlignmentType.Right
            st1.Font.Name = "Times New Roman"
            st1.Font.Size = 11
            xlcell.CreateRange("C5:H" & i - 1).ApplyStyle(st1, New StyleFlag With {.All = True})
            xlcell(2, 0).Value = "工序"
            xlcell(2, 1).Value = "物料名称"
            xlcell(2, 2).Value = "消耗"
            xlcell(2, 5).Value = "产量"
            st1.Font.Name = "仿宋"
            st1.HorizontalAlignment = TextAlignmentType.Center
            st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Double
            st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.None
            st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.None
            xlcell.CreateRange("A2:I2").ApplyStyle(st1, New StyleFlag With {.All = True})
            With xlsheet.PageSetup
                .LeftMargin = 0.8
                .RightMargin = 0.8
                .TopMargin = 0.8
                .BottomMargin = 0.8
                .HeaderMargin = 0
                .FooterMargin = 0
                .CenterHorizontally = True
                .CenterVertically = False
                .Orientation = PageOrientationType.Portrait
                .PaperSize = PaperSizeType.PaperA4
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
            xlsheet.Cells(1, 8).Value = Format(dt, "yyyy/MM/dd")
            s18(xlsheet, 4, 8, "消耗产量表", dt)
        Catch ex As Exception
            MsgBox(Format(dt, "yyyy-MM-dd") & "消耗产量表生成发生错误！" & vbCrLf & ex.Message)
            Return
        End Try
    End Sub
    Private Sub R_MouseDown(sender As Object, e As EventArgs)
        rc = DirectCast(sender, RadioButton).Checked
    End Sub
    Private Sub R_MouseUp(sender As Object, e As EventArgs)
        DirectCast(sender, RadioButton).Checked = Not rc
    End Sub
    Public Sub CL_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CL5.ItemCheck, CL6.ItemCheck
        Try
            cnct.Open()
            cmd = New SqlCommand("update " & CStr(DirectCast(sender, Control).Tag) & " set 可用性=" & e.NewValue & " where " & CStr(DirectCast(sender, Control).Tag) & "='" & CStr(DirectCast(sender, CheckedListBox).SelectedItem) & "'", cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
            whbl = True
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub TC1_MouseWheel(sender As Object, e As MouseEventArgs) Handles TC1.MouseWheel
        If TypeOf ActiveControl Is TextBox Then
            TBEC = DirectCast(ActiveControl, TextBox)
            If TBEC Is T38 AndAlso T38.Text = "" AndAlso suer <> 4 Then D1.Checked = False
            Fcsb.s37(TBEC, Math.Sign(e.Delta))
            If DirectCast(sender, TabControl).SelectedIndex < 5 Then TBEC.Tag = TBEC.Text
        End If
    End Sub
    Private Sub D_KeyUp(sender As Object, e As KeyEventArgs) Handles D1.KeyUp, D2.KeyUp, D3.KeyUp, D4.KeyUp, D9.KeyUp, D10.KeyUp
        Dim D As DateTimePicker = DirectCast(sender, DateTimePicker)
        If e.KeyCode = Keys.Escape Then D.Text = Format(D.Value, "yyyy-MM-dd 00:00")
    End Sub
    Public Sub MouseWheel(sender As Object, e As MouseEventArgs)
        Dim color As Color, DA As DataGridView = DirectCast(sender, DataGridView)
        If Not DA.IsCurrentCellInEditMode AndAlso dabl Then
            Dim item As Integer = Math.Sign(e.Delta)
            Dim order, ecl As Integer, bl As Boolean
            ecl = CInt(DA Is DA3 OrElse DA Is DA5 OrElse DA Is DA6 OrElse DA Is DA10 OrElse DA Is DA11)
            If er > -1 Then
                If DA.SelectedCells.Count > 0 AndAlso ec > -1 AndAlso DA.Rows(er).Cells(ec).Selected Then
                    For Each cell As DataGridViewCell In DA.SelectedCells
                        If cell.ColumnIndex > ecl Then dacell.Add(cell)
                    Next
                ElseIf ec = -1 Then
                    For Each row As DataGridViewRow In DA.SelectedRows
                        If er = row.Index Then
                            bl = True
                            Exit For
                        End If
                    Next
                    If bl Then
                        For Each row As DataGridViewRow In DA.SelectedRows
                            For i = ecl + 1 To DA.Columns.Count - 1
                                dacell.Add(row.Cells(i))
                            Next
                        Next
                    Else
                        For i = ecl + 1 To DA.Columns.Count - 1
                            dacell.Add(DA.Rows(er).Cells(i))
                        Next
                    End If
                ElseIf ec > ecl Then
                    dacell.Add(DA.Rows(er).Cells(ec))
                End If
            ElseIf ec > -1 Then
                If ec > ecl Then
                    If DA.SelectedCells.Count = 0 Then
                        For i = 0 To DA.Rows.Count - 2
                            dacell.Add(DA.Rows(i).Cells(ec))
                        Next
                    Else
                        For Each cell As DataGridViewCell In DA.SelectedCells
                            dacell.Add(DA.Rows(cell.RowIndex).Cells(ec))
                        Next
                    End If
                End If
            ElseIf er = -1 AndAlso ec = -1 Then
                For Each row As DataGridViewRow In DA.Rows
                    For Each cell As DataGridViewCell In row.Cells
                        If cell.ColumnIndex > ecl Then dacell.Add(cell)
                    Next
                Next
            End If
            If dacell.Count > 0 Then
                For Each dacel As DataGridViewCell In dacell
                    If item > 0 Then
                        If dacl(DA).Contains(dacel) Then
                            If dacel.RowIndex Mod 2 = 0 Then
                                dacel.Style.BackColor = DA.RowsDefaultCellStyle.BackColor
                            Else
                                dacel.Style.BackColor = DA.AlternatingRowsDefaultCellStyle.BackColor
                            End If
                        End If
                        Continue For
                    ElseIf scm.Keys.Contains(dacel.Style.BackColor) Then
                        color = dacel.Style.BackColor
                    Else
                        color = scm.Keys.Last()
                    End If
                    Try
                        cnct.Open()
                        order = (scm(color) + 1) Mod CInt(New SqlCommand("select count(*)-1 from 单耗类别", cnct).ExecuteScalar())
                        cnct.Close()
                        For Each cl As Color In scm.Keys
                            If scm(cl) = order Then
                                dacel.Style.BackColor = cl
                                If Not dacl(DA).Contains(dacel) Then dacl(DA).Add(dacel)
                                Exit For
                            End If
                        Next
                    Catch ex As Exception
                        cnct.Close()
                    End Try
                Next
                dacell.Clear()
            End If
        End If
    End Sub
    Public Sub CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs)
        er = e.RowIndex : ec = e.ColumnIndex
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If Form2 IsNot Nothing AndAlso DA Is Form2.DA1 AndAlso TypeOf ActiveControl IsNot DataGridView Then DA.Focus()
        If er = -1 AndAlso ec = -1 Then DA.ClearSelection()
    End Sub
    Public Sub CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs)
        er = -2 : ec = -2
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        dabl = e.KeyCode = Keys.ShiftKey
    End Sub
    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        dabl = False
    End Sub
    Private Sub LI7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LI7.SelectedIndexChanged
        Dim LI As ListBox = DirectCast(sender, ListBox)
        RemoveHandler T51.TextChanged, AddressOf T51_TextChanged
        If LI7.SelectedItems.Count = 1 Then
            Try
                cnct.Open()
                T51.Text = If(New SqlCommand("select 位号 from 储槽特性 where 储槽名称='" & Replace(LI.SelectedItem.ToString, "'", "''") & "'", cnct).ExecuteScalar, "").ToString
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        Else
            T51.Text = "储槽名称："
        End If
        AddHandler T51.TextChanged, AddressOf T51_TextChanged
    End Sub
    Private Sub DA10_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA10.CellBeginEdit
        If e.ColumnIndex = 3 Then DirectCast(sender, DataGridView).Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub
    Private Sub D8_MouseUp(sender As Object, e As MouseEventArgs) Handles D8.MouseUp
        If R17.Checked Then DirectCast(sender, DateTimePicker).Checked = True
    End Sub

    Private Sub G12_Enter(sender As Object, e As EventArgs) Handles G12.Enter

    End Sub

    Private Sub G15_Enter(sender As Object, e As EventArgs) Handles G15.Enter

    End Sub

    Private Sub G10_Enter(sender As Object, e As EventArgs) Handles G10.Enter

    End Sub

    Private Sub G13_Enter(sender As Object, e As EventArgs) Handles G13.Enter

    End Sub

    Private Sub G16_Enter(sender As Object, e As EventArgs) Handles G16.Enter

    End Sub

    Private Sub DA10_KeyDown(sender As Object, e As KeyEventArgs) Handles DA10.KeyDown
        ctlbl = e.Control
    End Sub
    Private Sub DA10_KeyUp(sender As Object, e As KeyEventArgs) Handles DA10.KeyUp
        ctlbl = False
    End Sub
    Sub s1(li1 As ListBox, li2 As ListBox)
        If IsNothing(li1.SelectedItem) Then Return
        For Each r In li1.SelectedItems
            If Not li2.Items.Contains(r) Then li2.Items.Add(r)
        Next
        For i = 0 To li1.SelectedItems.Count - 1
            li1.Items.Remove(li1.SelectedItems.Item(0))
        Next
    End Sub
    Sub s2(li1 As ListBox, li2 As ListBox)
        If li1.Items.Count = 0 Then Return
        For i = 0 To li1.Items.Count - 1
            If Not li2.Items.Contains(li1.Items.Item(0)) Then li2.Items.Add(li1.Items.Item(0))
            li1.Items.RemoveAt(0)
        Next i
    End Sub
    Sub s3(ByRef str As String, ByRef str1 As String, ByRef str2 As String, ByRef TSMIM As ToolStripMenuItem, ByRef current As DataGridViewCell)
        cmd = New SqlCommand("select dbo." & str2 & "(@物料名称,@日期,@盘存类型,@日库存标记,@月消耗标记)", cnct)
        cmd.Parameters.AddWithValue("物料名称", str1)
        cmd.Parameters.AddWithValue("日期", str)
        cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
        cmd.Parameters.AddWithValue("日库存标记", TSMIM.OwnerItem.Name)
        cmd.Parameters.AddWithValue("月消耗标记", TSMIM.Name)
        If IsDBNull(cmd.ExecuteScalar()) Then
            DA10.Rows(current.RowIndex).Cells(3).Value = Nothing
        Else
            DA10.Rows(current.RowIndex).Cells(3).Value = cmd.ExecuteScalar()
        End If
        current.Tag = {CInt(TSMIM.OwnerItem.Name), CInt(TSMIM.Name)}
    End Sub
    Sub s4(DA As DataGridView, ByRef n As Integer, lsum As TextBox, lcnt As TextBox, lavg As TextBox, lmax As TextBox, lmin As TextBox)
        Dim i As Single, j As New List(Of Decimal)
        For Each cell As DataGridViewCell In DA.SelectedCells
            If cell.ColumnIndex = n AndAlso cell.Value IsNot Nothing Then j.Add(CDec(cell.Value))
        Next
        lmin.Text = CStr(IIf(Single.TryParse(lmin.Text, i), j.Min(), String.Concat(lmin.Text, j.Min())))
        lmax.Text = CStr(IIf(Single.TryParse(lmax.Text, i), j.Max(), String.Concat(lmax.Text, j.Max())))
        lcnt.Text = CStr(IIf(Single.TryParse(lcnt.Text, i), j.Count, String.Concat(lcnt.Text, j.Count)))
        lsum.Text = CStr(IIf(Single.TryParse(lsum.Text, i), j.Sum(), String.Concat(lsum.Text, j.Sum())))
        lavg.Text = CStr(IIf(Single.TryParse(lavg.Text, i), Format(j.Average(), "0.000"), String.Concat(lavg.Text, Format(j.Average(), "0.000"))))
    End Sub
    Sub s5(ByRef ws As Worksheet, ByRef DA As DataGridView)
        Dim j, k, rm, mr, a(DA.Columns.Count - 1) As Integer, xlcell As Cells = ws.Cells
        ws.Name = "DATA"
        mr = xlcell.MaxDataRow + 1
        Dim st As New Style
        st.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
        st.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
        st.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
        st.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
        st.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
        st.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
        For i = 0 To UBound(a)
            a(DA.Columns(i).DisplayIndex) = i
        Next
        For i = 0 To DA.Columns.Count - 1
            If DA.Columns(a(i)).Visible Then
                xlcell(mr, k).Value = DA.Columns(a(i)).HeaderText
                k += 1
            End If
        Next
        xlcell.CreateRange(mr, 0, 1, DA.DisplayedColumnCount(False)).ApplyStyle(st, New StyleFlag With {.All = True})
        k = 0
        For i = 0 To DA.Columns.Count - 1
            If DA.Columns(a(i)).Visible Then
                rm = mr
                For j = 0 To DA.Rows.Count - 2
                    xlcell(mr + 1, k).Value = DA.Rows(j).Cells(a(i)).Value
                    If DA.Rows(j).Cells(a(i)).Style.BackColor.Name <> "0" AndAlso DA.Rows(j).Cells(a(i)).Style.BackColor <> DA.RowsDefaultCellStyle.BackColor AndAlso DA.Rows(j).Cells(a(i)).Style.BackColor <> DA.AlternatingRowsDefaultCellStyle.BackColor Then
                        st.ForegroundColor = DA.Rows(j).Cells(a(i)).Style.BackColor
                        st.Pattern = BackgroundType.Solid
                    Else
                        st.Pattern = BackgroundType.None
                    End If
                    xlcell.CreateRange(mr + 1, k, 1, 1).ApplyStyle(st, New StyleFlag With {.All = True})
                    mr += 1
                Next
                mr = rm
                k += 1
            End If
        Next
        ws.AutoFitRows() : ws.AutoFitColumns()
    End Sub
    Sub s6(CB As ComboBox)
        CB.Items.Clear()
        Try
            dr = New SqlCommand(cmdstr, cnct).ExecuteReader
            While dr.Read()
                CB.Items.Add(dr(0))
            End While
            dr.Close()
        Catch ex As Exception
            dr.Close()
            MsgBox("程序未能正确加载" & vbCrLf & ex.Message)
            Return
        End Try
    End Sub
    Sub s7(tb1 As DataTable, tb2 As DataTable, ByRef clbl As Boolean)
        tb2.Reset()
        tb2.Columns.Add("物料名称")
        For Each r In LI5.Items
            tb2.Rows.Add(r)
        Next
        For Each r In LI6.Items
            tb2.Rows.Add(r)
        Next
        DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Clear()
        If LI3.Items.Count + LI4.Items.Count > 0 Then
            cmd = New SqlCommand("消耗产量", cnct)
            cmd.Parameters.Add(New SqlParameter("操作工序", tb2))
            cmd.Parameters.Add(New SqlParameter("物料类型", tb1))
            cmd.Parameters.Add(New SqlParameter("类型", CByte(IIf(clbl, 2, 1))))
            cmd.CommandType = CommandType.StoredProcedure
            Fcsb.s53(dtn, cmd)
            dr = cmd.ExecuteReader
            While dr.Read
                LI1.Items.Add(dr(0))
                DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Add(dr(0))
            End While
            DirectCast(DA1.Columns.Item(3), DataGridViewComboBoxColumn).Items.Add("")
            dr.Close()
        End If
    End Sub
    Sub s8(ByRef xx As Integer)
        Dim bl As Boolean
        Dim dgv3 As DataGridViewComboBoxColumn = DirectCast(DA1.Columns(3), DataGridViewComboBoxColumn)
        Dim dgv6 As DataGridViewComboBoxColumn = DirectCast(DA1.Columns(6), DataGridViewComboBoxColumn)
        Dim dgv8 As DataGridViewComboBoxColumn = DirectCast(DA1.Columns(8), DataGridViewComboBoxColumn)
        Try
            cnct.Open()
            cmdstr = "select * from 物料数量 where Id=" & xx
            cmd = New SqlCommand(cmdstr, cnct)
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    If Not dgv8.Items.Contains(IIf(IsDBNull(dr(8)), "", dr(8))) Then
                        MsgBox("所选物料: " & CStr(dr(3)) & " 所在工序 " & CStr(dr(8)) & " 不在工序列表中！")
                    ElseIf Not dgv3.Items.Contains(dr(3)) Then
                        MsgBox("所选物料: " & CStr(dr(3)) & " 不在物料列表中！")
                    ElseIf Not dgv6.Items.Contains(dr(6)) Then
                        MsgBox("所选物料: " & CStr(dr(3)) & " 所在类型 " & CStr(dr(6)) & " 不在类型列表中！")
                    Else
                        DA1.Rows.Add()
                        For i = 0 To 9
                            DA1.Rows(DA1.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                        Next
                        DA1.Rows(DA1.Rows.Count - 2).Cells(1).Value = Format(CDate(dr(1)), "yyyy-MM-dd HH:mm")
                    End If
                End While
                cnct.Close()
            Else
                dr.Close()
                cmdstr = "select ident_current('物料数量')"
                cmd = New SqlCommand(cmdstr, cnct)
                dr = cmd.ExecuteReader
                While dr.Read
                    If xx > 0 Then
                        If CInt(dr(0)) < xx Then
                            MsgBox("所查的记录不存在")
                        Else
                            bl = True
                        End If
                    Else
                        MsgBox("记录号必须大于0")
                    End If
                End While
                cnct.Close()
                If bl Then Fcsb.s11(xx, DA1)
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox("物料查询过程中有错误！" & vbCrLf & ex.Message)
            Return
        End Try
        Fcsb.s14(B16, DA1.Rows.Count - 2, DA1, idt1, Color.DarkViolet)
    End Sub
    Sub s9(ByRef xx As Integer)
        Dim bl As Boolean
        Try
            cnct.Open()
            dr = New SqlCommand("select * from 储槽液位 where Id=" & xx, cnct).ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    If DirectCast(DA2.Columns(2), DataGridViewComboBoxColumn).Items.Contains(dr(2)) Then
                        DA2.Rows.Add()
                        For i = 0 To dr.FieldCount - 1
                            DA2.Rows(DA2.Rows.Count - 2).Cells(i).Value = dr(i)
                        Next
                        DA2.Rows(DA2.Rows.Count - 2).Cells(1).Value = Format(CDate(dr(1)), "yyyy-MM-dd HH:mm")
                        bl = True
                    Else
                        MsgBox("所选储槽: " & CStr(dr(2)) & " 不可用！")
                    End If
                End While
                cnct.Close()
                If bl Then
                    Fcsb.s17(DA2.Rows.Count - 2)
                    Fcsb.s18(DA2.Rows.Count - 2)
                End If
            Else
                dr.Close()
                dr = New SqlCommand("select ident_current('储槽液位')", cnct).ExecuteReader
                While dr.Read
                    If xx > 0 Then
                        If CInt(dr(0)) < xx Then
                            MsgBox("所查的记录不存在")
                        Else
                            bl = True
                        End If
                    Else
                        MsgBox("记录号必须大于0")
                    End If
                End While
                cnct.Close()
                If bl Then Fcsb.s47(xx, DA2)
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox("储槽查询中有错误！" & vbCrLf & ex.Message)
            Return
        End Try
        Fcsb.s14(B26, DA2.Rows.Count - 2, DA2, idt2, Color.Pink)
    End Sub
    Private Sub L101_TextChanged(sender As Object, e As EventArgs) Handles L101.TextChanged, T3.TextChanged
        L102.Text = ""
        If L101.Text <> "" AndAlso T3.Text <> "" Then L102.Text = T3.Text & "/" & L101.Text
    End Sub
    Sub s10(ByRef str As String, ByRef str1 As String, ByRef str2 As String, ByRef TSMIM As ToolStripMenuItem, ByRef current As DataGridViewCell)
        cmd = New SqlCommand("select dbo." & str2 & "(@物料名称,@日期,@盘存类型,@消耗标记,@模式)", cnct)
        cmd.Parameters.AddWithValue("物料名称", str1)
        cmd.Parameters.AddWithValue("日期", str)
        cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
        cmd.Parameters.AddWithValue("消耗标记", TSMIM.OwnerItem.Name)
        cmd.Parameters.AddWithValue("模式", TSMIM.Name)
        If IsDBNull(cmd.ExecuteScalar()) Then
            DA10.Rows(current.RowIndex).Cells(3).Value = Nothing
        Else
            DA10.Rows(current.RowIndex).Cells(3).Value = cmd.ExecuteScalar()
        End If
        current.Tag = {CInt(TSMIM.OwnerItem.Name), CInt(TSMIM.Name)}
    End Sub
    Sub s11(e As DataGridViewCellEventArgs)
        RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
        RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
        DA1.Columns(e.ColumnIndex).Visible = True
        DA1.CurrentCell = DA1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        dgvcell = DA1.Rows(e.RowIndex).Cells(e.ColumnIndex) : skip(1) = True
        DA1.BeginEdit(False)
        AddHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
        AddHandler DA1.RowValidating, AddressOf DA1_RowValidating
    End Sub
    Sub s12(T() As TextBox, ByRef str As String)
        Dim i As Integer = -CInt(T(1).Text = "" AndAlso T(0).Text <> "" OrElse T(0).Focused)
        T(i).Text = T(1 - i).Text
        If IsNumeric(T(i).Text) Then
            T(i).Text = Format(CDec(T(i).Text), str)
            T(1 - i).Text = T(i).Text
        End If
        T(i).Focus()
        T(i).SelectionStart = T(i).TextLength
    End Sub
    Sub s13(e As DataGridViewCellEventArgs)
        RemoveHandler DA2.RowValidating, AddressOf DA2_RowValidating
        RemoveHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
        DA2.Columns(e.ColumnIndex).Visible = True
        DA2.CurrentCell = DA2.Rows(e.RowIndex).Cells(e.ColumnIndex)
        dgvcell = DA2.Rows(e.RowIndex).Cells(e.ColumnIndex) : skip(1) = True
        DA2.BeginEdit(False)
        AddHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
        AddHandler DA2.RowValidating, AddressOf DA2_RowValidating
    End Sub
    Sub s14()
        RemoveHandler DA3.SelectionChanged, AddressOf DA3_SelectionChanged
        Dim dac As DataGridViewTextBoxColumn
        DA3.Columns.Clear()
        DA3.Columns.Add("CC1", "时间")
        DA3.Columns(0).Width = 125
        DA3.Columns(0).Frozen = True
        Try
            cmd = New SqlCommand(cmdstr, cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                dac = New DataGridViewTextBoxColumn
                dac.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                dac.HeaderText = CStr(dr(0))
                dac.Name = CStr(dr(0))
                dac.Resizable = DataGridViewTriState.True
                DA3.Columns.Add(dac)
            End While
            dr.Close()
        Catch ex As Exception
            dr.Close()
            MsgBox("储槽载入失败" & vbCrLf & ex.Message)
        End Try
        AddHandler DA3.SelectionChanged, AddressOf DA3_SelectionChanged
    End Sub
    Sub s15(tb1 As DataTable, tb2 As DataTable)
        Dim i, v As Integer, dac As DataGridViewTextBoxColumn, dt As New DataTable, gd, ge, cn As String
        For i = 1 To DA5.Columns.Count - 5
            DA5.Columns.RemoveAt(5)
        Next
        Try
            cmd = New SqlCommand("消耗产量", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("操作工序", tb2))
            cmd.Parameters.Add(New SqlParameter("物料类型", tb1))
            cmd.Parameters.Add(New SqlParameter("类型", 0))
            dr = cmd.ExecuteReader
            While dr.Read()
                v += 1
                gd = CStr(dr(0))
                If ge <> gd Then
                    DA5.Columns.Add(CStr(dr(0)), "")
                    If v = 1 Then
                        DA5.Columns(CStr(dr(0))).Width = 50
                        DA5.Columns(CStr(dr(0))).HeaderText = "|" & gd
                    Else
                        DA5.Columns(CStr(dr(0))).Width = 76
                        DA5.Columns(CStr(dr(0))).HeaderText = ge & "|" & gd
                    End If
                    DA5.Columns(CStr(dr(0))).ReadOnly = True
                    DA5.Columns(CStr(dr(0))).Resizable = DataGridViewTriState.False
                    DA5.Columns(CStr(dr(0))).DefaultCellStyle.BackColor = Color.FromArgb(13, 184, 246)
                    DA5.Columns(CStr(dr(0))).SortMode = DataGridViewColumnSortMode.NotSortable
                    DA5.Columns(CStr(dr(0))).Resizable = DataGridViewTriState.True
                    ge = gd
                End If
                cn = CStr(dr(0)) & CStr(v) & CStr(dr(2))
                dac = New DataGridViewTextBoxColumn
                dac.SortMode = DataGridViewColumnSortMode.NotSortable
                dac.Name = cn
                dac.HeaderText = CStr(dr(1)) & CStr(dr(2))
                dac.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                DA5.Columns.Add(dac)
            End While
            DA5.Columns.Add("Last", ge & "|")
            DA5.Columns("Last").Width = 50
            DA5.Columns("Last").ReadOnly = True
            DA5.Columns("Last").Resizable = DataGridViewTriState.False
            DA5.Columns("Last").Resizable = DataGridViewTriState.True
            DA5.Columns("Last").SortMode = DataGridViewColumnSortMode.NotSortable
            DA5.Columns("Last").DefaultCellStyle.BackColor = Color.FromArgb(13, 184, 246)
            dr.Close()
        Catch ex As Exception
            dr.Close()
        End Try
    End Sub
    Sub s16(sender As TextBox)
        Dim TND As Decimal
        Dim bl As Boolean
        RemoveHandler DirectCast(TN(sender)(0), TextBox).TextChanged, AddressOf T_TextChanged
        Try
            If sender.Text.Replace(" ", "") = Fcsb.s49(sender.Text, bl, TND) AndAlso bl Then
                cnct.Open()
                cmdstr = "select dbo." & CStr(TN(sender)(2)) & "计算(@储槽名称," & CStr(TN(sender)(1)) & ",@时间)"
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.Parameters.Add(New SqlParameter("储槽名称", CO8.Text))
                cmd.Parameters.Add(New SqlParameter(CStr(TN(sender)(1)), TND))
                cmd.Parameters.Add(New SqlParameter("时间", D11.Text))
                dr = cmd.ExecuteReader
                While dr.Read
                    DirectCast(TN(sender)(0), TextBox).Text = Format(dr(0), "0.000")
                End While
                cnct.Close()
            Else
                DirectCast(TN(sender)(0), TextBox).Text = ""
            End If
        Catch ex As Exception
            cnct.Close()
            DirectCast(TN(sender)(0), TextBox).Text = ""
        End Try
        DirectCast(TN(sender)(0), TextBox).Tag = DirectCast(TN(sender)(0), TextBox).Text
        AddHandler DirectCast(TN(sender)(0), TextBox).TextChanged, AddressOf T_TextChanged
    End Sub
    Sub s17(current As DataGridViewCell, TSMIM As ToolStripMenuItem, ByRef str2 As String, ByRef str1 As String, ByRef str As String, Optional ByRef strm As String = "", Optional ByRef strn As String = "")
        current.Tag = CInt(TSMIM.Name)
        cmd = New SqlCommand("select dbo." & str2 & "(@物料名称,@日期" & strm & strn & ")", cnct)
        cmd.Parameters.AddWithValue("物料名称", str1)
        cmd.Parameters.AddWithValue("日期", str)
        If strm <> "" Then cmd.Parameters.AddWithValue(Strings.Right(strm, Len(strm) - 2), Fcsb.s23())
        If strn <> "" Then cmd.Parameters.AddWithValue(Strings.Right(strn, Len(strn) - 2), current.Tag)
        If IsDBNull(cmd.ExecuteScalar()) Then
            DA10.Rows(current.RowIndex).Cells(3).Value = Nothing
        Else
            DA10.Rows(current.RowIndex).Cells(3).Value = cmd.ExecuteScalar()
        End If
    End Sub
    Sub s18(xlsheet As Worksheet, ByRef b As Integer, ByRef c As Integer, ByRef d As String, ByRef rq As Date)
        Dim dt As String
        Try
            cnct.Open()
            xlsheet.Cells(b, c).Value = New SqlCommand("select 报表备注 from 报表备注 where 备注日期='" & rq & "' and 报表名称='" & d & "'", cnct).ExecuteScalar
            dt = New SqlCommand("select max(备注日期) from 长期备注 where 备注日期<='" & rq & "' and 报表名称='" & Strings.Right(xlsheet.Workbook.FileName, 5) & "'", cnct).ExecuteScalar.ToString
            xlsheet.Cells(1, 0).Value = New SqlCommand("select 长期备注 from 长期备注 where 备注日期='" & dt & "' and 报表名称='" & Strings.Right(xlsheet.Workbook.FileName, 5) & "'", cnct).ExecuteScalar
            xlsheet.Cells(0, 0).Value = New SqlCommand("select 报表标签 from 报表配置 where 报表名称='" & d & "'", cnct).ExecuteScalar
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Sub s19(ByRef b As Object, c As DataTable, d As DataTable, f As DataTable)
        For i = 1 To d.Rows.Count
            Try
                cnct.Open()
                cmd.CommandTimeout = 0
                cmd = New SqlCommand("时点统计", cnct)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("日期", f))
                cmd.Parameters.Add(New SqlParameter("物料名称", c))
                cmd.Parameters.Add(New SqlParameter("盘存类型", b))
                cmd.Parameters.Add(New SqlParameter("统计类型", d.Rows(i - 1)(0)))
                dr = cmd.ExecuteReader
                RemoveHandler DA10.CellValueChanged, AddressOf DA10_CellValueChanged
                While dr.Read
                    DA10.Rows.Add()
                    DA10.Rows(DA10.Rows.Count - 2).Cells(0).Value = Format(CDate(dr(0)), "yyyy-MM-dd")
                    DA10.Rows(DA10.Rows.Count - 2).Cells(1).Value = dr(1)
                    DA10.Rows(DA10.Rows.Count - 2).Cells(2).Value = dr(2)
                    If IsDBNull(dr(3)) Then
                        DA10.Rows(DA10.Rows.Count - 2).Cells(3).Value = Nothing
                    Else
                        DA10.Rows(DA10.Rows.Count - 2).Cells(3).Value = dr(3)
                    End If
                    If dr.FieldCount = 5 Then
                        For Each key As Color In scm.Keys
                            If scm(key) = CByte(dr(4)) Then
                                DA10.Rows(DA10.Rows.Count - 2).Cells(3).Style.BackColor = key
                                If Not dacl(DA10).Contains(DA10.Rows(DA10.Rows.Count - 2).Cells(3)) Then dacl(DA10).Add(DA10.Rows(DA10.Rows.Count - 2).Cells(3))
                                Exit For
                            End If
                        Next
                    End If
                End While
                AddHandler DA10.CellValueChanged, AddressOf DA10_CellValueChanged
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
                MsgBox("时点统计未能完成统计" & vbCrLf & ex.Message)
                Return
            End Try
        Next
    End Sub
    Sub s20(li1 As ListBox, li2 As ListBox, nm As String, ky As Boolean)
        If IsNothing(li1.SelectedItem) Then Return
        For i = 0 To li1.SelectedItems.Count - 1
            li2.Items.Add(li1.SelectedItems.Item(i))
            Try
                cnct.Open()
                cmdstr = "update " & nm & "特性 set 可用性='" & ky & "' where " & nm & "名称='" & Replace(CStr(li1.SelectedItems.Item(i)), "'", "''") & "'"
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
                MsgBox("改变可用性失败!表:" & nm & "特性;" & nm & ":" & CStr(li1.SelectedItems.Item(i)) & ";程序已退出!")
                Return
            End Try
        Next
        For i = 0 To li1.SelectedItems.Count - 1
            li1.Items.Remove(li1.SelectedItems.Item(0))
        Next
    End Sub
    Sub s21(e As DataGridViewCellEventArgs)
        RemoveHandler DA9.RowValidating, AddressOf DA9_RowValidating
        RemoveHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
        DA9.Columns(e.ColumnIndex).Visible = True
        DA9.CurrentCell = DA9.Rows(e.RowIndex).Cells(e.ColumnIndex)
        dgvcell = DA9.Rows(e.RowIndex).Cells(e.ColumnIndex) : skip(1) = True
        DA9.BeginEdit(False)
        AddHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
        AddHandler DA9.RowValidating, AddressOf DA9_RowValidating
    End Sub
    Sub s22(ByRef bl As Boolean, ByRef ex As Integer)
        Dim txt, file, filea As String, dt, dtt As New DataTable, sf As String, SBF As New FolderBrowserDialog
        Try
            If Not PB.Visible Then
                bbbl = False
                B109.Show()
                PB.Value = 0
                PB.Show()
                Application.DoEvents()
                If bl Then
                    If SBF.ShowDialog = Windows.Forms.DialogResult.OK Then
                        txt = SBF.SelectedPath
                    Else
                        PB.Hide()
                        B109.Hide()
                        L124bl = False
                        Return
                    End If
                Else
                    txt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If
                If ex >= 0 AndAlso ex <= 26 Then
                    sf = ".xls"
                ElseIf ex >= 27 AndAlso ex <= 50 Then
                    sf = ".pdf"
                ElseIf ex >= 51 AndAlso ex <= 82 Then
                    sf = ".xlsx"
                End If
                Dim dt1, dt2 As Date
                If D5.Text = D6.Text OrElse Not D5.Checked Then
                    dt.Columns.Add("日期", Type.GetType("System.DateTime"))
                    dt.Rows.Add(CDate(D6.Text))
                ElseIf L39.Text = "~" Then
                    dt1 = Date.FromOADate(Math.Min(CDate(D5.Text).ToOADate, CDate(D6.Text).ToOADate))
                    dt2 = Date.FromOADate(Math.Max(CDate(D5.Text).ToOADate, CDate(D6.Text).ToOADate))
                    cmd = New SqlCommand("Select * from dbo.日期序列('" & dt1 & "','" & dt2 & "',1)", cnct)
                    da = New SqlDataAdapter(cmd)
                    da.Fill(dt)
                Else
                    dt1 = Date.FromOADate(Math.Min(CDate(D5.Text).ToOADate, CDate(D6.Text).ToOADate))
                    dt2 = Date.FromOADate(Math.Max(CDate(D5.Text).ToOADate, CDate(D6.Text).ToOADate))
                    dt.Columns.Add("日期", Type.GetType("System.DateTime"))
                    dt.Rows.Add(dt1)
                    dt.Rows.Add(dt2)
                End If
                cmd = New SqlCommand("select * from dbo.盘存序列(@t,@盘存类型)", cnct)
                cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
                Dim p As SqlParameter = New SqlParameter("t", SqlDbType.Structured)
                p.Value = dt
                p.TypeName = "dbo.统计时段"
                cmd.Parameters.Add(p)
                da = New SqlDataAdapter(cmd)
                da.Fill(dtt)
                PB.Maximum = -dtt.Rows.Count * (CInt(CH22.CheckState = 1) + CInt(IIf(CH22.Text = "阶段核算表" AndAlso CH22.CheckState = 2, -1, 0))) - dt.Rows.Count * (CInt(CH20.Checked) + CInt(CH22.CheckState = 1) + CInt(CH21.Checked) + CInt(IIf(CH22.Text = "平均核算表" AndAlso CH22.CheckState = 2, -1, 0)))
                Dim dat As New DataTable
                dat.Columns.Add("物料名称")
                If CL1.CheckedItems.Count = 0 Then
                    For Each r In CL1.Items
                        dat.Rows.Add(r)
                    Next
                Else
                    For Each r In CL1.CheckedItems
                        dat.Rows.Add(r)
                    Next
                End If
                If dat.Rows(0)(0).ToString = "全部" Then dat.Rows.RemoveAt(0)
                Dim tad As New DataTable
                tad.Columns.Add("物料名称")
                If CL2.CheckedItems.Count = 0 Then
                    For Each r In CL2.Items
                        tad.Rows.Add(r)
                    Next
                Else
                    For Each r In CL2.CheckedItems
                        tad.Rows.Add(r)
                    Next
                End If
                If tad.Rows.Item(0)(0).ToString = "全部" Then tad.Rows.RemoveAt(0)
                If Not (CH20.Checked OrElse CH21.Checked OrElse CH22.Checked) Then
                    MsgBox("至少选择一种报表！")
                Else
                    For Each CH In {CH20, CH21, CH22}
                        For Each s In DirectCast(CH.Tag, String())
                            If s <> "" Then
                                filea = ""
                                s34(DirectCast(IIf(s = "阶段核算表", dtt, dt), DataTable), txt, filea, s, sf, PB, dat, tad)
                                file += filea
                            End If
                        Next
                    Next
                    If file <> "" Then
                        If bbbl Then
                            MsgBox("报表生成已中断，部分报表可能已保存，文件名为:" & vbCrLf & file)
                        Else
                            MsgBox("若没有错误提示，报表可能已经保存，文件名为:" & vbCrLf & file)
                        End If
                    End If
                End If
                PB.Hide()
                B109.Hide()
            End If
            L124bl = False
        Catch xe As Exception
            cnct.Close()
            PB.Hide()
            B109.Hide()
            L124bl = False
            MsgBox(xe.Message)
        End Try
    End Sub
    Sub s23(current As DataGridViewCell, TSMIM As ToolStripMenuItem, ByRef str2 As String, ByRef strn As String)
        Dim dt As DataTable
        current.Tag = CBool(-CInt(TSMIM.Name))
        cmd = New SqlCommand("select 物料数量,单耗标记 from dbo." & str2 & "(@物料名称,@日期,@盘存类型," & strn & ")", cnct)
        cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
        cmd.Parameters.AddWithValue(strn, current.Tag)
        dt = New DataTable
        dt.Columns.Add("物料名称")
        dt.Rows.Add(DA10.Rows(current.RowIndex).Cells(1).Value.ToString)
        Dim sp As SqlParameter = New SqlParameter("物料名称", SqlDbType.Structured) With {.Value = dt, .TypeName = "dbo.库存物料"}
        cmd.Parameters.Add(sp)
        dt = New DataTable
        dt.Columns.Add("时间", Type.GetType("System.DateTime"))
        dt.Rows.Add(CDate(DA10.Rows(current.RowIndex).Cells(0).Value))
        sp = New SqlParameter("日期", SqlDbType.Structured) With {.Value = dt, .TypeName = "dbo.统计时段"}
        cmd.Parameters.Add(sp)
        dr = cmd.ExecuteReader()
        DA10.Rows(current.RowIndex).Cells(3).Value = Nothing
        While dr.Read()
            If scm.Keys.Contains(DA10.Rows(current.RowIndex).Cells(3).Style.BackColor) AndAlso scm(DA10.Rows(current.RowIndex).Cells(3).Style.BackColor) = CByte(dr(1)) Then DA10.Rows(current.RowIndex).Cells(3).Value = dr(0) : Exit While
        End While
    End Sub
    Sub s24(ByRef blct As Boolean)
        Dim xlbook As Workbook, xlsheet As Worksheet, xlcell As Cells, st1 As New Style, txt, txtpd, file0, file As String, i As Integer
        st1.Font.Name = "Times New Roman"
        st1.Font.Size = 12
        st1.HorizontalAlignment = TextAlignmentType.Center
        st1.VerticalAlignment = TextAlignmentType.Center
        st1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
        st1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
        st1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
        st1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
        st1.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
        st1.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
        lbl(L125)(0) = False
        For Each value As String In bn.Values.Distinct
            Try
                If blct Then
                    SFD.FileName = value
                    Do
                        If SFD.ShowDialog = Windows.Forms.DialogResult.OK Then
                            txt = Strings.Left(SFD.FileName, SFD.FileName.LastIndexOf("."))
                            txtpd = Strings.Right(LCase(SFD.FileName), SFD.FileName.Length - SFD.FileName.LastIndexOf("."))
                            If txtpd = ".xls" OrElse txtpd = ".xlsx" Then
                                Exit Do
                            Else
                                SFD.FileName = value
                                MsgBox("不支持的文件格式（目前只支持Excel文档）！")
                            End If
                        Else
                            Return
                        End If
                    Loop
                    SFD.FileName = ""
                Else
                    txt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & value
                    txtpd = ".xls"
                End If
                Dim xg As String = Strings.Left(gd(value), Len(gd(value)) - 1)
                xg = Strings.Right(xg, Len(xg) - 1)
                Dim str() As String = xg.Split(CChar("|"))
                Dim bl(1) As Boolean, st2 As New Style, ds As New DataSet, dt As New DataTable
                st2.Copy(st1)
                st2.Font.IsBold = True : st2.Font.Color = Color.Red
                dt.Columns.Add("批号", Type.GetType("System.String"))
                For Each key In bn.Keys
                    If gd(value).Contains("|" & Fcsb.s13(key) & "|") Then dt.Rows.Add(key)
                Next
                cmdstr = "declare @i bit declare @游标批号 varchar(30)"
                If UBound(str) = 0 Then
                    cmdstr += "declare @" & Replace(Replace(str(0), "$", "$$"), "'", "$") & "批号 table(批号 varchar(50))"
                    cmdstr += "insert into @" & Replace(Replace(str(0), "$", "$$"), "'", "$") & "批号 select[" & str(0) & "批号]from " & str(0) & "工艺 where[" & str(0) & "批号]COLLATE Chinese_PRC_CI_AS in(select 批号 from @批号)union all select '" & vbCrLf & "'"
                    cmdstr += "set @i=0declare cusr cursor for(select 批号 from @" & Replace(Replace(str(0), "$", "$$"), "'", "$") & "批号)open cusr fetch next from cusr into @游标批号 WHILE @@FETCH_STATUS = 0begin if @i=0begin set @i=1select*into[#" & str(0) & "]from dbo.[" & str(0) & "输出](@游标批号)end else begin insert into[#" & str(0) & "]select*from dbo.[" & str(0) & "输出](@游标批号)end fetch next from cusr into @游标批号 end close cusr deallocate cusr"
                    cmdstr += " select*from[#" & str(0) & "]where 批号<>'" & vbCrLf & "'"
                Else
                    For Each s As String In str
                        cmdstr += "declare @" & Replace(Replace(s, "$", "$$"), "'", "$") & " table(批号 varchar(50))declare @" & Replace(Replace(s, "$", "$$"), "'", "$") & "批号 table(批号 varchar(50))"
                    Next
                    For Each s As String In str
                        cmdstr += "insert into @" & Replace(Replace(s, "$", "$$"), "'", "$") & " select[" & s & "批号]from[" & s & "工艺]where[" & s & "批号]COLLATE Chinese_PRC_CI_AS in(select 批号 from @批号)union all select'" & vbCrLf & "'"
                    Next
                    If CH39.Checked Then
                        For i = UBound(str) - 1 To 0 Step -1
                            cmdstr += "insert into @" & Replace(Replace(str(i), "$", "$$"), "'", "$") & " select 上一工序批号 from[" & str(i + 1) & "工艺]where[" & str(i + 1) & "批号]COLLATE Chinese_PRC_CI_AS In(select 批号 from @" & Replace(Replace(str(i + 1), "$", "$$"), "'", "$") & ")"
                        Next
                        For i = 1 To UBound(str)
                            cmdstr += "insert into @" & Replace(Replace(str(i), "$", "$$"), "'", "$") & " select[" & str(i) & "批号]from[" & str(i) & "工艺]where 上一工序批号 COLLATE Chinese_PRC_CI_AS in(select 批号 from @" & Replace(Replace(str(i - 1), "$", "$$"), "'", "$") & ")"
                        Next
                    End If
                    For Each s As String In str
                        cmdstr += " insert into @" & Replace(Replace(s, "$", "$$"), "'", "$") & "批号 select distinct 批号 from @" & Replace(Replace(s, "$", "$$"), "'", "$") & " where 批号 is not NULL"
                    Next
                    For Each s As String In str
                        cmdstr += " set @i=0declare cusr cursor for(select 批号 from @" & Replace(Replace(s, "$", "$$"), "'", "$") & "批号)open cusr fetch next from cusr into @游标批号 WHILE @@FETCH_STATUS = 0begin if @i=0begin set @i=1select * into[#" & s & "]from dbo.[" & s & "输出](@游标批号)end else begin insert into[#" & s & "]select*from dbo.[" & s & "输出](@游标批号)end fetch next from cusr into @游标批号 end close cusr deallocate cusr"
                    Next
                    For Each s As String In str
                        cmdstr += " select*from[#" & s & "]where[" & s & "批号]<>'" & vbCrLf & "'"
                    Next
                    If CH39.Checked Then
                        cmdstr += "select"
                        For i = 0 To UBound(str) - 1
                            cmdstr += "[#" & str(i) & "].*,"
                        Next
                        cmdstr += "[#" & str(UBound(str)) & "].*into #Result from"
                        For i = 0 To UBound(str)
                            cmdstr += "[#" & str(i) & "],"
                        Next
                        For i = 1 To UBound(str) - 1
                            cmdstr += "[" & str(i) & "工艺],"
                        Next
                        cmdstr += "[" & str(UBound(str)) & "工艺]where"
                        For i = 0 To UBound(str) - 2
                            cmdstr += "[#" & str(i) & "].[" & str(i) & "批号]COLLATE Chinese_PRC_CI_AS=[" & str(i + 1) & "工艺].上一工序批号 and[" & str(i + 1) & "工艺].[" & str(i + 1) & "批号]COLLATE Chinese_PRC_CI_AS=[#" & str(i + 1) & "].[" & str(i + 1) & "批号]and"
                        Next
                        cmdstr += "[#" & str(UBound(str) - 1) & "].[" & str(UBound(str) - 1) & "批号]COLLATE Chinese_PRC_CI_AS=[" & str(UBound(str)) & "工艺].上一工序批号 and [" & str(UBound(str)) & "工艺].[" & str(UBound(str)) & "批号]COLLATE Chinese_PRC_CI_AS=[#" & str(UBound(str)) & "].[" & str(UBound(str)) & "批号]"
                        cmdstr += "select*from #Result where #Result.[" & str(UBound(str)) & "批号]<>'" & vbCrLf & "'"
                    End If
                End If
                cmd = New SqlCommand(cmdstr, cnct) With {.CommandTimeout = 0}
                cmd.Parameters.Add(New SqlParameter("批号", SqlDbType.Structured) With {.Value = dt, .TypeName = "dbo.批号"})
                da = New SqlDataAdapter(cmd)
                da.Fill(ds)
                xlbook = New Workbook
                For i = 0 To ds.Tables.Count - 1
                    If ds.Tables(i).Rows.Count > 0 Then
                        If bl(0) Then
                            xlbook.Worksheets.Add()
                        Else
                            bl(0) = True
                        End If
                        xlsheet = xlbook.Worksheets(xlbook.Worksheets.Count - 1)
                        If CH39.Checked AndAlso i = ds.Tables.Count - 1 AndAlso UBound(str) > 0 Then
                            If Not bl(1) Then
                                xlsheet.Name = value
                            Else
                                xlbook.Worksheets.RemoveAt(xlbook.Worksheets.Count - 1)
                                Continue For
                            End If
                        ElseIf CL4.Items.Contains(Fcsb.s13(ds.Tables(i).Rows(0)(0).ToString)) Then
                            xlsheet.Name = Fcsb.s13(ds.Tables(i).Rows(0)(0).ToString)
                        Else
                            xlbook.Worksheets.RemoveAt(xlbook.Worksheets.Count - 1)
                            bl(1) = True
                            Continue For
                        End If
                        xlcell = xlsheet.Cells
                        xlcell.ImportDataTable(ds.Tables(i), True, "A1")
                        xlcell.CreateRange(0, 0, xlcell.MaxDataRow + 1, xlcell.MaxDataColumn + 1).ApplyStyle(st1, New StyleFlag With {.All = True})
                        For h = 0 To xlcell.MaxDataColumn
                            For k = 0 To xlcell.MaxDataRow
                                xlcell(k, h).Value = IIf(IsNothing(xlcell(k, h).Value), "", xlcell(k, h).Value)
                                If xlcell(k, h).Value.ToString.Contains("不合格") Then
                                    Try
                                        xlcell(k, h).SetStyle(st2)
                                    Catch ex As Exception
                                        xlcell.CreateRange(k, h, 1, 1).ApplyStyle(st2, New StyleFlag With {.All = True})
                                    End Try
                                End If
                            Next
                        Next
                    End If
                Next
                file0 = txt & txtpd
                i = 0
                Do While IO.File.Exists(file0) AndAlso Not blct
                    i += 1
                    file0 = txt & i & txtpd
                Loop
                xlbook.Save(file0)
                file += file0 & vbCrLf
            Catch ex As Exception
                cnct.Close()
                xlbook = Nothing
                MsgBox(value & "导出遇到问题" & vbCrLf & ex.Message)
            End Try
        Next
        If file IsNot Nothing Then MsgBox("工段工艺表已保存，文件名为：" & vbCrLf & file)
    End Sub
    Sub s25(DA As DataGridView, ByRef er As Integer)
        Dim str As Byte
        Dim cl As Color = DA.Rows(er).Cells(0).Style.BackColor
        If DA.Rows(er).Cells(6).Value Is "" Then DA.Rows(er).Cells(6).Value = 0
        Dim data As Decimal = CDec(DA.Rows(er).Cells(6).Value)
        If CDec(DA.Rows(er).Cells(6).Value) = 0 Then DA.Rows(er).Cells(6).Value = Nothing
        Try
            cnct.Open()
            cmdstr = "select Id from 单耗类别 where R=@R and G=@G and B=@B"
            cmd = New SqlCommand(cmdstr, cnct)
            cmd.Parameters.AddWithValue("R", cl.R)
            cmd.Parameters.AddWithValue("G", cl.G)
            cmd.Parameters.AddWithValue("B", cl.B)
            dr = cmd.ExecuteReader
            While dr.Read
                str = CByte(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            Return
        End Try
        cmd = New SqlCommand()
        cmdstr = "update 工序类型 set 单耗预估值=NULL where 物料名称=@物料名称 and 单耗标记=@单耗标记"
        If DA.Rows(er).Cells(6).Value IsNot Nothing Then
            cmdstr += vbCrLf
            cmdstr += "update 工序类型 set 单耗预估值=@单耗预估值 where Id=@Id"
            cmd.Parameters.AddWithValue("单耗预估值", data)
            cmd.Parameters.AddWithValue("Id", DA.Rows(er).Cells(0).Value)
        End If
        cmd.Parameters.AddWithValue("物料名称", CStr(DA.Rows(er).Cells(1).Value))
        cmd.Parameters.AddWithValue("单耗标记", str)
        cmd.CommandText = cmdstr
        cmd.Connection = cnct
        Try
            cnct.Open()
            cmd.ExecuteNonQuery()
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        For i = 1 To DA.Rows.Count
            If DA.Rows(i - 1).Cells(0).Style.BackColor = cl AndAlso i - 1 <> er AndAlso CStr(DA.Rows(er).Cells(1).Value) = CStr(DA.Rows(i - 1).Cells(1).Value) Then
                DA.Rows(i - 1).Cells(6).Value = Nothing
            End If
        Next
    End Sub
    Sub s26(CL As CheckedListBox, ByRef gxlx As String)
        CL.Items.Clear()
        cnct.Open()
        dr = New SqlCommand("select " & gxlx & ",可用性 from " & gxlx & " order by Id", cnct).ExecuteReader
        RemoveHandler CL.ItemCheck, AddressOf CL_ItemCheck
        While dr.Read
            CL.Items.Add(dr(0))
            CL.SetItemChecked(CL.Items.Count - 1, CBool(dr(1)))
        End While
        AddHandler CL.ItemCheck, AddressOf CL_ItemCheck
        cnct.Close()
    End Sub
    Sub s27(ByRef blct As Boolean)
        lbl(L130)(0) = False
        If Not DA10.Rows(0).IsNewRow Then
            If Not CH30.Checked AndAlso Not CH31.Checked Then
                MsgBox("请至少选择一种文件类型！")
            Else
                Dim txt As String, i, j As Integer, pvt As Pivot.PivotTable, fd As New FolderBrowserDialog, xlbook As New Workbook, ws As Worksheet
                If blct Then
                    If fd.ShowDialog = DialogResult.OK Then txt = fd.SelectedPath
                Else
                    txt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If
                If txt <> "" Then
                    Try
                        xlbook.Worksheets.Add()
                        s5(xlbook.Worksheets(1), DA10)
                        i = xlbook.Worksheets(1).Cells.MaxDataRow + 1
                        j = xlbook.Worksheets(1).Cells.MaxDataColumn
                        ws = xlbook.Worksheets(0)
                        ws.Name = "数据透视表"
                        ws.PivotTables.Add("=DATA!A1:" & Chr(65 + j) & i, "A2", "pvt")
                        pvt = ws.PivotTables(0)
                        pvt.IsAutoFormat = True
                        pvt.AutoFormatType = Pivot.PivotTableAutoFormatType.Classic
                        pvt.RowGrand = False
                        pvt.ColumnGrand = False
                        If j > 0 Then pvt.AddFieldToArea(Pivot.PivotFieldType.Row, 0)
                        If j > 1 Then pvt.AddFieldToArea(Pivot.PivotFieldType.Column, 1)
                        If j > 2 Then pvt.AddFieldToArea(Pivot.PivotFieldType.Page, 2)
                        pvt.AddFieldToArea(Pivot.PivotFieldType.Data, j)
                        pvt.DataFields(0).NumberFormat = "0.000"
                        ws.AutoFitRows()
                        ws.AutoFitColumns()
                        Try
                            Dim file As String = txt & "\数据透视表.xls"
                            Dim file0 As String
                            Dim file1 As String
                            If CH31.Checked Then
                                file1 = file & "x"
                                i = 0
                                Do While IO.File.Exists(file1)
                                    i += 1
                                    file1 = txt & "\数据透视表" & i & ".xlsx"
                                Loop
                                xlbook.Save(file1)
                                file1 += vbCrLf
                            End If
                            If CH30.Checked Then
                                file0 = file
                                i = 0
                                Do While IO.File.Exists(file0)
                                    i += 1
                                    file0 = txt & "\数据透视表" & i & ".xls"
                                Loop
                                xlbook.Save(file0)
                                file0 += vbCrLf
                                If DA10.Rows.Count > 65535 Then
                                    file = "请注意由于你勾选了xls格式且数据表超过65535行（不包含标头），超出部分将被截断！"
                                Else
                                    file = ""
                                End If
                            End If
                            MsgBox("数据透视表已保存，文件名为：" & vbCrLf & file0 & file1 & file)
                        Catch ex As Exception
                            MsgBox("文件保存失败！" & vbCrLf & ex.Message)
                        End Try
                    Catch ex As Exception
                        MsgBox("生成数据透视表出错！" & vbCrLf & ex.Message)
                    End Try
                End If
            End If
        Else
            MsgBox("没有数据无法生成数据透视表！")
        End If
    End Sub
    Sub s28(CL As CheckedListBox, ByRef field As String)
        cmdstr = ""
        Dim i As Integer
        For i = 0 To CL.Items.Count - 1
            cmdstr += "update " & field & " set id=" & i + 1 & ",可用性='" & CL.GetItemChecked(i) & "' where " & field & "='" & Replace(CStr(CL.Items(i)), "'", "''") & "'"
        Next
        Try
            cnct.Open()
            i = New SqlCommand(cmdstr, cnct).ExecuteNonQuery()
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub s29(CL As CheckedListBox, ByRef i As Integer)
        If CL.SelectedItems.Count = 0 Then Return
        If CL.SelectedIndex = 0 AndAlso i = -1 OrElse CL.SelectedIndex = CL.Items.Count - 1 AndAlso i = 1 Then Return
        Dim itm As String = CStr(CL.Items(CL.SelectedIndex))
        Dim bl As Boolean = CL.GetItemChecked(CL.SelectedIndex)
        CL.SetItemChecked(CL.SelectedIndex, CL.GetItemChecked(CL.SelectedIndex + i))
        CL.Items(CL.SelectedIndex) = CL.Items(CL.SelectedIndex + i)
        CL.Items(CL.SelectedIndex + i) = itm
        CL.SetItemChecked(CL.SelectedIndex + i, bl)
        CL.SetSelected(CL.SelectedIndex + i, True)
    End Sub
    Sub s30(DA As DataGridView, e As DataGridViewCellMouseEventArgs, ByRef bl As Boolean)
        If DA.Columns.Item(e.ColumnIndex).HeaderText.Contains("|") Then
            Dim cit As Integer = e.ColumnIndex
            Do
                If cit < DA.Columns.Count - 2 Then
                    cit += 1
                    If DA.Columns.Item(cit).HeaderText.Contains("|") Then Exit Do
                    DA.Columns.Item(cit).Visible = bl
                Else
                    Exit Do
                End If
            Loop
        Else
            If bl Then
                For i = 1 To DA.Columns.Count
                    DA.Columns.Item(i - 1).Visible = True
                Next
            Else
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        End If
    End Sub
    Sub s31(CB As ComboBox)
        dr = New SqlCommand(cmdstr, cnct).ExecuteReader
        CB.Items.Clear()
        While dr.Read()
            CB.Items.Add(dr(0))
        End While
        CB.Items.Add("")
        dr.Close()
    End Sub
    Sub s32(ByRef str2 As String, ByRef str1 As String, ByRef str As String, TSMIM As ToolStripMenuItem, current As DataGridViewCell, ByRef strn As String)
        cmd = New SqlCommand("select dbo." & str2 & "(@物料名称,@日期,@盘存类型,@日库存标记,@月消耗标记," & strn & ")", cnct)
        cmd.Parameters.AddWithValue("物料名称", str1)
        cmd.Parameters.AddWithValue("日期", str)
        cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
        If IsNothing(TSMIM.OwnerItem) Then
            cmd.Parameters.AddWithValue("日库存标记", DBNull.Value)
            cmd.Parameters.AddWithValue("月消耗标记", DBNull.Value)
            cmd.Parameters.AddWithValue(strn, TSMIM.Name)
            current.Tag = {CInt(TSMIM.Name), Nothing, Nothing}
        Else
            cmd.Parameters.AddWithValue("日库存标记", TSMIM.OwnerItem.Name)
            cmd.Parameters.AddWithValue("月消耗标记", TSMIM.Name)
            cmd.Parameters.AddWithValue(strn, 0)
            current.Tag = {0, CInt(TSMIM.OwnerItem.Name), CInt(TSMIM.Name)}
        End If
        If IsDBNull(cmd.ExecuteScalar()) Then
            DA10.Rows(current.RowIndex).Cells(3).Value = Nothing
        Else
            DA10.Rows(current.RowIndex).Cells(3).Value = CDec(cmd.ExecuteScalar())
        End If
    End Sub
    Sub s33(ByRef str As String, DA As DataGridView, e As DataGridViewCellEventArgs)
        RemoveHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
        RemoveHandler DA.RowValidating, AddressOf DA12_RowValidating
        skip(0) = True
        MsgBox(str + " 输入有误", MsgBoxStyle.OkOnly)
        DA.Columns(e.ColumnIndex).Visible = True
        DA.CurrentCell = DA.Rows(e.RowIndex).Cells(e.ColumnIndex)
        dgvcell = DA.Rows(e.RowIndex).Cells(e.ColumnIndex)
        skip(1) = True
        DA.BeginEdit(False)
        AddHandler DA.RowValidating, AddressOf DA12_RowValidating
        AddHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
    End Sub
    Sub s34(dt As DataTable, ByRef txt As String, ByRef file As String, ByRef str As String, sf As String, PB As ProgressBar, dat As DataTable, Optional tad As DataTable = Nothing)
        Dim xlbook As New Workbook, xlsheet As Worksheet
        Try
            For Each tm As DataRow In dt.Rows
                If Not bbbl Then
                    If xlbook.Worksheets.Count = 1 Then
                        xlbook.FileName = Format(tm(0), "yyyy-MM-dd") & str
                    Else
                        xlbook.FileName = str
                    End If
                    xlsheet = xlbook.Worksheets(xlbook.Worksheets.Count - 1)
                    xlsheet.Name = Format(tm(0), "yyyy-MM-dd")
                    [GetType].GetMethod(str).Invoke(Me, {tm(0), xlsheet, dat, tad})
                    PB.Value = Math.Min(10000, PB.Value + 1)
                    Application.DoEvents()
                    PB.CreateGraphics().DrawString(Format(PB.Value / PB.Maximum, IIf(PB.Value / PB.Maximum = 1, "0.0% ", "00.00% ").ToString) + Format(tm(0), "yyyy-MM-dd") + str, New System.Drawing.Font("宋体", 10.0!, FontStyle.Regular), Brushes.Red, 98, 5)
                    Application.DoEvents()
                    xlbook.Worksheets.Add()
                Else
                    Exit For
                End If
            Next
            If xlbook.Worksheets.Count > 1 AndAlso Not IsDate(xlbook.Worksheets(xlbook.Worksheets.Count - 1).Name) Then
                Dim i As Integer
                xlbook.Worksheets.RemoveAt(xlbook.Worksheets.Count - 1)
                file = txt & "\" & xlbook.FileName & sf
                Do While IO.File.Exists(file)
                    i += 1
                    file = txt & "\" & xlbook.FileName & i & sf
                Loop
                xlbook.Save(file)
                file += vbCrLf
            End If
        Catch ex As Exception
            MsgBox(xlbook.FileName & "生成失败。详细信息：" & ex.Message)
            PB.Hide()
            B109.Hide()
        End Try
    End Sub
    Sub s35(cell As DataGridViewCell, ByRef bl As Boolean)
        cmd = New SqlCommand("select 物料数量,单耗标记 from dbo." & CStr(DA10.Rows(cell.RowIndex).Cells(2).Value) & "(@物料名称,@日期" + IIf(bl, ",@盘存类型", "").ToString + ")", cnct)
        Dim dt = New DataTable
        dt.Columns.Add("物料名称")
        dt.Rows.Add(DA10.Rows(cell.RowIndex).Cells(1).Value.ToString)
        cmd.Parameters.Add(New SqlParameter("物料名称", SqlDbType.Structured) With {.Value = dt, .TypeName = "dbo.库存物料"})
        dt = New DataTable
        dt.Columns.Add("时间", Type.GetType("System.DateTime"))
        dt.Rows.Add(CDate(DA10.Rows(cell.RowIndex).Cells(0).Value))
        cmd.Parameters.Add(New SqlParameter("日期", SqlDbType.Structured) With {.Value = dt, .TypeName = "dbo.统计时段"})
        If bl Then cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
        dr = cmd.ExecuteReader()
        While dr.Read()
            If scm.Keys.Contains(DA10.Rows(cell.RowIndex).Cells(3).Style.BackColor) AndAlso scm(DA10.Rows(cell.RowIndex).Cells(3).Style.BackColor) = CByte(dr(1)) Then DA10.Rows(cell.RowIndex).Cells(3).Value = CDec(dr(0)) : Exit While
        End While
        dr.Close()
    End Sub
    Sub s36(DA As DataGridView, e As DataGridViewCellCancelEventArgs, ByRef cmdstr As String, ByRef datestr As String, D1 As DateTimePicker, D2 As DateTimePicker, Optional ByRef D3 As DateTimePicker = Nothing, Optional ByRef D4 As DateTimePicker = Nothing)
        sv = Nothing
        If suer <> 4 Then
            D1.Enabled = False
            D2.Enabled = False
        End If
        If D3 IsNot Nothing Then
            D3.Enabled = False
            D4.Enabled = False
        End If
        For Each col As DataGridViewColumn In DA.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        If IsNothing(DA.Rows(e.RowIndex).Cells(0).Value) Then Return
        cmdstr += CStr(DA.Rows(e.RowIndex).Cells(0).Value)
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    sv = IIf(IsDBNull(dr(e.ColumnIndex)), Nothing, dr(e.ColumnIndex))
                    If Not skip(1) Then
                        For Each col As DataGridViewColumn In DA.Columns
                            If TypeOf col Is DataGridViewComboBoxColumn AndAlso Not DirectCast(col, DataGridViewComboBoxColumn).Items.Contains(CStr(IIf(IsDBNull(dr(col.Index)), "", dr(col.Index)))) Then
                                MsgBox("数据库值:" & CStr(IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))) & " 不在" & col.HeaderText & "列表中！")
                            ElseIf col.Index = 1 Then
                                DA.Rows(e.RowIndex).Cells(1).Value = Format(CDate(dr(1)), datestr)
                            Else
                                DA.Rows(e.RowIndex).Cells(col.Index).Value = IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))
                            End If
                        Next
                    End If
                End While
            Else
                e.Cancel = True
                DA.Rows(e.RowIndex).Tag = New Integer() {Nothing, CInt(DA.Rows(e.RowIndex).Cells(0).Value)}
                DA.Rows(e.RowIndex).Cells(0).Value = 0
                DA.Rows(e.RowIndex).ReadOnly = True
            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Sub s37(ByRef bl As Boolean)
        Dim i As Integer, lix As Integer = DA9.Rows.Count - 1
        If bl Then
            cmdstr = "select * from (select top 4 * from 报表备注 order by 备注日期 desc)A order by 备注日期"
        Else
            cmdstr = "select * from 报表备注 where 备注日期"
            If Not D5.Checked Then
                cmdstr += "='" & D6.Text & "'"
            ElseIf L39.Text = "、" Then
                cmdstr += " in('" & D5.Text & "','" & D6.Text & "')"
            Else
                cmdstr += " between '" & D5.Text & "' and '" & D6.Text & "'"
            End If
            cmdstr += " order by 备注日期"
        End If
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader()
            While dr.Read()
                DA9.Rows.Add()
                For i = 1 To dr.FieldCount
                    DA9.Rows(DA9.Rows.Count - 2).Cells(i - 1).Value = dr(i - 1)
                Next
                DA9.Rows(DA9.Rows.Count - 2).Cells(1).Value = Format(DA9.Rows(DA9.Rows.Count - 2).Cells(1).Value, "yyyy-MM-dd")
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox("报表备注载入失败！" & vbCrLf & ex.Message)
        End Try
        Dim CL As Color = Color.FromArgb(255, 204, 204, 204)
        Fcsb.s14(B80, lix, DA9, idt4, CL)
        DA9.ClearSelection()
        ckbl = False
    End Sub
    Sub s38(ByRef bltc As Boolean)
        Dim k As New List(Of String), lix As Integer = DA2.Rows.Count - 1
        If LI8.Items.Count = 0 Then
            If LI7.Items.Count = 0 Then Return
            For Each r In LI7.Items
                k.Add(CStr(r))
            Next
        Else
            For Each r In LI8.Items
                k.Add(CStr(r))
            Next
        End If
        Dim cmdstr0 As String = Fcsb.s2(k, "储槽液位.储槽名称")
        Dim cmdstr1 As String = Fcsb.s5(D3, D4, "日期")
        cmdstr = "select 储槽液位.id,日期,储槽液位.储槽名称,储槽液位,dbo.储槽计算(储槽液位.储槽名称,储槽液位,日期) as 储槽储量,物料名称,操作工序 from 储槽液位,储槽特性 where 储槽液位.储槽名称=储槽特性.储槽名称"
        If cmdstr0 <> "(" Then cmdstr += " and " & cmdstr0
        If Not bltc Then
            If cmdstr1 <> "(" Then cmdstr += " and " & cmdstr1
            cmdstr += " order by 日期"
        End If
        If bltc Then
            cmdstr = "select * from (select top 6 * from (" & cmdstr & ") as A order by 日期 desc) as B order by 日期"
        End If
        Try
            cmd = New SqlCommand(cmdstr, cnct)
            cnct.Open()
            dr = cmd.ExecuteReader
            While dr.Read
                DA2.Rows.Add()
                For i = 0 To 6
                    DA2.Rows(DA2.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                Next
                DA2.Rows(DA2.Rows.Count - 2).Cells(1).Value = Format(CDate(dr(1)), "yyyy-MM-dd HH:mm")
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox("储槽查询发生错误！" & vbCrLf & ex.Message)
        End Try
        Fcsb.s14(B26, lix, DA2, idt2, Color.Pink)
        DA2.ClearSelection()
        ttbl = False
    End Sub
    Sub s39(ByRef blct As Boolean)
        nn = True
        Try
            cnctk.Open()
            If suer = 4 Then D9.Value = Date.FromOADate(Math.Max(DateAdd(DateInterval.Day, -2, CDate(New SqlCommand("select getdate()", cnctk).ExecuteScalar)).ToOADate, D9.Value.ToOADate))
            cnctk.Close()
        Catch ex As Exception
            cnctk.Close()
        End Try
        Dim col As New DataGridViewLinkColumn
        Dim cmdstr1 As String = " order by id desc)A order by rid asc"
        Dim cmdstr2 As String = " 操作员='" & usr & "'"
        Dim cmdstr3 As String = " from (select top 25 * from 操作记录"
        Dim cmdstr4 As String = " from 操作记录 where 操作时间 between '" & Date.FromOADate(Math.Min(D9.Value.ToOADate, D10.Value.ToOADate)) & "' and '" & Date.FromOADate(Math.Max(D9.Value.ToOADate, D10.Value.ToOADate)) & "'"
        Dim cmdstr5 As String = "select 操作员 as U,操作时间,SQL语句,记录表 as 表,记录Id as Id,Id as RId,计算机名"
        DA11.Columns.Clear()
        If blct Then
            If sbl(1) Then
                cmdstr = cmdstr5 & cmdstr3 & " where" & cmdstr2 & cmdstr1
            Else
                cmdstr = cmdstr5 & cmdstr3 & cmdstr1
            End If
        Else
            cmdstr = cmdstr5 & cmdstr4
            If sbl(1) Then cmdstr += " and" & cmdstr2
        End If
        cmd = New SqlCommand(cmdstr, cnctm)
        Try
            cnctm.Open()
            dr = cmd.ExecuteReader()
            For i = 1 To dr.FieldCount
                If i = dr.FieldCount - 2 Then
                    col.SortMode = DataGridViewColumnSortMode.Automatic
                    col.HeaderText = "Id"
                    DA11.Columns.Add(col)
                Else
                    DA11.Columns.Add("", dr.GetName(i - 1))
                End If
            Next
            RemoveHandler DA11.CellValueChanged, AddressOf DA11_CellValueChanged
            While dr.Read()
                DA11.Rows.Add()
                For i = 1 To dr.FieldCount
                    DA11.Rows(DA11.Rows.Count - 2).Cells(i - 1).Value = IIf(IsDBNull(dr(i - 1)), Nothing, dr(i - 1))
                Next
                DA11.Rows(DA11.Rows.Count - 2).Cells(2).Value = Fcsb.s19(CStr(DA11.Rows(DA11.Rows.Count - 2).Cells(2).Value), CStr(DA11.Rows(DA11.Rows.Count - 2).Cells(3).Value))
            End While
            AddHandler DA11.CellValueChanged, AddressOf DA11_CellValueChanged
            cnctm.Close()
            DA11.Columns(0).Width = 45 : DA11.Columns(1).Width = 120
            DA11.Columns(2).Width = 500 : DA11.Columns(3).Width = 55 : DA11.Columns(3).ReadOnly = True
            DA11.Columns(4).Width = 54 : DA11.Columns(5).Visible = False : DA11.Columns(6).Visible = False
            Fcsb.s56(DA11)
            If blct Then DA11.FirstDisplayedScrollingRowIndex = Math.Max(0, DA11.NewRowIndex - 13)
        Catch ex As Exception
            cnctm.Close()
            MsgBox("操作记录查询失败" & vbCrLf & ex.Message)
        End Try
        DA11.ClearSelection()
        ccbl = False
    End Sub
    Sub s40(ByRef blct As Boolean)
        If Not IO.Directory.Exists("D:\" & st(2)) Then IO.Directory.CreateDirectory("D:\" & st(2))
        If blct Then
            IO.File.SetAttributes("D:\" & st(2), IO.FileAttributes.System)
            IO.File.SetAttributes("D:\" & st(2), IO.FileAttributes.Hidden)
            Try
                cnct.Open()
                cmd = New SqlCommand("backup database " & st(2) & " to disk='D:\" & st(2) & "\" & Format(Now, "yyMMddHHmmss") & "'", cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
                MsgBox("数据备份成功！")
            Catch ex As Exception
                cnct.Close()
                MsgBox("数据备份失败！" & vbCrLf & ex.Message)
            End Try
        Else
            Dim cnctm, cnctn As New SqlConnection("data source=" & st(3) & ";initial catalog=master;user id=" & usr & ";password=" & pswd）, i As Integer
            OFD.InitialDirectory = "D:\" & st(2)
            If OFD.ShowDialog = DialogResult.OK Then
                If MsgBox("这将会删除自 " & System.IO.File.GetLastWriteTime(OFD.FileName) & " 以来的所有数据，确定继续吗？", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    Try
                        cnctm.Open()
                        cmd = New SqlCommand("select spid from master..sysprocesses where dbid=db_id('" & st(2) & "')", cnctm)
                        dr = cmd.ExecuteReader()
                        While dr.Read()
                            cnctn.Open()
                            i = New SqlCommand("kill " & CStr(dr("spid")), cnctn).ExecuteNonQuery()
                            cnctn.Close()
                        End While
                        cnctn.Open()
                        i = New SqlCommand("restore database " & st(2) & " from disk='" & OFD.FileName & "'", cnctn).ExecuteNonQuery()
                        MsgBox("恢复备份成功，请自己重新启动该程序即可！")
                        fc = True
                        Application.Exit()
                    Catch ex As Exception
                        fc = True
                        Application.Exit()
                    End Try
                End If
            End If
        End If
        bcbl = False
    End Sub
    Sub s41()
        Try
            cnct.Open()
            dr = New SqlCommand("select * from 单耗类别 where Id is not NULL", cnct).ExecuteReader()
            While dr.Read()
                If scm.ContainsValue(CByte(dr(0))) Then
                    scm(Color.FromArgb(CInt(dr(1)), CInt(dr(2)), CInt(dr(3)))) = CByte(dr(0))
                Else
                    scm.Add(Color.FromArgb(CInt(dr(1)), CInt(dr(2)), CInt(dr(3))), CByte(dr(0)))
                End If
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Sub s42(ByRef bl As Boolean)
        Dim lb, blm As Boolean, dt As Date, xn As Decimal
        bl = True
        Try
            For Each r As DataGridViewRow In DA3.Rows
                If r.IsNewRow AndAlso r.Index = 0 Then MsgBox("请输入数据后重试！") : Return
                If Not r.IsNewRow Then
                    If Date.TryParse(CStr(r.Cells(0).Value), dt) = False Then
                        MsgBox("第" & r.Index + 1 & "行日期出现错误,请重输！")
                        bl = False
                        DA3.CurrentCell = r.Cells(0)
                        DA3.BeginEdit(False)
                        Return
                    End If
                End If
                lb = True
                For Each k As DataGridViewCell In r.Cells
                    blm = True
                    If k.ColumnIndex > 0 Then
                        Try
                            If CStr(k.Value) = "" Then
                                blm = False
                            Else
                                k.Value = Fcsb.s49(CStr(k.Value), True, xn)
                            End If
                            If k.ReadOnly Then lb = False
                            If blm Then
                                lb = False
                                cmdstr = "insert into 储槽液位 values('" & CStr(r.Cells(0).Value) & "','" & Replace(DA3.Columns(k.ColumnIndex).Name, "'", "''") & "'," & CStr(k.Value) & ")"
                                cmdstr += "select max(Id) from 储槽液位"
                                cmd = New SqlCommand(cmdstr, cnct)
                                cnct.Open()
                                idt2.Add(CInt(cmd.ExecuteScalar))
                                cnct.Close()
                            End If
                            k.Value = Nothing
                            k.ReadOnly = True
                        Catch ex As Exception
                            cnct.Close()
                            MsgBox("储槽液位:第" & r.Index + 1 & "行 " & DA3.Columns(k.ColumnIndex).Name & " 录入错误。")
                            DA3.Columns(k.ColumnIndex).Visible = True
                            DA3.CurrentCell = k
                            DA3.BeginEdit(False)
                            bl = False
                            Return
                        End Try
                    End If
                Next
                If lb Then
                    If Not r.IsNewRow Then
                        bl = False
                        MsgBox("第" & r.Index + 1 & "行：请至少输入一个数据")
                        r.ReadOnly = False
                        Return
                    End If
                End If
            Next
            bl = True
            MsgBox("操作已成功！")
        Catch ex As Exception
            MsgBox("储槽液位录入过程中有错误发生，请立即与管理员联系" & vbCrLf & ex.Message)
        End Try
    End Sub
    Sub s43(ByRef bl As Boolean)
        Dim lb As Boolean, dt As Date, msb As MsgBoxResult, blm As Boolean, xn As Decimal, yz(4) As String
        Try
            For Each r As DataGridViewRow In DA5.Rows
                If r.IsNewRow AndAlso r.Index = 0 Then MsgBox("请输入数据后重试！") : Return
                If Not r.IsNewRow Then
                    If Date.TryParse(CStr(r.Cells(0).Value), dt) = False Then
                        MsgBox("消耗产量:第" & r.Index + 1 & "行日期出现错误,请重输！")
                        bl = False
                        DA5.Columns(0).Visible = True
                        DA5.CurrentCell = r.Cells(0)
                        DA5.BeginEdit(False)
                        Return
                    End If
                    Do
                        If CH32.Checked AndAlso Fcsb.s10(CStr(r.Cells(1).Value), True) = 0 AndAlso CStr(r.Cells(1).Value) <> "" Then
                            msb = MsgBox("消耗产量:第" & r.Index + 1 & "行批号格式不正确！", DirectCast(MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.DefaultButton1, MsgBoxStyle))
                            If msb = MsgBoxResult.Abort Then
                                bl = False
                                DA5.Columns(1).Visible = True
                                DA5.CurrentCell = r.Cells(1)
                                DA5.BeginEdit(False)
                                Return
                            ElseIf msb = MsgBoxResult.Ignore Then
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    Dim bln As Boolean = True
                    r.Cells(4).Value = Fcsb.s49(CStr(r.Cells(4).Value), bln, xn)
                    If bln Then
                        If xn <= 0 AndAlso CStr(r.Cells(4).Value) <> "" Then bln = False
                    End If
                    If Not bln Then
                        MsgBox("消耗产量:第" & r.Index + 1 & "行物料含量输入有误，请检查后重输！")
                        bl = False
                        DA5.Columns(4).Visible = True
                        DA5.CurrentCell = r.Cells(4)
                        DA5.BeginEdit(False)
                        Return
                    End If
                End If
                lb = True
                For Each k As DataGridViewCell In r.Cells
                    blm = True
                    If k.ColumnIndex > 5 Then
                        Try
                            If CStr(k.Value) = "" Then
                                blm = False
                            Else
                                k.Value = Fcsb.s49(CStr(k.Value), True, xn)
                            End If
                            If Not DA5.Columns(k.ColumnIndex).HeaderText.Contains("|") Then
                                If k.ReadOnly Then lb = False
                            End If
                            If blm Then
                                lb = False
                                Dim sg As String = Strings.Left(DA5.Columns(k.ColumnIndex).HeaderText, Len(DA5.Columns(k.ColumnIndex).HeaderText) - 2)
                                Try
                                    yz(0) = Replace(sg, "'", "''")
                                    yz(1) = Replace(Strings.Right(DA5.Columns(k.ColumnIndex).Name, 2), "'", "''")
                                    yz(2) = Replace(Strings.Left(DA5.Columns(k.ColumnIndex).Name, 2), "'", "''")
                                    yz(3) = CStr(r.Cells(1).Value)
                                    yz(4) = Replace(CStr(r.Cells(3).Value), "'", "''")
                                    Do
                                        If Fcsb.s15(yz) AndAlso CH33.Checked Then
                                            msb = MsgBox("消耗产量:第" & k.RowIndex + 1 & "行 " & DA5.Columns(k.ColumnIndex).HeaderText & " 输入的条目不匹配！", DirectCast(MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.DefaultButton1, MsgBoxStyle))
                                            If msb = MsgBoxResult.Abort Then
                                                bl = False
                                                DA5.Columns(1).Visible = True
                                                DA5.CurrentCell = r.Cells(1)
                                                DA5.BeginEdit(False)
                                                Return
                                            ElseIf msb = MsgBoxResult.Ignore Then
                                                Exit Do
                                            End If
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                Catch ex As Exception
                                    cnct.Close()
                                    DA5.ClearSelection()
                                    k.Selected = True
                                    bl = False
                                    Return
                                End Try
                                If yz(3) = "" Then yz(3) = "NULL"
                                yz(3).Replace("'", "''")
                                cmdstr = "insert into 物料数量 values('" & CStr(r.Cells(0).Value) & "'," & CStr(IIf(yz(3) = "NULL", yz(3), "'" & yz(3) & "'")) & ",'" & yz(0) & "'," & CStr(k.Value) & "," & CStr(IIf(CStr(r.Cells(4).Value) = "", "NULL", r.Cells(4).Value)) & ",'" & yz(1) & "'," & CStr(IIf(CStr(r.Cells(2).Value) = "", "NULL", "'" & Replace(CStr(r.Cells(2).Value), "'", "''") & "'")) & ",'" & yz(2) & "'," & CStr(IIf(yz(4) = "", "NULL", "'" & yz(4) & "'")) & ")"
                                cmdstr += "select max(Id) from 物料数量"
                                cmd = New SqlCommand(cmdstr, cnct)
                                cnct.Open()
                                idt1.Add(CInt(cmd.ExecuteScalar))
                                cnct.Close()
                            End If
                            k.Value = Nothing
                            k.ReadOnly = True
                        Catch ex As Exception
                            cnct.Close()
                            MsgBox("消耗产量:第" & k.RowIndex + 1 & "行 " & DA5.Columns(k.ColumnIndex).HeaderText & " 录入错误。")
                            DA5.ClearSelection()
                            DA5.Columns(k.ColumnIndex).Visible = True
                            DA5.CurrentCell = k
                            DA5.BeginEdit(False)
                            bl = False
                            Return
                        End Try
                    End If
                Next
                If lb Then
                    If Not r.IsNewRow Then
                        bl = False
                        MsgBox("第" & r.Index + 1 & "行：请至少输入一个数据")
                        r.ReadOnly = False
                        Return
                    End If
                End If
            Next
            bl = True
            MsgBox("操作已成功！")
        Catch ex As Exception
            MsgBox("消耗产量录入过程中有错误发生，请立即与管理员联系" & vbCrLf & ex.Message)
        End Try
    End Sub
    Sub s44(ByRef bl As Boolean)
        Dim lb, blm As Boolean, xn As Decimal, dt As Date, msb As MsgBoxResult
        bl = True
        Try
            If DA6.Rows(0).IsNewRow AndAlso DA6.Rows(0).Index = 0 Then MsgBox("请输入数据后重试！") : bl = False : Return
            If CStr(DA6.Rows(0).Cells(5).Value) = "" Then MsgBox("请选择类型后重试！") : bl = False : Return
            If Not DA6.Rows(0).IsNewRow Then
                If Date.TryParse(CStr(DA6.Rows(0).Cells(0).Value), dt) = False Then
                    MsgBox("消耗产量:第" & DA6.Rows(0).Index + 1 & "行日期格式出现错误,请重输！")
                    bl = False
                    DA6.Columns(0).Visible = True
                    DA6.CurrentCell = DA6.Rows(0).Cells(0)
                    DA6.BeginEdit(False)
                    Return
                End If
            End If
            Do
                If CH32.Checked AndAlso Fcsb.s10(CStr(DA6.Rows(0).Cells(1).Value), True) = 0 AndAlso CStr(DA6.Rows(0).Cells(1).Value) <> "" Then
                    msb = MsgBox("消耗产量:批号格式不正确！", DirectCast(MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.DefaultButton1, MsgBoxStyle))
                    If msb = MsgBoxResult.Abort Then
                        bl = False
                        DA6.Columns(1).Visible = True
                        DA6.CurrentCell = DA6.Rows(0).Cells(1)
                        DA6.BeginEdit(False)
                        Return
                    ElseIf msb = MsgBoxResult.Ignore Then
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            lb = True
            Dim bln As Boolean = True
            DA6.Rows(0).Cells(4).Value = Fcsb.s49(CStr(DA6.Rows(0).Cells(4).Value), bln, xn)
            If bln Then
                If xn <= 0 AndAlso CStr(DA6.Rows(0).Cells(4).Value) <> "" Then bln = False
            End If
            If Not bln Then
                MsgBox("消耗产量:物料含量输入有误，请检查后重输！")
                bl = False
                DA6.Columns(4).Visible = True
                DA6.CurrentCell = DA6.Rows(0).Cells(4)
                DA6.BeginEdit(False)
                Return
            End If
            For Each k As DataGridViewCell In DA6.Rows(0).Cells
                blm = True
                If k.ColumnIndex > 5 Then
                    Try
                        If CStr(k.Value) = "" Then
                            blm = False
                        Else
                            k.Value = Fcsb.s49(CStr(k.Value), True, xn)
                        End If
                        If k.ReadOnly Then lb = False
                        If blm Then
                            Dim yz(4) As String
                            Try
                                yz(0) = Replace(DA6.Columns(k.ColumnIndex).HeaderText, "'", "''")
                                yz(1) = Replace(CStr(DA6.Rows(0).Cells(5).Value), "'", "''")
                                yz(2) = ""
                                yz(3) = CStr(DA6.Rows(0).Cells(1).Value)
                                yz(4) = Replace(CStr(DA6.Rows(0).Cells(3).Value), "'", "''")
                                Do
                                    If Fcsb.s15(yz) AndAlso CH33.Checked Then
                                        msb = MsgBox("消耗产量:" & DA6.Columns(k.ColumnIndex).HeaderText & " 输入的条目不匹配！", DirectCast(MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.DefaultButton1, MsgBoxStyle))
                                        If msb = MsgBoxResult.Abort Then
                                            bl = False
                                            DA6.Columns(1).Visible = True
                                            DA6.CurrentCell = DA6.Rows(0).Cells(1)
                                            DA6.BeginEdit(False)
                                            Return
                                        ElseIf msb = MsgBoxResult.Ignore Then
                                            Exit Do
                                        End If
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            Catch ex As Exception
                                cnct.Close()
                                DA6.ClearSelection()
                                k.Selected = True
                                bl = False
                                Return
                            End Try
                            If yz(3) = "" Then yz(3) = "NULL"
                            lb = False
                            cmdstr = "insert into 物料数量 values('" & CStr(DA6.Rows(0).Cells(0).Value) & "'," & CStr(IIf(yz(3) = "NULL", yz(3), "'" & yz(3) & "'")) & ",'" & yz(0) & "'," & CStr(k.Value) & "," & CStr(IIf(CStr(DA6.Rows(0).Cells(4).Value) = "", "NULL", DA6.Rows(0).Cells(4).Value)) & ",'" & yz(1) & "'," & CStr(IIf(CStr(DA6.Rows(0).Cells(2).Value) = "", "NULL", "'" & Replace(CStr(DA6.Rows(0).Cells(2).Value), "'", "''") & "'")) & ",NULL," & CStr(IIf(CStr(DA6.Rows(0).Cells(3).Value) = "", "NULL", "'" & yz(4) & "'")) & ")"
                            cmdstr += "select max(Id) from 物料数量"
                            cmd = New SqlCommand(cmdstr, cnct)
                            cnct.Open()
                            idt1.Add(CInt(cmd.ExecuteScalar))
                            cnct.Close()
                        End If
                        k.Value = Nothing
                        k.ReadOnly = True
                    Catch ex As Exception
                        cnct.Close()
                        MsgBox("消耗产量:" & DA6.Columns(k.ColumnIndex).Name & " 录入错误。")
                        DA6.Columns(k.ColumnIndex).Visible = True
                        DA6.CurrentCell = k
                        DA6.BeginEdit(False)
                        bl = False
                        Return
                    End Try
                End If
            Next
            If lb Then
                bl = False
                MsgBox("消耗产量:请至少输入一个数据")
                DA6.Rows(0).ReadOnly = False
                Return
            End If
            bl = True
            MsgBox("操作已成功！")
        Catch ex As Exception
            MsgBox("录入过程中有错误发生，请立即与管理员联系" & vbCrLf & ex.Message)
        End Try
    End Sub
    Sub s45(DA As DataGridView, ByRef sv As Object, Optional ByRef time As String = Nothing, Optional D() As DateTimePicker = Nothing)
        Dim bl As Boolean
        If skip(1) Then
            dgvcell.Value = sv
            If time IsNot Nothing AndAlso IsDate(sv) Then dgvcell.Value = Format(sv, time)
            DA.CancelEdit() : DA.EndEdit()
            skip(1) = False
            bl = DA.Rows(dgvcell.RowIndex).Cells(0).Value IsNot Nothing
        Else
            DA.CancelEdit() : DA.EndEdit()
            bl = DA.Rows.Count = 1 OrElse IsNothing(DA.Rows(DA.Rows.Count - 1).Cells(0).Value) AndAlso DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing
        End If
        If bl OrElse DA.CurrentRow.Cells(0).Value IsNot Nothing AndAlso DA.CurrentRow.IsNewRow Then
            If D IsNot Nothing Then
                For Each t As DateTimePicker In D
                    If suer <> 4 Then t.Enabled = True
                Next
            End If
            For Each col As DataGridViewColumn In DA.Columns
                col.SortMode = DataGridViewColumnSortMode.Automatic
            Next
            skip(0) = False
        End If
        s58(DA)
    End Sub
    Sub s46(ByRef bl As Boolean)
        T1.Text = "" : CO9.Text = "" : T38.Text = ""
        T39.Text = "" : T40.Text = "" : T52.Text = "" : T53.Text = ""
        D1.Text = Format(DateAdd(DateInterval.Day, -2, Now), "yyyy-MM-dd 00:00")
        D2.Text = Format(DateAdd(DateInterval.Day, 1, Now), "yyyy-MM-dd 00:00")
        D1.Checked = suer <> 4
        D2.Checked = suer <> 4
        If bl Then
            CL2.SetItemChecked(0, True)
            Dim e As EventArgs
            CL2_MouseUp(CL2, e)
            Dim em As ComponentModel.CancelEventArgs
            CMS4_Opening(CMS4, em)
        Else
            s2(LI2, LI1)
            s2(LI4, LI3)
            s2(LI6, LI5)
        End If
        b105bl = False
    End Sub
    Function s47(ByRef dgvc As List(Of Integer)) As Boolean
        If dgvc.Count > 1 Then
            dgvc.Sort()
            Dim str As New List(Of String)
            For Each i As Integer In dgvc
                If DA10.Rows(dgvc.Item(0)).Cells(2).Tag IsNot Nothing AndAlso DA10.Rows(i).Cells(2).Tag IsNot Nothing Then
                    If IsNumeric(DA10.Rows(i).Cells(2).Tag) Then
                        If CInt(DA10.Rows(dgvc.Item(0)).Cells(2).Tag) <> CInt(DA10.Rows(i).Cells(2).Tag) Then Return True
                    Else
                        For j = 0 To DirectCast(DA10.Rows(dgvc.Item(0)).Cells(2).Tag, Integer()).Length - 1
                            If DirectCast(DA10.Rows(dgvc.Item(0)).Cells(2).Tag, Integer())(j) <> DirectCast(DA10.Rows(i).Cells(2).Tag, Integer())(j) Then Return True
                        Next
                    End If
                ElseIf IsNothing(DA10.Rows(dgvc.Item(0)).Cells(2).Tag) Xor IsNothing(DA10.Rows(i).Cells(2).Tag) Then
                    Return True
                End If
                str.Add(DA10.Rows(DA10.Rows(i).Cells(2).RowIndex).Cells(1).Value.ToString)
            Next
            Return str.Distinct.ToList.Count > 1
        End If
    End Function
    Sub s48(ByRef str As String, ByRef str1 As String, ByRef str2 As String, ByRef TSMIM As ToolStripMenuItem, ByRef current As DataGridViewCell)
        cmd = New SqlCommand("select dbo." & str2 & "(@物料名称,@日期,@盘存类型,@消耗标记,@模式,DEFAULT)", cnct)
        cmd.Parameters.AddWithValue("物料名称", str1)
        cmd.Parameters.AddWithValue("日期", str)
        cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
        cmd.Parameters.AddWithValue("消耗标记", TSMIM.OwnerItem.Name)
        cmd.Parameters.AddWithValue("模式", TSMIM.Name)
        If IsDBNull(cmd.ExecuteScalar()) Then
            DA10.Rows(current.RowIndex).Cells(3).Value = Nothing
        Else
            DA10.Rows(current.RowIndex).Cells(3).Value = cmd.ExecuteScalar()
        End If
        current.Tag = {CInt(TSMIM.OwnerItem.Name), CInt(TSMIM.Name)}
    End Sub
    Sub s49(ByRef CMS As ContextMenuStrip, ByRef ary As List(Of Integer), ByRef num As Integer)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "盘存差", .Name = "0"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "累加型", .Name = "1"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "回收型", .Name = "2"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "单耗型", .Name = "3"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "计划型", .Name = "4"})
        If Not s47(ary) Then
            If IsNothing(DA10.SelectedCells(0).Tag) Then
                DirectCast(CMS.Items(0), ToolStripMenuItem).Checked = True
            Else
                Do
                    DirectCast(CMS.Items(num), ToolStripMenuItem).Checked = num = CInt(DA10.SelectedCells(0).Tag)
                    num += 1
                Loop While num <= 4
            End If
        End If
    End Sub
    Sub s50(ByRef CMS As ContextMenuStrip, ByRef ary As List(Of Integer), ByRef str() As String)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = str(0), .Name = "0"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = str(1), .Name = "1"})
        If Not s47(ary) Then
            If IsNothing(DA10.SelectedCells(0).Tag) Then
                DirectCast(CMS.Items(1), ToolStripMenuItem).Checked = True
            Else
                DirectCast(CMS.Items(0), ToolStripMenuItem).Checked = Not CBool(DA10.SelectedCells(0).Tag)
                DirectCast(CMS.Items(1), ToolStripMenuItem).Checked = CBool(DA10.SelectedCells(0).Tag)
            End If
        End If
    End Sub
    Sub s51(ByRef CMS As ContextMenuStrip, ByRef ary As List(Of Integer), ByRef num As Integer)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "累加型", .Name = "1"})
        TSMIA = DirectCast(CMS.Items(0), ToolStripMenuItem)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "回收型", .Name = "2"})
        TSMIB = DirectCast(CMS.Items(1), ToolStripMenuItem)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "单耗型", .Name = "3"})
        TSMIC = DirectCast(CMS.Items(2), ToolStripMenuItem)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "计划型", .Name = "4"})
        TSMID = DirectCast(CMS.Items(3), ToolStripMenuItem)
        num = 0
        Do
            DirectCast(CMS.Items(num), ToolStripMenuItem).DropDownItems.Add(New ToolStripMenuItem() With {.Text = "绝对值", .Name = "0"})
            DirectCast(CMS.Items(num), ToolStripMenuItem).DropDownItems.Add(New ToolStripMenuItem() With {.Text = "相对值", .Name = "1"})
            num += 1
        Loop While num <= 3
        If Not s47(ary) Then
            If IsNothing(DA10.SelectedCells(0).Tag) Then
                DirectCast(CMS.Items(0), ToolStripMenuItem).Checked = True
                DirectCast(DirectCast(CMS.Items(0), ToolStripMenuItem).DropDownItems(1), ToolStripMenuItem).Checked = True
            Else
                num = 0
                Do
                    DirectCast(CMS.Items(num), ToolStripMenuItem).Checked = num = DirectCast(DA10.SelectedCells(0).Tag, Integer())(0) - 1
                    num += 1
                Loop While num <= 3
                num = 0
                Do
                    DirectCast(DirectCast(CMS.Items(num \ 2), ToolStripMenuItem).DropDownItems(num Mod 2), ToolStripMenuItem).Checked = num Mod 2 = DirectCast(DA10.SelectedCells(0).Tag, Integer())(1) AndAlso DirectCast(CMS.Items(num \ 2), ToolStripMenuItem).Checked
                    num += 1
                Loop While num <= 7
            End If
        End If
    End Sub
    Sub s52(ByRef CMS As ContextMenuStrip, ByRef ary As List(Of Integer), ByRef num As Integer, ByRef dt As DataTable, ByRef i() As Integer)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "盘存型", .Name = "0"})
        TSMIA = DirectCast(CMS.Items(0), ToolStripMenuItem)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "储槽型", .Name = "1"})
        TSMIB = DirectCast(CMS.Items(1), ToolStripMenuItem)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = "智能型", .Name = "2"})
        TSMIC = DirectCast(CMS.Items(2), ToolStripMenuItem)
        For Each ddi As ToolStripMenuItem In CMS.Items
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "累加型", .Name = "1"})
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "回收型", .Name = "2"})
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "单耗型", .Name = "3"})
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "计划型", .Name = "4"})
        Next
        If Not s47(ary) AndAlso DA10.SelectedCells(0).Tag IsNot Nothing Then
            num = 0
            Do
                DirectCast(CMS.Items(num), ToolStripMenuItem).Checked = num = DirectCast(DA10.SelectedCells(0).Tag, Integer())(0)
                num += 1
            Loop While num <= 2
            num = 0
            Do
                DirectCast(DirectCast(CMS.Items(num \ 4), ToolStripMenuItem).DropDownItems(num Mod 4), ToolStripMenuItem).Checked = num Mod 4 = DirectCast(DA10.SelectedCells(0).Tag, Integer())(1) - 1 AndAlso DirectCast(CMS.Items(num \ 4), ToolStripMenuItem).Checked
                num += 1
            Loop While num <= 11
        ElseIf Not s47(ary) Then
            cmd = New SqlCommand(String.Concat("select 日库存标记,月消耗标记 from 物料特性 where 物料名称='", CStr(DA10.Rows(DA10.SelectedCells(0).RowIndex).Cells(1).Value), "'"), cnct)
            da = New SqlDataAdapter(cmd)
            dt = New DataTable()
            da.Fill(dt)
            Try
                i(0) = CInt(dt.Rows(0)(0))
                DirectCast(CMS.Items(CInt(dt.Rows(0)(0))), ToolStripMenuItem).Checked = True
            Catch ex As Exception
            End Try
            Try
                num = 0
                i(1) = CInt(dt.Rows(0)(1))
                Do
                    DirectCast(DirectCast(CMS.Items(num), ToolStripMenuItem).DropDownItems(i(1) - 1), ToolStripMenuItem).Checked = DirectCast(CMS.Items(num), ToolStripMenuItem).Checked
                    num += 1
                Loop While num <= 3
            Catch ex As Exception
            End Try
            If i(0) <> Nothing AndAlso i(1) <> Nothing Then
                DA10.SelectedCells(0).Tag = i
            ElseIf i(0) = Nothing AndAlso i(1) <> Nothing Then
                num = 0
                Do
                    DirectCast(DirectCast(CMS.Items(num), ToolStripMenuItem).DropDownItems(i(1) - 1), ToolStripMenuItem).Checked = True
                    num += 1
                Loop While num <= 3
            End If
        End If
    End Sub
    Sub s53(ByRef CMS As ContextMenuStrip, ByRef ary As List(Of Integer), ByRef str() As String)
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = str(0), .Name = "0"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = str(1), .Name = "1"})
        CMS.Items.Add(New ToolStripMenuItem() With {.Text = str(2), .Name = "2"})
        If Not s47(ary) Then
            If IsNothing(DA10.SelectedCells(0).Tag) Then
                DirectCast(CMS.Items(0), ToolStripMenuItem).Checked = True
            Else
                For j = 0 To 2
                    DirectCast(CMS.Items(j), ToolStripMenuItem).Checked = j = CInt(DA10.SelectedCells(0).Tag)
                Next
            End If
        End If
    End Sub
    Sub s54(ByRef cms As ContextMenuStrip, ByRef ary As List(Of Integer), ByRef num As Integer, ByRef dt As DataTable, ByRef i() As Integer, ByRef tp As String)
        cms.Items.Add(New ToolStripMenuItem() With {.Text = "库存差", .Name = "0"})
        cms.Items.Add(New ToolStripMenuItem() With {.Text = "累加型", .Name = "1"})
        cms.Items.Add(New ToolStripMenuItem() With {.Text = "回收型", .Name = "2"})
        cms.Items.Add(New ToolStripMenuItem() With {.Text = "单耗型", .Name = "3"})
        cms.Items.Add(New ToolStripMenuItem() With {.Text = "计划型", .Name = "4"})
        DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems.Add(New ToolStripMenuItem() With {.Text = "盘存型", .Name = "0"})
        TSMIA = DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(0), ToolStripMenuItem)
        DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems.Add(New ToolStripMenuItem() With {.Text = "储槽型", .Name = "1"})
        TSMIB = DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(1), ToolStripMenuItem)
        DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems.Add(New ToolStripMenuItem() With {.Text = "智能型", .Name = "2"})
        TSMIC = DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(2), ToolStripMenuItem)
        For Each ddi As ToolStripMenuItem In DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "累加型", .Name = "1"})
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "回收型", .Name = "2"})
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "单耗型", .Name = "3"})
            ddi.DropDownItems.Add(New ToolStripMenuItem() With {.Text = "计划型", .Name = "4"})
        Next
        If Not s47(ary) AndAlso DA10.SelectedCells(0).Tag IsNot Nothing Then
            num = 0
            Do
                DirectCast(cms.Items(num), ToolStripMenuItem).Checked = num = DirectCast(DA10.SelectedCells(0).Tag, Integer())(0)
                num += 1
            Loop While num <= 4
            num = 0
            Do
                DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num), ToolStripMenuItem).Checked = num = DirectCast(DA10.SelectedCells(0).Tag, Integer())(1) AndAlso DirectCast(cms.Items(0), ToolStripMenuItem).Checked
                num += 1
            Loop While num <= 2
            num = 0
            Do
                DirectCast(DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num \ 4), ToolStripMenuItem).DropDownItems(num Mod 4), ToolStripMenuItem).Checked = num Mod 4 = DirectCast(DA10.SelectedCells(0).Tag, Integer())(2) - 1 AndAlso DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num \ 4), ToolStripMenuItem).Checked
                num += 1
            Loop While num <= 11
        ElseIf Not s47(ary) Then
            cmd = New SqlCommand(String.Concat("select 日库存标记,月消耗标记," & tp & " from 物料特性 where 物料名称='", CStr(DA10.Rows(DA10.SelectedCells(0).RowIndex).Cells(1).Value), "'"), cnct)
            da = New SqlDataAdapter(cmd)
            dt = New DataTable()
            da.Fill(dt)
            Try
                i(0) = CInt(dt.Rows(0)(2))
                DirectCast(cms.Items(i(0)), ToolStripMenuItem).Checked = True
            Catch ex As Exception
            End Try
            Try
                i(1) = CInt(dt.Rows(0)(0))
                DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(i(1)), ToolStripMenuItem).Checked = DirectCast(cms.Items(0), ToolStripMenuItem).Checked
            Catch ex As Exception
            End Try
            Try
                num = 0
                i(2) = CInt(dt.Rows(0)(1))
                Do
                    DirectCast(DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num), ToolStripMenuItem).DropDownItems(i(2) - 1), ToolStripMenuItem).Checked = DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num), ToolStripMenuItem).Checked
                    num += 1
                Loop While num <= 2
            Catch ex As Exception
            End Try
            If i(0) <> Nothing AndAlso i(1) <> Nothing AndAlso i(2) <> Nothing Then
                DA10.SelectedCells(0).Tag = i
            ElseIf i(0) = 0 AndAlso i(1) = Nothing AndAlso i(2) <> Nothing Then
                num = 0
                Do
                    DirectCast(DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num), ToolStripMenuItem).DropDownItems(i(2) - 1), ToolStripMenuItem).Checked = True
                    num += 1
                Loop While num <= 2
            ElseIf i(0) = Nothing AndAlso i(1) <> Nothing Then
                DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(i(1)), ToolStripMenuItem).Checked = True
                If i(2) <> Nothing Then
                    num = 0
                    Do
                        DirectCast(DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num), ToolStripMenuItem).DropDownItems(i(2) - 1), ToolStripMenuItem).Checked = DirectCast(DirectCast(cms.Items(0), ToolStripMenuItem).DropDownItems(num), ToolStripMenuItem).Checked
                        num += 1
                    Loop While num <= 2
                End If
            End If
        End If
    End Sub
    Sub s55(ByRef cell As DataGridViewCell)
        cmd = New SqlCommand("select dbo." & CStr(DA10.Rows(cell.RowIndex).Cells(2).Value) & "(@物料名称,@日期)", cnct)
        cmd.Parameters.AddWithValue("物料名称", DA10.Rows(cell.RowIndex).Cells(1).Value)
        cmd.Parameters.AddWithValue("日期", DA10.Rows(cell.RowIndex).Cells(0).Value)
        If IsDBNull(cmd.ExecuteScalar()) Then
            DA10.Rows(cell.RowIndex).Cells(3).Value = Nothing
        Else
            DA10.Rows(cell.RowIndex).Cells(3).Value = cmd.ExecuteScalar()
        End If
    End Sub
    Sub s56(ByRef e As System.ComponentModel.CancelEventArgs, Optional ByRef bl As Object = Nothing)
        e.Cancel = True
        cnct.Open()
        For Each cell As DataGridViewCell In DA10.SelectedCells
            If CL1.Items.Contains(DA10.Rows(cell.RowIndex).Cells(1).Value.ToString) AndAlso CStr(DA10.Rows(cell.RowIndex).Cells(1).Value) <> "全部" Then
                If IsNothing(bl) Then
                    s55(cell)
                Else
                    s35(cell, CBool(bl))
                End If
            End If
        Next
        cnct.Close()
    End Sub
    Sub s57(ByRef e As System.ComponentModel.CancelEventArgs, ByRef bl As Boolean)
        e.Cancel = True
        cnct.Open()
        For Each cell As DataGridViewCell In DA10.SelectedCells
            If CL1.Items.Contains(DA10.Rows(cell.RowIndex).Cells(1).Value.ToString) AndAlso CStr(DA10.Rows(cell.RowIndex).Cells(1).Value) <> "全部" Then
                If bl Then
                    cmd = New SqlCommand("select dbo." & CStr(DA10.Rows(cell.RowIndex).Cells(2).Value) & "(@物料名称,@日期,@盘存类型)", cnct)
                Else
                    cmd = New SqlCommand("select convert(numeric(12,3),convert(char(10),dbo." & CStr(DA10.Rows(cell.RowIndex).Cells(2).Value) & "(@物料名称,@日期,@盘存类型),112))", cnct)
                End If
                cmd.Parameters.AddWithValue("物料名称", DA10.Rows(cell.RowIndex).Cells(1).Value)
                cmd.Parameters.AddWithValue("日期", DA10.Rows(cell.RowIndex).Cells(0).Value)
                cmd.Parameters.AddWithValue("盘存类型", Fcsb.s23())
                If IsDBNull(cmd.ExecuteScalar()) Then
                    DA10.Rows(cell.RowIndex).Cells(3).Value = Nothing
                Else
                    DA10.Rows(cell.RowIndex).Cells(3).Value = cmd.ExecuteScalar()
                End If
            End If
        Next
        cnct.Close()
    End Sub
    Private Sub LI7_GotFocus(sender As Object, e As EventArgs) Handles LI7.GotFocus
        If skip(0) Then
            RemoveHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
            RemoveHandler DA2.RowValidating, AddressOf DA2_RowValidating
            If skip(1) Then DA2.Columns(dgvcell.ColumnIndex).Visible = True
            DA2.Select()
            If skip(1) Then DA2.CurrentCell = dgvcell
            RemoveHandler DA2.CellBeginEdit, AddressOf DA2_CellBeginEdit
            DA2.BeginEdit(False)
            AddHandler DA2.CellBeginEdit, AddressOf DA2_CellBeginEdit
            AddHandler DA2.CellEndEdit, AddressOf DA2_CellEndEdit
            AddHandler DA2.RowValidating, AddressOf DA2_RowValidating
        End If
    End Sub
    Private Sub B45_GotFocus(sender As Object, e As EventArgs) Handles B45.GotFocus
        If skip(0) Then
            RemoveHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
            RemoveHandler DA9.RowValidating, AddressOf DA9_RowValidating
            If skip(1) Then DA9.Columns(dgvcell.ColumnIndex).Visible = True
            DA9.Select()
            If skip(1) Then DA9.CurrentCell = dgvcell
            RemoveHandler DA9.CellBeginEdit, AddressOf DA9_CellBeginEdit
            DA9.BeginEdit(False)
            AddHandler DA9.CellBeginEdit, AddressOf DA9_CellBeginEdit
            AddHandler DA9.CellEndEdit, AddressOf DA9_CellEndEdit
            AddHandler DA9.RowValidating, AddressOf DA9_RowValidating
        End If
    End Sub
    Private Sub DA12_GotFocus(sender As Object, e As EventArgs)
        If skip(0) Then
            Dim DA As DataGridView = DirectCast(sender, DataGridView)
            RemoveHandler DA.GotFocus, AddressOf DA12_GotFocus
            RemoveHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
            RemoveHandler DA.RowValidating, AddressOf DA12_RowValidating
            CB1.Focus()
            If skip(1) Then DA.Columns(dgvcell.ColumnIndex).Visible = True
            DA.Select()
            If skip(1) Then DA.CurrentCell = dgvcell
            RemoveHandler DA.CellBeginEdit, AddressOf DA12_CellBeginEdit
            DA.BeginEdit(False)
            AddHandler DA.CellBeginEdit, AddressOf DA12_CellBeginEdit
            AddHandler DA.CellEndEdit, AddressOf DA12_CellEndEdit
            AddHandler DA.RowValidating, AddressOf DA12_RowValidating
        End If
    End Sub
    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If TC1.SelectedIndex = 5 Then
            WindowState = FormWindowState.Normal
            RemoveHandler SizeChanged, AddressOf Form1_SizeChanged
            Size = New Size(1114, 663)
            AddHandler SizeChanged, AddressOf Form1_SizeChanged
            CenterToScreen()
        End If
    End Sub
    Private Sub CH_MouseUp(sender As Object, e As MouseEventArgs) Handles CH20.MouseUp, CH21.MouseUp, CH22.MouseUp
        Dim CH As CheckBox = DirectCast(sender, CheckBox)
        Select Case CH.Text
            Case "平均核算表"
                CH.Text = "阶段核算表"
                CH.CheckState = CheckState.Indeterminate
            Case "阶段核算表"
                CH.Text = "消耗核算表"
                CH.Checked = True
            Case "消耗核算表"
                If CH.Checked Then
                    CH.Text = "平均核算表"
                    CH.CheckState = CheckState.Indeterminate
                End If
        End Select
        CH.Tag = IIf(CH.Checked, IIf(CH.Text = "消耗核算表", {"平均核算表", "阶段核算表"}, {CH.Text}), {""})
    End Sub
    Private Sub DA11_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA11.CellBeginEdit
        If e.ColumnIndex = 2 Then
            If DirectCast(sender, DataGridView).Rows(e.RowIndex).Cells(e.ColumnIndex).Tag IsNot Nothing Then
                sv = DirectCast(sender, DataGridView).Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                DirectCast(sender, DataGridView).Rows(e.RowIndex).Cells(e.ColumnIndex).Value = DirectCast(sender, DataGridView).Rows(e.RowIndex).Cells(e.ColumnIndex).Tag
            End If
        End If
    End Sub
    Private Sub DA10_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DA10.CellValueChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.ColumnIndex = 2 Then
            DirectCast(sender, DataGridView).Rows(e.RowIndex).Cells(2).Tag = Nothing
        ElseIf e.ColumnIndex = 3 Then
            Dim dec As Decimal
            If Decimal.TryParse(CStr(DA.Rows(e.RowIndex).Cells(3).Value), dec) Then
                DA.Rows(e.RowIndex).Cells(3).Value = dec
            ElseIf CStr(DA.Rows(e.RowIndex).Cells(3).Value) = "" Then
                DA.Rows(e.RowIndex).Cells(3).Value = Nothing
            End If
            For i = 0 To DA.Rows.Count - 2
                If DA.Rows(i).Cells(3).Value IsNot Nothing AndAlso Not IsNumeric(DA.Rows(i).Cells(3).Value) Then
                    DA.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                    Return
                End If
            Next
            DA.Columns(3).SortMode = DataGridViewColumnSortMode.Automatic
        End If
    End Sub
    Private Sub DA10_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA10.CellEndEdit
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.ColumnIndex = 0 Then
            If DA.NewRowIndex = e.RowIndex Then
                ni = 10
                ri = e.RowIndex
                TM1.Enabled = True
            End If
        ElseIf e.ColumnIndex = 3 Then
            Dim dec As Decimal
            If Decimal.TryParse(CStr(DA.Rows(e.RowIndex).Cells(3).Value), dec) Then
                DA.Rows(e.RowIndex).Cells(3).Value = dec
            ElseIf CStr(DA.Rows(e.RowIndex).Cells(3).Value) = "" Then
                DA.Rows(e.RowIndex).Cells(3).Value = Nothing
            End If
            For i = 0 To DA.Rows.Count - 2
                If DA.Rows(i).Cells(3).Value IsNot Nothing AndAlso Not IsNumeric(DA.Rows(i).Cells(3).Value) Then Return
            Next
            DA.Columns(3).SortMode = DataGridViewColumnSortMode.Automatic
        End If
    End Sub
End Class