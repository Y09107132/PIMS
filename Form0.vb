Imports System.Data.SqlClient
Public Class Form0
    Public st() As String = My.Settings.setting.Split(CChar("|"))
    Dim gx, state1 As Boolean, flag As Integer, da As SqlDataAdapter, state As New List(Of Integer), js As String, cnctm As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & st(0) & ";password=" & st(1) & ";timeout=1")
    Private Sub Form0_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        C2.Items.Add(st(0))
        Try
            cnctm.Open()
            dr = New SqlCommand("sp_helpuser", cnctm).ExecuteReader
            While dr.Read
                If CInt(dr(5)) > 4 Then C2.Items.Add(dr(0))
            End While
            dr.Close()
            dr = New SqlCommand("select name from sys.sql_logins where is_disabled=0 and name<>'calc'", cnctm).ExecuteReader
            While dr.Read
                C1.Items.Add(dr(0))
            End While
            dr.Close()
            dr = New SqlCommand("select 操作工序 from 操作工序 where 可用性=1 order by Id", cnctm).ExecuteReader
            While dr.Read
                CL1.Items.Add(dr(0))
            End While
            dr.Close()
            dr = New SqlCommand("sp_helprole", cnctm).ExecuteReader
            While dr.Read
                If CInt(dr(1)) <= 16384 AndAlso CInt(dr(1)) > 0 Then C3.Items.Add(dr(0))
            End While
            cnctm.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Application.Exit()
        End Try
        C3.Items.Add("")
        s2("注意事项", T1)
        s2("系统标题", L1)
        C1.Items.Remove(st(0))
        If Screen.PrimaryScreen.Bounds.Width <= Width OrElse Screen.PrimaryScreen.Bounds.Height <= Height Then MsgBox("屏幕分辨率不得小于690×436！")
        TT1.SetToolTip(C3, "MDataReader有物料数据访问权限" & vbCrLf & "Joperator有录入数据权限" & vbCrLf & "PDataReader有工艺数据访问权限" & vbCrLf & "Soperator有数据访问和录入权限" & vbCrLf & "DataReader有数据访问权限" & vbCrLf & "db_owner是数据库管理员" & vbCrLf & "所有权限归" & st(0) & "拥有")
    End Sub
    Private Sub C2_GotFocus(sender As Object, e As EventArgs) Handles C2.GotFocus
        AcceptButton = B2
    End Sub
    Private Sub C2_TextChanged(sender As Object, e As EventArgs) Handles C2.TextChanged
        Dim dt As New DataTable
        RemoveHandler T4.TextChanged, AddressOf L1_Text
        T4.Text = ""
        RemoveHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
        For i = 0 To CL1.Items.Count - 1
            CL1.SetItemChecked(i, False)
        Next
        Try
            cmd = New SqlCommand("select 密码 from 人员设置 where 操作人员=@操作人员 and 所用电脑=@所用电脑 and 用户名=@用户名", cnctm)
            cmd.Parameters.Add(New SqlParameter("操作人员", C2.Text))
            cmd.Parameters.Add(New SqlParameter("所用电脑", Environment.MachineName))
            cmd.Parameters.Add(New SqlParameter("用户名", Environment.UserName))
            cnctm.Open()
            dr = cmd.ExecuteReader
            CH1.Checked = False
            While dr.Read
                CH1.Checked = True
                T4.Text = CStr(dr(0))
            End While
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
        End Try
        AddHandler T4.TextChanged, AddressOf L1_Text
        B2.Enabled = Not (C1.Text <> "" AndAlso Fcsb.s1(C2.Text) = 1)
        state.Clear()
        Try
            cmd = New SqlCommand("select 操作工序 from 用户工序 where 操作人员=@操作人员", cnctm)
            cmd.Parameters.Add(New SqlParameter("操作人员", C2.Text))
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            state1 = dt.Rows.Count > 0
            For Each dtr As DataRow In dt.Rows
                For j = 1 To CL1.Items.Count - 1
                    If CL1.Items(j).ToString = dtr(0).ToString Then
                        CL1.SetItemChecked(j, True)
                        state.Add(j)
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
        End Try
        If CL1.CheckedItems.Count = CL1.Items.Count Then
            CL1.SetItemCheckState(0, CheckState.Checked)
        ElseIf CL1.CheckedItems.Count = 0 AndAlso state.Count > 0 OrElse CL1.Items.Count > CL1.CheckedItems.Count AndAlso CL1.CheckedItems.Count > 0 Then
            Try
                cnctm.Open()
                If Not CL1.Items.Count - 1 = CInt(New SqlCommand("select count(*) from 操作工序", cnctm).ExecuteScalar) OrElse CL1.Items.Count - 1 > state.Count Then
                    CL1.SetItemCheckState(0, CheckState.Indeterminate)
                ElseIf CL1.Items.Count - 1 = CInt(New SqlCommand("select count(*) from 操作工序", cnctm).ExecuteScalar) Then
                    CL1.SetItemCheckState(0, CheckState.Checked)
                End If
                cnctm.Close()
            Catch ex As Exception
                cnctm.Close()
            End Try
        End If
        AddHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
    End Sub
    Private Sub B2_Click(sender As Object, e As EventArgs) Handles B2.Click
        If C2.Text.Contains("]") OrElse C2.Text = "calc" OrElse C2.Text = "" OrElse C2.Text.Contains("'") Then MsgBox("用户名格式不规范！") : Return
        If T4.Text.Contains("'") Then MsgBox("密码中不能包含'") : Return
        Dim cnct As New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        If CL1.CheckedItems.Count = 1 AndAlso CL1.GetItemCheckState(0) = CheckState.Indeterminate Then
            MsgBox("请选择自己的工序！")
        Else
            Try
                cnct.Open()
                Dim i = New SqlCommand("select 1", cnct).ExecuteNonQuery
                Try
                    cmdstr = "delete from 人员设置 where 操作人员=@操作人员 and 所用电脑=@所用电脑 and 用户名=@用户名"
                    If CH1.Checked Then cmdstr += " insert into 人员设置 values(@操作人员,@密码,@所用电脑,@用户名)"
                    cnctm.Open()
                    cmd = New SqlCommand(cmdstr, cnctm)
                    cmd.Parameters.Add(New SqlParameter("操作人员", C2.Text))
                    cmd.Parameters.Add(New SqlParameter("密码", T4.Text))
                    cmd.Parameters.Add(New SqlParameter("所用电脑", Environment.MachineName))
                    cmd.Parameters.Add(New SqlParameter("用户名", Environment.UserName))
                    cmd.ExecuteNonQuery()
                    cnctm.Close()
                Catch ex As Exception
                    cnctm.Close()
                End Try
                cnct.Close()
                Form1.Show()
                Close()
            Catch ex As Exception
                cnct.Close()
                MsgBox("用户 " & C2.Text & " 登录失败")
                T4.Focus()
                Return
            End Try
            Try
                cnctm.Open()
                If Not CBool(New SqlCommand("select-1from sys.sql_logins where name='calc'", cnctm).ExecuteScalar) Then Dim j As Integer = New SqlCommand("exec sp_addlogin 'calc','','msdb' use msdb grant connect to guest", cnctm).ExecuteNonQuery
                cnctm.Close()
            Catch ex As Exception
                MsgBox("创建计算用账户失败！" & vbCrLf & ex.Message)
                Application.Exit()
            End Try
        End If
    End Sub
    Private Sub B1_Click(sender As Object, e As EventArgs) Handles B1.Click
        Dim sign As Boolean = True
        s1(sign)
        If sign Then
            Dim cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text & ";timeout=1")
            Dim da As SqlDataAdapter
            Dim dt1 As New DataTable
            Dim dt2 As New DataTable
            Dim bl As Boolean
            If Not (CL1.GetItemCheckState(0) = CheckState.Indeterminate AndAlso CL1.CheckedItems.Count = 1) Then
                cmd = New SqlCommand("select 操作工序 from 用户工序 where 操作人员=@操作人员 order by 操作工序", cnctm)
                cmd.Parameters.Add(New SqlParameter("操作人员", C1.Text))
                da = New SqlDataAdapter(cmd)
                da.Fill(dt1)
                Try
                    cnct.Open()
                    cmd = New SqlCommand("delete from 用户工序 where 操作人员=@操作人员", cnct)
                    cmd.Parameters.Add(New SqlParameter("操作人员", C1.Text))
                    cmd.ExecuteNonQuery()
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                    Return
                End Try
                dt2.Columns.Add("操作工序")
                For i = 1 To CL1.CheckedItems.Count - 1
                    dt2.Rows.Add(CL1.CheckedItems(i).ToString)
                    Try
                        cnct.Open()
                        cmd = New SqlCommand("insert into 用户工序 values(@操作人员,@操作工序)", cnct)
                        cmd.Parameters.Add(New SqlParameter("操作人员", C1.Text))
                        cmd.Parameters.Add(New SqlParameter("操作工序", CL1.CheckedItems(i).ToString))
                        cmd.ExecuteNonQuery()
                        cnct.Close()
                    Catch ex As Exception
                        cnct.Close()
                    End Try
                Next
                If gx Then
                    bl = True
                    gx = False
                Else
                    Dim dtr1() As DataRow = dt1.Select("")
                    Dim dtr2() As DataRow = dt2.Select("", "操作工序")
                    If dtr1.Length <> dtr2.Length Then
                        bl = True
                    Else
                        For i = 0 To Math.Min(dtr1.Length, dtr2.Length) - 1
                            If dtr1(i)(0).ToString <> dtr2(i)(0).ToString Then
                                bl = True
                                Exit For
                            End If
                        Next
                    End If
                End If
                If bl Then MsgBox("用户工序更新成功！")
            End If
        End If
    End Sub
    Private Sub L10_Click(sender As Object, e As EventArgs) Handles L10.Click
        Close()
    End Sub
    Private Sub T1_GotFocus(sender As Object, e As EventArgs) Handles T1.GotFocus
        s2("注意事项", T1)
        AcceptButton = Nothing
    End Sub
    Private Sub T1_LostFocus(sender As Object, e As EventArgs) Handles T1.LostFocus
        s3("注意事项", "公告栏", T1)
    End Sub
    Private Sub L1_Text(sender As Object, e As EventArgs) Handles L1.GotFocus, C2.TextChanged, T4.TextChanged
        If T4.Text = st(1) AndAlso C2.Text = st(0) Then
            If sender Is L1 Then
                s2("系统标题", L1)
                AcceptButton = Nothing
            End If
            L1.ReadOnly = False
        Else
            L1.ReadOnly = True
        End If
        If sender Is L1 Then L1.Tag = L1.Text
    End Sub
    Private Sub L1_LostFocus(sender As Object, e As EventArgs) Handles L1.LostFocus
        s3("系统标题", "标题栏", L1)
    End Sub
    Private Sub C1_GotFocus(sender As Object, e As EventArgs) Handles C1.GotFocus
        AcceptButton = B1
    End Sub
    Public Sub C1_TextChanged(sender As Object, e As EventArgs) Handles C1.TextChanged
        Dim bl As Boolean = True
        Dim dt As New DataTable
        Dim C As ComboBox = DirectCast(sender, ComboBox)
        T3.Text = "" : T5.Text = ""
        If C.Items.Contains(C.Text) Then
            Try
                cnctm.Open()
                C3.Text = CStr(New SqlCommand("select a.name from sys.database_principals b,sys.database_principals a,sys.database_role_members roles where roles.member_principal_id=b.principal_id and roles.role_principal_id=a.principal_id and b.name='" & Replace(C.Text, "'", "''") & "'", cnctm).ExecuteScalar)
                cnctm.Close()
            Catch ex As Exception
                cnctm.Close()
            End Try
            If C3.Text <> "" Then bl = False : js = C3.Text
        End If
        If bl Then C3.Text = "" : js = ""
        B2.Enabled = Not (C.Text <> "" AndAlso Fcsb.s1(C2.Text) = 1)
        RemoveHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
        For j = 1 To CL1.Items.Count
            CL1.SetItemChecked(j - 1, False)
        Next
        AddHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
        state.Clear()
        Try
            cmd = New SqlCommand("select 操作工序 from 用户工序 where 操作人员=@操作人员", cnctm)
            If Fcsb.s1(C2.Text) = 1 Then
                cmd.Parameters.Add(New SqlParameter("操作人员", C2.Text))
            Else
                cmd.Parameters.Add(New SqlParameter("操作人员", C.Text))
            End If
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)
            state1 = dt.Rows.Count > 0
            For Each dtr As DataRow In dt.Rows
                For j = 1 To CL1.Items.Count - 1
                    If CL1.Items(j).ToString = dtr(0).ToString Then
                        CL1.SetItemChecked(j, True)
                        state.Add(j)
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
        End Try
        RemoveHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
        If CL1.CheckedItems.Count = CL1.Items.Count Then
            CL1.SetItemCheckState(0, CheckState.Checked)
        ElseIf CL1.CheckedItems.Count = 0 AndAlso state.Count > 0 OrElse CL1.Items.Count > CL1.CheckedItems.Count AndAlso CL1.CheckedItems.Count > 0 Then
            Try
                cnctm.Open()
                If Not CL1.Items.Count - 1 = CInt(New SqlCommand("select count(*) from 操作工序", cnctm).ExecuteScalar) Then CL1.SetItemCheckState(0, CheckState.Indeterminate)
                cnctm.Close()
            Catch ex As Exception
                cnctm.Close()
            End Try
        End If
        AddHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
    End Sub
    Public Sub CL1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CL1.ItemCheck
        Dim CL As CheckedListBox = DirectCast(sender, CheckedListBox)
        RemoveHandler CL.ItemCheck, AddressOf CL1_ItemCheck
        If e.Index = 0 AndAlso Fcsb.s1(C2.Text) > 0 AndAlso (Not CL.GetItemCheckState(0) = CheckState.Indeterminate AndAlso (state.Count = 0 OrElse state.Count = CL.Items.Count - 1) OrElse CL.GetItemCheckState(0) = CheckState.Indeterminate AndAlso state.Count = 0) OrElse e.Index = 0 AndAlso Fcsb.s1(C2.Text) = 0 Then
            If e.NewValue = CheckState.Checked Then
                For i = 1 To CL.Items.Count - 1
                    CL.SetItemChecked(i, True)
                Next
            ElseIf e.NewValue = CheckState.Unchecked Then
                For i = 1 To CL.Items.Count - 1
                    CL.SetItemChecked(i, False)
                Next
            End If
        ElseIf Not (e.Index = 0 AndAlso Fcsb.s1(C2.Text) > 0 AndAlso CL.GetItemCheckState(0) = CheckState.Indeterminate) Then
            s12(DirectCast(sender, CheckedListBox), e)
        End If
        Try
            cnctm.Open()
            If Not CL.Items.Count - 1 = CInt(New SqlCommand("select count(*) from 操作工序", cnctm).ExecuteScalar) Then CL.SetItemCheckState(0, CheckState.Indeterminate)
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
        End Try
        flag = e.Index
        AddHandler CL.ItemCheck, AddressOf CL1_ItemCheck
    End Sub
    Private Sub B3_Click(sender As Object, e As EventArgs) Handles B3.Click
        If C1.Text.Contains("]") OrElse C1.Text = "calc" OrElse C1.Text = "" OrElse C1.Text.Contains("'") Then MsgBox("登录名格式不规范！") : Return
        If T3.Text.Contains("'") OrElse T5.Text.Contains("'") Then MsgBox("密码格式不规范！") : Return
        If C2.Text.Contains("]") OrElse C2.Text = "calc" OrElse C2.Text = "" OrElse C2.Text.Contains("'") Then MsgBox("登录名格式不规范！(管理员)") : Return
        If T4.Text.Contains("'") Then MsgBox("密码格式不规范！(管理员)") : Return
        Dim cnct As New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        If MsgBox("是否删除用户？", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Try
                cnct.Open()
                cmd = New SqlCommand("drop user [" & C1.Text & "]", cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
                MsgBox("删除用户成功！")
                C2.Items.Remove(C1.Text)
                T3.Text = ""
                T5.Text = ""
                C3.Text = ""
                js = ""
            Catch ex As Exception
                cnct.Close()
                MsgBox("删除用户失败！" & Replace(ex.Message, "'", ""))
            End Try
        End If
        If MsgBox("是否删除登录名？", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Try
                cnct.Open()
                cmd = New SqlCommand("drop login [" & C1.Text & "]", cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
                C1.Items.Remove(C1.Text)
                C1.Text = ""
                MsgBox("删除登录名成功！")
            Catch ex As Exception
                cnct.Close()
                MsgBox("删除登录名失败！" & Replace(ex.Message, "'", ""))
            End Try
        End If
    End Sub
    Sub s1(ByRef sign As Boolean)
        If C1.Text.Contains("]") OrElse C1.Text = "calc" OrElse C1.Text = "" OrElse C1.Text.Contains("'") Then MsgBox("登录名格式不规范！") : sign = False : Return
        If T3.Text.Contains("'") OrElse T5.Text.Contains("'") Then MsgBox("密码格式不规范！") : sign = False : Return
        If C2.Text.Contains("]") OrElse C2.Text = "calc" OrElse C2.Text = "" OrElse C2.Text.Contains("'") Then MsgBox("登录名格式不规范！(授权方)") : sign = False : Return
        If T4.Text.Contains("'") Then MsgBox("密码格式不规范！(授权方)") : sign = False : Return
        Dim cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        Dim bl As Boolean = False
        If C1.Items.Contains(C1.Text) Then
            If Not (T3.Text = "" AndAlso T4.Text = "" AndAlso T5.Text = "") AndAlso C1.Text = C2.Text Then
                If T3.Text = T5.Text Then
                    Try
                        cmd = New SqlCommand("sp_password", cnct)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("old", T4.Text))
                        cmd.Parameters.Add(New SqlParameter("new", T3.Text))
                        cmd.Parameters.Add(New SqlParameter("loginame", C1.Text))
                        cnct.Open()
                        cmd.ExecuteNonQuery()
                        cnctm.Close()
                        T3.Text = "" : T4.Text = "" : T5.Text = ""
                        MsgBox("密码更改成功！")
                    Catch ex As Exception
                        cnct.Close()
                        sign = False
                        MsgBox("更改密码时发生错误，权限不足或连接失败！")
                        Return
                    End Try
                Else
                    sign = False
                    MsgBox("两次密码不一致，请重输")
                    Return
                End If
            End If
        Else
            If T3.Text <> T5.Text Then MsgBox("两次密码输入不一致！") : sign = False : Return
            Try
                cnct.Open()
                cmd = New SqlCommand("sp_addlogin", cnct)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("loginame", C1.Text))
                cmd.Parameters.Add(New SqlParameter("passwd", T3.Text))
                cmd.ExecuteNonQuery()
                If Not C1.Items.Contains(C1.Text) Then C1.Items.Add(C1.Text)
                cnct.Close()
                MsgBox("成功建立登录名！")
                T3.Text = "" : T5.Text = ""
            Catch ex As Exception
                cnct.Close()
                MsgBox("未能建立登录名！" & vbCrLf & Replace(ex.Message, "'", ""))
                sign = False
                Return
            End Try
            If C3.Text = "" Then js = "" : C2.Items.Remove(C1.Text) : sign = False : Return
        End If
        RemoveHandler C1.TextChanged, AddressOf C1_TextChanged
        If Fcsb.s1(C1.Text) = 1 AndAlso Fcsb.s1(C2.Text) = 1 AndAlso C3.Text <> "db_owner" Then
            MsgBox("角色db_owner不能相互更改。")
            sign = False
            C3.Text = js
            AddHandler C1.TextChanged, AddressOf C1_TextChanged
            Return
        End If
        AddHandler C1.TextChanged, AddressOf C1_TextChanged
        If Fcsb.s1(C2.Text) = 1 AndAlso C3.Text = "db_owner" AndAlso Fcsb.s1(C1.Text) > 1 Then
            If MsgBox("你正在尝试将 " & C1.Text & " 的权限升高到报表数据库的最高级别" & vbCrLf & "一旦生效你将没有权限将其降级，是否继续？", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                bl = True
            End If
        End If
        If bl Then
            sign = False
            C3.Text = js
            Return
        End If
        If js = C3.Text Then
            If js = "" Then
                sign = False
            End If
            Return
        End If
        Try
            cnct.Open()
            cmd = New SqlCommand("drop user[" & C1.Text & "]", cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
            If C3.Text = "" AndAlso js <> C3.Text Then
                MsgBox("去除角色成功！")
                js = ""
                C2.Items.Remove(C1.Text)
                C3.Text = ""
                sign = False
                Return
            End If
        Catch ex As Exception
            cnct.Close()
        End Try
        Try
            cnct.Open()
            cmd = New SqlCommand("create user[" & C1.Text & "]for login[" & C1.Text & "]with default_schema=dbo", cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        Try
            cnct.Open()
            cmd = New SqlCommand("sp_addrolemember", cnct)
            cmd.Parameters.Add(New SqlParameter("rolename", C3.Text))
            cmd.Parameters.Add(New SqlParameter("membername", C1.Text))
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteNonQuery()
            cnct.Close()
            MsgBox("角色赋予成功！")
            If js = "" Then
                gx = True
            End If
            If Not C2.Items.Contains(C1.Text) Then C2.Items.Add(C1.Text)
            js = C3.Text
        Catch ex As Exception
            cnct.Close()
            MsgBox("角色未能正常更改,因为你没有所需的权限！")
            C3.Text = js
            sign = False
            Return
        End Try
    End Sub
    Sub s2(ByRef str As String, T As TextBox)
        Try
            cnctm.Open()
            cmd = New SqlCommand("select " & str & " from 系统配置", cnctm)
            dr = cmd.ExecuteReader
            While dr.Read
                T.Text = CStr(dr(0))
            End While
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
            MsgBox(ex.Message)
            Application.Exit()
        End Try
    End Sub
    Sub s3(ByRef str1 As String, ByRef str2 As String, T As TextBox)
        Try
            cnctm.Open()
            cmd = New SqlCommand("update 系统配置 set " & str1 & "=@Content", cnctm)
            cmd.Parameters.Add(New SqlParameter("Content", T.Text))
            cmd.ExecuteNonQuery()
            cnctm.Close()
        Catch ex As Exception
            cnctm.Close()
            MsgBox(str2 & "可能没有正确更新" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub L1_KeyDown(sender As Object, e As KeyEventArgs) Handles L1.KeyDown
        If e.KeyCode = Keys.Escape Then DirectCast(sender, TextBox).Text = CStr(DirectCast(sender, TextBox).Tag)
    End Sub
    Private Sub CL1_MouseUp(sender As Object, e As MouseEventArgs) Handles CL1.MouseUp
        Dim CL As CheckedListBox = DirectCast(sender, CheckedListBox)
        If Fcsb.s1(C2.Text) > 0 Then
            If flag = 0 Then
                If CL.CheckedItems.Count = 0 AndAlso state1 AndAlso state.Count = 0 OrElse CL.CheckedItems.Count > 0 AndAlso CL.Items.Count <> CL.CheckedItems.Count Then CL.SetItemCheckState(0, CheckState.Indeterminate)
            Else
                RemoveHandler CL.ItemCheck, AddressOf CL1_ItemCheck
                If CL.CheckedItems.Count < 2 Then
                    If CL.CheckedItems.Count > 0 Then
                        CL.SetItemChecked(flag, True)
                    ElseIf Not CL.GetItemChecked(0) Then
                        CL.SetItemChecked(flag, Not CL.GetItemChecked(flag))
                    End If
                Else
                    CL.SetItemChecked(flag, (state.Contains(flag) OrElse Not state1) AndAlso CL.GetItemChecked(flag))
                End If
                AddHandler CL.ItemCheck, AddressOf CL1_ItemCheck
                Try
                    cnctm.Open()
                    If CL.Items.Count - 1 = CInt(New SqlCommand("select count(*) from 操作工序", cnctm).ExecuteScalar) Then
                        If CL.CheckedItems.Count < 2 Then
                            CL.SetItemCheckState(0, CheckState.Indeterminate)
                        ElseIf CL.CheckedItems.Count = CL.Items.Count Then
                            CL.SetItemCheckState(0, CheckState.Checked)
                        End If
                    Else
                        CL.SetItemCheckState(0, CheckState.Indeterminate)
                    End If
                    cnctm.Close()
                Catch ex As Exception
                    cnctm.Close()
                End Try
            End If
        End If
    End Sub
End Class