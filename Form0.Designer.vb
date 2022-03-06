<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form0
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form0))
        Me.L1 = New System.Windows.Forms.TextBox()
        Me.L10 = New System.Windows.Forms.Label()
        Me.L4 = New System.Windows.Forms.Label()
        Me.L6 = New System.Windows.Forms.Label()
        Me.T3 = New System.Windows.Forms.TextBox()
        Me.CL1 = New System.Windows.Forms.CheckedListBox()
        Me.L8 = New System.Windows.Forms.Label()
        Me.C2 = New System.Windows.Forms.ComboBox()
        Me.L9 = New System.Windows.Forms.Label()
        Me.T4 = New System.Windows.Forms.TextBox()
        Me.CH1 = New System.Windows.Forms.CheckBox()
        Me.L3 = New System.Windows.Forms.Label()
        Me.L7 = New System.Windows.Forms.Label()
        Me.B1 = New System.Windows.Forms.Button()
        Me.B2 = New System.Windows.Forms.Button()
        Me.C1 = New System.Windows.Forms.ComboBox()
        Me.P1 = New System.Windows.Forms.PictureBox()
        Me.L2 = New System.Windows.Forms.Label()
        Me.T1 = New System.Windows.Forms.TextBox()
        Me.T5 = New System.Windows.Forms.TextBox()
        Me.L11 = New System.Windows.Forms.Label()
        Me.L12 = New System.Windows.Forms.Label()
        Me.C3 = New System.Windows.Forms.ComboBox()
        Me.B3 = New System.Windows.Forms.Button()
        Me.TT1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.P1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'L1
        '
        Me.L1.BackColor = System.Drawing.Color.FromArgb(CType(CType(162, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(155, Byte), Integer))
        Me.L1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.L1.Font = New System.Drawing.Font("等线", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(220, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.L1.Location = New System.Drawing.Point(18, 40)
        Me.L1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.L1.Name = "L1"
        Me.L1.Size = New System.Drawing.Size(999, 32)
        Me.L1.TabIndex = 10
        Me.L1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'L10
        '
        Me.L10.AutoSize = True
        Me.L10.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L10.Location = New System.Drawing.Point(975, 14)
        Me.L10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L10.Name = "L10"
        Me.L10.Size = New System.Drawing.Size(43, 29)
        Me.L10.TabIndex = 1
        Me.L10.Text = "×"
        '
        'L4
        '
        Me.L4.AutoSize = True
        Me.L4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L4.Location = New System.Drawing.Point(34, 426)
        Me.L4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L4.Name = "L4"
        Me.L4.Size = New System.Drawing.Size(106, 24)
        Me.L4.TabIndex = 14
        Me.L4.Text = "登录名："
        '
        'L6
        '
        Me.L6.AutoSize = True
        Me.L6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L6.Location = New System.Drawing.Point(34, 484)
        Me.L6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L6.Name = "L6"
        Me.L6.Size = New System.Drawing.Size(106, 24)
        Me.L6.TabIndex = 16
        Me.L6.Text = "新密码："
        '
        'T3
        '
        Me.T3.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.T3.Location = New System.Drawing.Point(129, 478)
        Me.T3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.T3.Name = "T3"
        Me.T3.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.T3.Size = New System.Drawing.Size(157, 32)
        Me.T3.TabIndex = 8
        '
        'CL1
        '
        Me.CL1.CheckOnClick = True
        Me.CL1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CL1.FormattingEnabled = True
        Me.CL1.Items.AddRange(New Object() {"全部"})
        Me.CL1.Location = New System.Drawing.Point(582, 362)
        Me.CL1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CL1.Name = "CL1"
        Me.CL1.Size = New System.Drawing.Size(128, 260)
        Me.CL1.TabIndex = 3
        Me.TT1.SetToolTip(Me.CL1, "用户配置时" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "可以在这里" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "限制工序。")
        '
        'L8
        '
        Me.L8.AutoSize = True
        Me.L8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L8.Location = New System.Drawing.Point(716, 424)
        Me.L8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L8.Name = "L8"
        Me.L8.Size = New System.Drawing.Size(82, 24)
        Me.L8.TabIndex = 17
        Me.L8.Text = "用户："
        '
        'C2
        '
        Me.C2.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.C2.FormattingEnabled = True
        Me.C2.Location = New System.Drawing.Point(784, 420)
        Me.C2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.C2.Name = "C2"
        Me.C2.Size = New System.Drawing.Size(205, 29)
        Me.C2.TabIndex = 1
        '
        'L9
        '
        Me.L9.AutoSize = True
        Me.L9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L9.Location = New System.Drawing.Point(716, 484)
        Me.L9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L9.Name = "L9"
        Me.L9.Size = New System.Drawing.Size(82, 24)
        Me.L9.TabIndex = 18
        Me.L9.Text = "密码："
        '
        'T4
        '
        Me.T4.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.T4.Location = New System.Drawing.Point(784, 478)
        Me.T4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.T4.Name = "T4"
        Me.T4.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.T4.Size = New System.Drawing.Size(205, 32)
        Me.T4.TabIndex = 2
        '
        'CH1
        '
        Me.CH1.AutoSize = True
        Me.CH1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CH1.Location = New System.Drawing.Point(722, 538)
        Me.CH1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CH1.Name = "CH1"
        Me.CH1.Size = New System.Drawing.Size(132, 28)
        Me.CH1.TabIndex = 4
        Me.CH1.Text = "记住密码"
        Me.CH1.UseVisualStyleBackColor = True
        '
        'L3
        '
        Me.L3.AutoSize = True
        Me.L3.Font = New System.Drawing.Font("微软雅黑", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L3.Location = New System.Drawing.Point(279, 358)
        Me.L3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L3.Name = "L3"
        Me.L3.Size = New System.Drawing.Size(133, 38)
        Me.L3.TabIndex = 12
        Me.L3.Text = "用户配置"
        '
        'L7
        '
        Me.L7.AutoSize = True
        Me.L7.Font = New System.Drawing.Font("微软雅黑", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L7.Location = New System.Drawing.Point(802, 358)
        Me.L7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L7.Name = "L7"
        Me.L7.Size = New System.Drawing.Size(133, 38)
        Me.L7.TabIndex = 13
        Me.L7.Text = "用户登录"
        '
        'B1
        '
        Me.B1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.B1.Location = New System.Drawing.Point(165, 530)
        Me.B1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.B1.Name = "B1"
        Me.B1.Size = New System.Drawing.Size(123, 45)
        Me.B1.TabIndex = 11
        Me.B1.Text = "确 定"
        Me.B1.UseVisualStyleBackColor = True
        '
        'B2
        '
        Me.B2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.B2.Location = New System.Drawing.Point(868, 531)
        Me.B2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.B2.Name = "B2"
        Me.B2.Size = New System.Drawing.Size(123, 45)
        Me.B2.TabIndex = 5
        Me.B2.Text = "确 定"
        Me.B2.UseVisualStyleBackColor = True
        '
        'C1
        '
        Me.C1.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.C1.FormattingEnabled = True
        Me.C1.Location = New System.Drawing.Point(129, 422)
        Me.C1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.C1.Name = "C1"
        Me.C1.Size = New System.Drawing.Size(157, 29)
        Me.C1.TabIndex = 6
        '
        'P1
        '
        Me.P1.Image = CType(resources.GetObject("P1.Image"), System.Drawing.Image)
        Me.P1.Location = New System.Drawing.Point(68, 111)
        Me.P1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.P1.Name = "P1"
        Me.P1.Size = New System.Drawing.Size(234, 220)
        Me.P1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.P1.TabIndex = 18
        Me.P1.TabStop = False
        '
        'L2
        '
        Me.L2.AutoSize = True
        Me.L2.Font = New System.Drawing.Font("Times New Roman", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L2.Location = New System.Drawing.Point(328, 128)
        Me.L2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L2.Name = "L2"
        Me.L2.Size = New System.Drawing.Size(69, 180)
        Me.L2.TabIndex = 11
        Me.L2.Text = "公" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "告 " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "栏"
        '
        'T1
        '
        Me.T1.BackColor = System.Drawing.Color.FromArgb(CType(CType(148, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(142, Byte), Integer))
        Me.T1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.T1.Location = New System.Drawing.Point(384, 111)
        Me.T1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.T1.Multiline = True
        Me.T1.Name = "T1"
        Me.T1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.T1.Size = New System.Drawing.Size(572, 218)
        Me.T1.TabIndex = 20
        '
        'T5
        '
        Me.T5.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.T5.Location = New System.Drawing.Point(414, 478)
        Me.T5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.T5.Name = "T5"
        Me.T5.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.T5.Size = New System.Drawing.Size(157, 32)
        Me.T5.TabIndex = 9
        '
        'L11
        '
        Me.L11.AutoSize = True
        Me.L11.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L11.Location = New System.Drawing.Point(296, 484)
        Me.L11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L11.Name = "L11"
        Me.L11.Size = New System.Drawing.Size(130, 24)
        Me.L11.TabIndex = 22
        Me.L11.Text = "确认密码："
        '
        'L12
        '
        Me.L12.AutoSize = True
        Me.L12.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.L12.Location = New System.Drawing.Point(344, 424)
        Me.L12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.L12.Name = "L12"
        Me.L12.Size = New System.Drawing.Size(82, 24)
        Me.L12.TabIndex = 22
        Me.L12.Text = "角色："
        '
        'C3
        '
        Me.C3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.C3.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.C3.FormattingEnabled = True
        Me.C3.Location = New System.Drawing.Point(414, 420)
        Me.C3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.C3.Name = "C3"
        Me.C3.Size = New System.Drawing.Size(157, 31)
        Me.C3.TabIndex = 10
        '
        'B3
        '
        Me.B3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.B3.Location = New System.Drawing.Point(414, 531)
        Me.B3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.B3.Name = "B3"
        Me.B3.Size = New System.Drawing.Size(123, 45)
        Me.B3.TabIndex = 12
        Me.B3.Text = "删 除"
        Me.B3.UseVisualStyleBackColor = True
        '
        'TT1
        '
        Me.TT1.AutoPopDelay = 5000
        Me.TT1.InitialDelay = 200
        Me.TT1.ReshowDelay = 500
        '
        'Form0
        '
        Me.AcceptButton = Me.B2
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(162, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(155, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1035, 654)
        Me.Controls.Add(Me.B3)
        Me.Controls.Add(Me.C3)
        Me.Controls.Add(Me.T5)
        Me.Controls.Add(Me.L11)
        Me.Controls.Add(Me.T1)
        Me.Controls.Add(Me.L2)
        Me.Controls.Add(Me.P1)
        Me.Controls.Add(Me.B2)
        Me.Controls.Add(Me.B1)
        Me.Controls.Add(Me.L7)
        Me.Controls.Add(Me.L3)
        Me.Controls.Add(Me.CH1)
        Me.Controls.Add(Me.C1)
        Me.Controls.Add(Me.C2)
        Me.Controls.Add(Me.L8)
        Me.Controls.Add(Me.CL1)
        Me.Controls.Add(Me.T3)
        Me.Controls.Add(Me.T4)
        Me.Controls.Add(Me.L6)
        Me.Controls.Add(Me.L4)
        Me.Controls.Add(Me.L10)
        Me.Controls.Add(Me.L1)
        Me.Controls.Add(Me.L9)
        Me.Controls.Add(Me.L12)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "Form0"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "众一合成一厂报表管理"
        CType(Me.P1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents L1 As System.Windows.Forms.TextBox
    Friend WithEvents L10 As System.Windows.Forms.Label
    Friend WithEvents L4 As System.Windows.Forms.Label
    Friend WithEvents L6 As System.Windows.Forms.Label
    Friend WithEvents T3 As System.Windows.Forms.TextBox
    Friend WithEvents CL1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents L8 As System.Windows.Forms.Label
    Friend WithEvents C2 As System.Windows.Forms.ComboBox
    Friend WithEvents L9 As System.Windows.Forms.Label
    Friend WithEvents T4 As System.Windows.Forms.TextBox
    Friend WithEvents CH1 As System.Windows.Forms.CheckBox
    Friend WithEvents L3 As System.Windows.Forms.Label
    Friend WithEvents L7 As System.Windows.Forms.Label
    Friend WithEvents B1 As System.Windows.Forms.Button
    Friend WithEvents B2 As System.Windows.Forms.Button
    Friend WithEvents P1 As System.Windows.Forms.PictureBox
    Friend WithEvents C1 As System.Windows.Forms.ComboBox
    Friend WithEvents L2 As System.Windows.Forms.Label
    Friend WithEvents T1 As System.Windows.Forms.TextBox
    Friend WithEvents T5 As System.Windows.Forms.TextBox
    Friend WithEvents L11 As System.Windows.Forms.Label
    Friend WithEvents L12 As System.Windows.Forms.Label
    Friend WithEvents C3 As System.Windows.Forms.ComboBox
    Friend WithEvents B3 As System.Windows.Forms.Button
    Friend WithEvents TT1 As ToolTip
End Class
