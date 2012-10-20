Public Class frmword
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            ' just to be sure!
            WinWordControl1.CloseControl()
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tpgjournal As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rtbSavejornal As System.Windows.Forms.Button
    Friend WithEvents WinWordControl1 As WinWordControl.WinWordControl
    Friend WithEvents btnrestoreword As System.Windows.Forms.Button
    Friend WithEvents btnpreactivate As System.Windows.Forms.Button
    Friend WithEvents btnload As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tpgjournal = New System.Windows.Forms.TabPage
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnrestoreword = New System.Windows.Forms.Button
        Me.btnpreactivate = New System.Windows.Forms.Button
        Me.btnload = New System.Windows.Forms.Button
        Me.WinWordControl1 = New WinWordControl.WinWordControl
        Me.rtbSavejornal = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.tpgjournal.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tpgjournal)
        Me.TabControl1.Location = New System.Drawing.Point(8, -32)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(432, 584)
        Me.TabControl1.TabIndex = 0
        '
        'tpgjournal
        '
        Me.tpgjournal.Controls.Add(Me.Panel1)
        Me.tpgjournal.Location = New System.Drawing.Point(4, 22)
        Me.tpgjournal.Name = "tpgjournal"
        Me.tpgjournal.Size = New System.Drawing.Size(424, 558)
        Me.tpgjournal.TabIndex = 0
        Me.tpgjournal.Text = "Journal Form"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnrestoreword)
        Me.Panel1.Controls.Add(Me.btnpreactivate)
        Me.Panel1.Controls.Add(Me.btnload)
        Me.Panel1.Controls.Add(Me.WinWordControl1)
        Me.Panel1.Controls.Add(Me.rtbSavejornal)
        Me.Panel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(2, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(400, 528)
        Me.Panel1.TabIndex = 0
        '
        'btnrestoreword
        '
        Me.btnrestoreword.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnrestoreword.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnrestoreword.Location = New System.Drawing.Point(296, 6)
        Me.btnrestoreword.Name = "btnrestoreword"
        Me.btnrestoreword.Size = New System.Drawing.Size(88, 24)
        Me.btnrestoreword.TabIndex = 4
        Me.btnrestoreword.Text = "Restore Word"
        '
        'btnpreactivate
        '
        Me.btnpreactivate.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnpreactivate.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnpreactivate.Location = New System.Drawing.Point(200, 6)
        Me.btnpreactivate.Name = "btnpreactivate"
        Me.btnpreactivate.Size = New System.Drawing.Size(88, 24)
        Me.btnpreactivate.TabIndex = 3
        Me.btnpreactivate.Text = "Pre Activate"
        '
        'btnload
        '
        Me.btnload.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnload.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnload.Location = New System.Drawing.Point(104, 5)
        Me.btnload.Name = "btnload"
        Me.btnload.Size = New System.Drawing.Size(88, 24)
        Me.btnload.TabIndex = 2
        Me.btnload.Text = "Load "
        '
        'WinWordControl1
        '
        Me.WinWordControl1.Location = New System.Drawing.Point(8, 40)
        Me.WinWordControl1.Name = "WinWordControl1"
        Me.WinWordControl1.Size = New System.Drawing.Size(384, 480)
        Me.WinWordControl1.TabIndex = 1
        '
        'rtbSavejornal
        '
        Me.rtbSavejornal.BackColor = System.Drawing.SystemColors.ControlLight
        Me.rtbSavejornal.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.rtbSavejornal.Location = New System.Drawing.Point(8, 5)
        Me.rtbSavejornal.Name = "rtbSavejornal"
        Me.rtbSavejornal.Size = New System.Drawing.Size(88, 24)
        Me.rtbSavejornal.TabIndex = 0
        Me.rtbSavejornal.Text = "Save Journal"
        '
        'frmword
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(448, 598)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "frmword"
        Me.Text = "Word"
        Me.TabControl1.ResumeLayout(False)
        Me.tpgjournal.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Sub btnrestoreword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrestoreword.Click
        Try
            Me.WinWordControl1.RestoreWord()
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnload.Click
        Try
            Dim ofd As New System.Windows.Forms.OpenFileDialog
            ofd.Filter = "Text files (*.txt)|*.txt|document (*.doc)|*.doc"
            ofd.Multiselect = False
            ofd.ShowDialog()
            Me.WinWordControl1.LoadDocument(ofd.FileName)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnpreactivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpreactivate.Click
        Try
            WinWordControl1.PreActivate()
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
End Class
