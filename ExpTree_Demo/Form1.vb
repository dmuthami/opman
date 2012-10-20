Imports System.Data
Imports System.Data.SqlClient

Public Class frmDataGrid
    Inherits System.Windows.Forms.Form
    'Please change these suit to your SQL Server settings.
    Private Const SQL_SERVER As String = "127.0.0.1"
    Private Const DATABASE As String = "TESC"
    Private Const USER_ID As String = "sa"
    Private Const PWD As String = "sa"

    'Capture the clicked cell
    Private hitTestGrid As DataGrid.HitTestInfo

    'Control definishion to add to DataGrid
    Private WithEvents datagridtextBox As DataGridTextBoxColumn
    Private WithEvents dataTable As DataTable
    Private WithEvents comboControl As System.Windows.Forms.ComboBox
    Private WithEvents dtp As New DateTimePicker
    Private WithEvents chk As New CheckBox

    'DataGrid Header Row
    Private arrstr() As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        InitializeComponent()
        'Method to create the customized data grid
        CreateGrid()
    End Sub 'New
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dgMember As System.Windows.Forms.DataGrid
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents TimeTextBox1 As AMS.TextBox.TimeTextBox
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.dgMember = New System.Windows.Forms.DataGrid
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.TimeTextBox1 = New AMS.TextBox.TimeTextBox
        CType(Me.dgMember, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(272, 152)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Load Grid"
        '
        'dgMember
        '
        Me.dgMember.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dgMember.CaptionVisible = False
        Me.dgMember.DataMember = ""
        Me.dgMember.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgMember.Location = New System.Drawing.Point(8, 16)
        Me.dgMember.Name = "dgMember"
        Me.dgMember.PreferredRowHeight = 20
        Me.dgMember.ReadOnly = True
        Me.dgMember.Size = New System.Drawing.Size(336, 128)
        Me.dgMember.TabIndex = 10
        '
        'TreeView1
        '
        Me.TreeView1.ImageIndex = -1
        Me.TreeView1.Location = New System.Drawing.Point(8, 165)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Nodes.AddRange(New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Process", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Casual"), New System.Windows.Forms.TreeNode("Accomodation"), New System.Windows.Forms.TreeNode("Travel")})})
        Me.TreeView1.SelectedImageIndex = -1
        Me.TreeView1.Size = New System.Drawing.Size(256, 139)
        Me.TreeView1.TabIndex = 11
        '
        'TimeTextBox1
        '
        Me.TimeTextBox1.Flags = 0
        Me.TimeTextBox1.Location = New System.Drawing.Point(80, 344)
        Me.TimeTextBox1.Name = "TimeTextBox1"
        Me.TimeTextBox1.RangeMax = New Date(1900, 1, 1, 23, 59, 59, 0)
        Me.TimeTextBox1.RangeMin = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.TimeTextBox1.ShowSeconds = False
        Me.TimeTextBox1.Size = New System.Drawing.Size(88, 20)
        Me.TimeTextBox1.TabIndex = 12
        '
        'frmDataGrid
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(352, 386)
        Me.Controls.Add(Me.TimeTextBox1)
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.dgMember)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.Name = "frmDataGrid"
        CType(Me.dgMember, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sconnectStr As String
        Try
            'Building the connection string
            sconnectStr = "Data Source=" & SQL_SERVER & ";Initial Catalog=" & DATABASE & ";User id=" & USER_ID & ";Password=" & PWD

            'Establish the connection to the data base and open it
            Dim sqlConn As New SqlConnection(sconnectStr)
            sqlConn.Open()

            'Create the sql command object and set its command type to execute the sql query to get the results
            Dim sc As New SqlCommand
            sc.Connection = sqlConn
            sc.CommandType = CommandType.Text
            sc.CommandText = "SELECT * FROM [DBT_CONTROL]"

            'Create the data set object to be used to fill the data grid with the data 
            Dim ds As New DataSet

            'Create the sql adapter that will be used to fill the data set created above
            Dim myReader As New SqlDataAdapter(sc)
            myReader.Fill(ds)

            'Fill the rows in the grid
            Dim i As Integer

            For i = 0 To (ds.Tables(0).Rows.Count) - 1
                dataTable.LoadDataRow(arrstr, True)
                dgMember(i, 0) = ds.Tables(0).Rows(i).ItemArray(0).ToString()
                dgMember(i, 1) = ds.Tables(0).Rows(i).ItemArray(1).ToString()
            Next i
            Button1.Enabled = False
        Catch ex As SqlException
            MsgBox(ex.Message().ToString())
        End Try
    End Sub

    Private Sub dgMember_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgMember.MouseDown
        Try
            hitTestGrid = dgMember.HitTest(e.X, e.Y)
            If Not (hitTestGrid Is Nothing) Then
                'Create the combo control to be added and set its properties
                comboControl = New ComboBox
                comboControl.Cursor = System.Windows.Forms.Cursors.Arrow
                comboControl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
                comboControl.Dock = DockStyle.Fill
                comboControl.Items.AddRange(New String(5) {"", "Information Technology", "Computer Science", "Bio Technology", "Electrical Engg", "aaa"})

                'Create the date time picker control to be added and set its properties
                dtp.Dock = DockStyle.Fill
                dtp.Cursor = Cursors.Arrow

                'Create the check box control to be added and set its properties
                chk.Dock = DockStyle.Fill
                chk.Cursor = Cursors.Arrow

                'Create the radio button control to be added and set its properties
                Dim rb As New RadioButton
                rb.Dock = DockStyle.Fill
                rb.Cursor = Cursors.Arrow

                'Add the controls to the respective columns in the data grid
                Dim i As Integer
                Dim sType As String

                chk.SendToBack()
                rb.SendToBack()
                dtp.SendToBack()
                comboControl.SendToBack()

                For i = 0 To dataTable.Rows.Count - 1
                    sType = dgMember(i, 0).ToString()
                    If hitTestGrid.Row = i Then
                        Select Case hitTestGrid.Row
                            Case 1
                                datagridtextBox.TextBox.Controls.Add(dtp)
                                dtp.BringToFront()
                            Case 0
                                datagridtextBox.TextBox.Controls.Add(comboControl)
                                comboControl.BringToFront()
                            Case 2
                                datagridtextBox.TextBox.Controls.Add(chk)
                                chk.BringToFront()
                            Case 3
                                datagridtextBox.TextBox.Controls.Add(rb)
                                rb.BringToFront()
                        End Select
                    End If
                    datagridtextBox.TextBox.BackColor = Color.White
                Next i
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub CreateGrid()
        'Declare and initialize local variables used
        Dim dtCol As DataColumn = Nothing 'Data Column variable
        'Create the String array object, initialize the array with the column names to be displayed
        arrstr = New String(1) {"Control Name", "Control"}

        'Create the Data Table object which will then be used to hold columns and rows
        dataTable = New DataTable("Controls")

        'Add the string array of columns to the DataColumn object 		
        Dim i As Integer
        For i = 0 To 1
            Dim str As String = arrstr(i)
            dtCol = New DataColumn(str)
            dtCol.DataType = System.Type.GetType("System.String")
            dtCol.DefaultValue = ""
            dataTable.Columns.Add(dtCol)
        Next i

        'Set the Data Grid Source as the Data Table created above
        dgMember.DataSource = dataTable

        'set style property when first time the grid loads, next time onwards it will maintain its property
        If Not dgMember.TableStyles.Contains("Controls") Then
            'Create a DataGridTableStyle object	
            Dim dgdtblStyle As New DataGridTableStyle
            'Set its properties
            dgdtblStyle.MappingName = dataTable.TableName 'its table name of dataset
            dgMember.TableStyles.Add(dgdtblStyle)
            dgdtblStyle.RowHeadersVisible = False
            dgdtblStyle.PreferredRowHeight = 22
            dgMember.BackgroundColor = Color.White

            'Take the columns in a GridColumnStylesCollection object and set the size of the individual columns	
            Dim colStyle As GridColumnStylesCollection
            colStyle = dgMember.TableStyles(0).GridColumnStyles
            colStyle(0).Width = 97
            colStyle(1).Width = 220
        End If
        'Take the text box from the second column of the grid where u will be adding the controls of your choice	
        datagridtextBox = CType(dgMember.TableStyles(0).GridColumnStyles(1), DataGridTextBoxColumn)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub comboControl_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles comboControl.SelectedValueChanged
        dgMember(hitTestGrid.Row, hitTestGrid.Column) = comboControl.Text
    End Sub

    Private Sub dtp_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp.ValueChanged
        dgMember(hitTestGrid.Row, hitTestGrid.Column) = dtp.Value.ToString
    End Sub

    Private Sub chk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chk.Click
        If (chk.Checked = True) Then
            dgMember(hitTestGrid.Row, hitTestGrid.Column) = "Selected"
        Else
            dgMember(hitTestGrid.Row, hitTestGrid.Column) = "Not Selected"
        End If
    End Sub
End Class
