Public Class Form2
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
    Friend WithEvents ColorCombo1 As VSEssentials.ColorCombo
    Friend WithEvents ColorList1 As VSEssentials.ColorList
    Friend WithEvents CoolList1 As VSEssentials.CoolList
    Friend WithEvents CustomEdit1 As VSEssentials.CustomEdit
    Friend WithEvents DirectoryEdit1 As VSEssentials.DirectoryEdit
    Friend WithEvents DirectoryTree1 As VSEssentials.DirectoryTree
    Friend WithEvents DirectoryTreeEdit1 As VSEssentials.DirectoryTreeEdit
    Friend WithEvents FileEdit1 As VSEssentials.FileEdit
    Friend WithEvents ImageToolBar1 As VSEssentials.ImageToolBar
    Friend WithEvents PowerListView1 As VSEssentials.PowerListView
    Friend WithEvents TabList1 As VSEssentials.TabList
    Friend WithEvents ThumbnailCombo1 As VSEssentials.ThumbnailCombo
    Friend WithEvents ThumbnailList1 As VSEssentials.ThumbnailList
    Friend WithEvents TitleList1 As VSEssentials.TitleList
    Friend WithEvents VsListView1 As VSEssentials.VSListView
    Friend WithEvents VsmdiNav1 As VSEssentials.VSMDINav
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents MultiLineCombo1 As VSEssentials.MultiLineCombo
    Friend WithEvents FlatCombo1 As VSEssentials.FlatCombo
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form2))
        Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, -1)
        Dim ListViewItem2 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, -1)
        Dim ListViewItem3 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))}, -1)
        Me.ColorCombo1 = New VSEssentials.ColorCombo()
        Me.ColorList1 = New VSEssentials.ColorList()
        Me.CoolList1 = New VSEssentials.CoolList()
        Me.CustomEdit1 = New VSEssentials.CustomEdit()
        Me.DirectoryEdit1 = New VSEssentials.DirectoryEdit()
        Me.DirectoryTree1 = New VSEssentials.DirectoryTree()
        Me.DirectoryTreeEdit1 = New VSEssentials.DirectoryTreeEdit()
        Me.FileEdit1 = New VSEssentials.FileEdit()
        Me.ImageToolBar1 = New VSEssentials.ImageToolBar()
        Me.PowerListView1 = New VSEssentials.PowerListView()
        Me.TabList1 = New VSEssentials.TabList()
        Me.ThumbnailCombo1 = New VSEssentials.ThumbnailCombo()
        Me.ThumbnailList1 = New VSEssentials.ThumbnailList()
        Me.TitleList1 = New VSEssentials.TitleList()
        Me.VsListView1 = New VSEssentials.VSListView()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.VsmdiNav1 = New VSEssentials.VSMDINav()
        Me.MultiLineCombo1 = New VSEssentials.MultiLineCombo()
        Me.FlatCombo1 = New VSEssentials.FlatCombo()
        Me.SuspendLayout()
        '
        'ColorCombo1
        '
        Me.ColorCombo1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ColorCombo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ColorCombo1.Flat = False
        Me.ColorCombo1.ItemHeight = 16
        Me.ColorCombo1.Items.AddRange(New Object() {"AliceBlue", "AntiqueWhite", "Aqua", "Aquamarine", "Azure", "Beige", "Bisque", "Black", "BlanchedAlmond", "Blue", "BlueViolet", "Brown", "BurlyWood", "CadetBlue", "Chartreuse", "Chocolate", "Coral", "CornflowerBlue", "Cornsilk", "Crimson", "Cyan", "DarkBlue", "DarkCyan", "DarkGoldenrod", "DarkGray", "DarkGreen", "DarkKhaki", "DarkMagenta", "DarkOliveGreen", "DarkOrange", "DarkOrchid", "DarkRed", "DarkSalmon", "DarkSeaGreen", "DarkSlateBlue", "DarkSlateGray", "DarkTurquoise", "DarkViolet", "DeepPink", "DeepSkyBlue", "DimGray", "DodgerBlue", "Firebrick", "FloralWhite", "ForestGreen", "Fuchsia", "Gainsboro", "GhostWhite", "Gold", "Goldenrod", "Gray", "Green", "GreenYellow", "Honeydew", "HotPink", "IndianRed", "Indigo", "Ivory", "Khaki", "Lavender", "LavenderBlush", "LawnGreen", "LemonChiffon", "LightBlue", "LightCoral", "LightCyan", "LightGoldenrodYellow", "LightGray", "LightGreen", "LightPink", "LightSalmon", "LightSeaGreen", "LightSkyBlue", "LightSlateGray", "LightSteelBlue", "LightYellow", "Lime", "LimeGreen", "Linen", "Magenta", "Maroon", "MediumAquamarine", "MediumBlue", "MediumOrchid", "MediumPurple", "MediumSeaGreen", "MediumSlateBlue", "MediumSpringGreen", "MediumTurquoise", "MediumVioletRed", "MidnightBlue", "MintCream", "MistyRose", "Moccasin", "NavajoWhite", "Navy", "OldLace", "Olive", "OliveDrab", "Orange", "OrangeRed", "Orchid", "PaleGoldenrod", "PaleGreen", "PaleTurquoise", "PaleVioletRed", "PapayaWhip", "PeachPuff", "Peru", "Pink", "Plum", "PowderBlue", "Purple", "Red", "RosyBrown", "RoyalBlue", "SaddleBrown", "Salmon", "SandyBrown", "SeaGreen", "SeaShell", "Sienna", "Silver", "SkyBlue", "SlateBlue", "SlateGray", "Snow", "SpringGreen", "SteelBlue", "Tan", "Teal", "Thistle", "Tomato", "Turquoise", "Violet", "Wheat", "White", "WhiteSmoke", "Yellow", "YellowGreen"})
        Me.ColorCombo1.Name = "ColorCombo1"
        Me.ColorCombo1.NormalColors = True
        Me.ColorCombo1.Size = New System.Drawing.Size(121, 22)
        Me.ColorCombo1.TabIndex = 0
        '
        'ColorList1
        '
        Me.ColorList1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ColorList1.ItemHeight = 16
        Me.ColorList1.Items.AddRange(New Object() {"AliceBlue", "AntiqueWhite", "Aqua", "Aquamarine", "Azure", "Beige", "Bisque", "Black", "BlanchedAlmond", "Blue", "BlueViolet", "Brown", "BurlyWood", "CadetBlue", "Chartreuse", "Chocolate", "Coral", "CornflowerBlue", "Cornsilk", "Crimson", "Cyan", "DarkBlue", "DarkCyan", "DarkGoldenrod", "DarkGray", "DarkGreen", "DarkKhaki", "DarkMagenta", "DarkOliveGreen", "DarkOrange", "DarkOrchid", "DarkRed", "DarkSalmon", "DarkSeaGreen", "DarkSlateBlue", "DarkSlateGray", "DarkTurquoise", "DarkViolet", "DeepPink", "DeepSkyBlue", "DimGray", "DodgerBlue", "Firebrick", "FloralWhite", "ForestGreen", "Fuchsia", "Gainsboro", "GhostWhite", "Gold", "Goldenrod", "Gray", "Green", "GreenYellow", "Honeydew", "HotPink", "IndianRed", "Indigo", "Ivory", "Khaki", "Lavender", "LavenderBlush", "LawnGreen", "LemonChiffon", "LightBlue", "LightCoral", "LightCyan", "LightGoldenrodYellow", "LightGray", "LightGreen", "LightPink", "LightSalmon", "LightSeaGreen", "LightSkyBlue", "LightSlateGray", "LightSteelBlue", "LightYellow", "Lime", "LimeGreen", "Linen", "Magenta", "Maroon", "MediumAquamarine", "MediumBlue", "MediumOrchid", "MediumPurple", "MediumSeaGreen", "MediumSlateBlue", "MediumSpringGreen", "MediumTurquoise", "MediumVioletRed", "MidnightBlue", "MintCream", "MistyRose", "Moccasin", "NavajoWhite", "Navy", "OldLace", "Olive", "OliveDrab", "Orange", "OrangeRed", "Orchid", "PaleGoldenrod", "PaleGreen", "PaleTurquoise", "PaleVioletRed", "PapayaWhip", "PeachPuff", "Peru", "Pink", "Plum", "PowderBlue", "Purple", "Red", "RosyBrown", "RoyalBlue", "SaddleBrown", "Salmon", "SandyBrown", "SeaGreen", "SeaShell", "Sienna", "Silver", "SkyBlue", "SlateBlue", "SlateGray", "Snow", "SpringGreen", "SteelBlue", "Tan", "Teal", "Thistle", "Tomato", "Turquoise", "Violet", "Wheat", "White", "WhiteSmoke", "Yellow", "YellowGreen"})
        Me.ColorList1.Location = New System.Drawing.Point(8, 48)
        Me.ColorList1.Name = "ColorList1"
        Me.ColorList1.NormalColors = True
        Me.ColorList1.Size = New System.Drawing.Size(120, 84)
        Me.ColorList1.TabIndex = 1
        '
        'CoolList1
        '
        Me.CoolList1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.CoolList1.ItemHeight = 20
        Me.CoolList1.Location = New System.Drawing.Point(320, 296)
        Me.CoolList1.Name = "CoolList1"
        Me.CoolList1.Size = New System.Drawing.Size(120, 84)
        Me.CoolList1.TabIndex = 3
        '
        'CustomEdit1
        '
        Me.CustomEdit1.DirectoryName = ""
        Me.CustomEdit1.Icon = CType(resources.GetObject("CustomEdit1.Icon"), System.Drawing.Icon)
        Me.CustomEdit1.Location = New System.Drawing.Point(16, 296)
        Me.CustomEdit1.Name = "CustomEdit1"
        Me.CustomEdit1.Size = New System.Drawing.Size(256, 24)
        Me.CustomEdit1.TabIndex = 4
        '
        'DirectoryEdit1
        '
        Me.DirectoryEdit1.DirectoryName = ""
        Me.DirectoryEdit1.Location = New System.Drawing.Point(24, 328)
        Me.DirectoryEdit1.Name = "DirectoryEdit1"
        Me.DirectoryEdit1.Size = New System.Drawing.Size(256, 24)
        Me.DirectoryEdit1.TabIndex = 5
        '
        'DirectoryTree1
        '
        Me.DirectoryTree1.Directory = "C:\"
        Me.DirectoryTree1.HideSelection = False
        Me.DirectoryTree1.Indent = 18
        Me.DirectoryTree1.ItemHeight = 18
        Me.DirectoryTree1.Location = New System.Drawing.Point(448, 176)
        Me.DirectoryTree1.Name = "DirectoryTree1"
        Me.DirectoryTree1.Nodes.AddRange(New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Desktop", 8, 8), New System.Windows.Forms.TreeNode("My Computer", 9, 9, New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("C:\", 0, 0), New System.Windows.Forms.TreeNode("D:\", 0, 0), New System.Windows.Forms.TreeNode("E:\", 0, 0), New System.Windows.Forms.TreeNode("F:\", 0, 0), New System.Windows.Forms.TreeNode("G:\", 2, 2), New System.Windows.Forms.TreeNode("H:\", 1, 1)}), New System.Windows.Forms.TreeNode("My Documents", 5, 6)})
        Me.DirectoryTree1.TabIndex = 6
        '
        'DirectoryTreeEdit1
        '
        Me.DirectoryTreeEdit1.ColWidth = 100
        Me.DirectoryTreeEdit1.Directory = ""
        Me.DirectoryTreeEdit1.DisplayCol = -1
        Me.DirectoryTreeEdit1.Filter = ""
        Me.DirectoryTreeEdit1.Icon = CType(resources.GetObject("DirectoryTreeEdit1.Icon"), System.Drawing.Icon)
        Me.DirectoryTreeEdit1.Location = New System.Drawing.Point(24, 392)
        Me.DirectoryTreeEdit1.Name = "DirectoryTreeEdit1"
        Me.DirectoryTreeEdit1.PopupHeight = 296
        Me.DirectoryTreeEdit1.PopupWidth = 336
        Me.DirectoryTreeEdit1.Selected = ""
        Me.DirectoryTreeEdit1.Size = New System.Drawing.Size(264, 24)
        Me.DirectoryTreeEdit1.TabIndex = 7
        '
        'FileEdit1
        '
        Me.FileEdit1.Filename = ""
        Me.FileEdit1.Location = New System.Drawing.Point(24, 360)
        Me.FileEdit1.Name = "FileEdit1"
        Me.FileEdit1.Size = New System.Drawing.Size(264, 24)
        Me.FileEdit1.TabIndex = 8
        '
        'ImageToolBar1
        '
        Me.ImageToolBar1.Location = New System.Drawing.Point(16, 136)
        Me.ImageToolBar1.Name = "ImageToolBar1"
        Me.ImageToolBar1.TabIndex = 9
        '
        'PowerListView1
        '
        Me.PowerListView1.BackFill_Gradient1 = System.Drawing.SystemColors.ControlLight
        Me.PowerListView1.BackFill_Gradient2 = System.Drawing.SystemColors.ControlDark
        Me.PowerListView1.BackFillGradient = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.PowerListView1.ItemNormalcolor = System.Drawing.SystemColors.Window
        Me.PowerListView1.ItemSelectedcolor = System.Drawing.SystemColors.Highlight
        Me.PowerListView1.Location = New System.Drawing.Point(184, 40)
        Me.PowerListView1.Name = "PowerListView1"
        Me.PowerListView1.TabIndex = 10
        '
        'TabList1
        '
        Me.TabList1.BackColor = System.Drawing.SystemColors.Control
        Me.TabList1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TabList1.Color_FrameColor = System.Drawing.SystemColors.ControlDark
        Me.TabList1.Color_OffColor = System.Drawing.SystemColors.Control
        Me.TabList1.Color_ONColor = System.Drawing.SystemColors.Window
        Me.TabList1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.TabList1.ItemHeight = 20
        Me.TabList1.Items.AddRange(New Object() {"ssss", "vv", "zz"})
        Me.TabList1.Location = New System.Drawing.Point(312, 40)
        Me.TabList1.Name = "TabList1"
        Me.TabList1.TabIndex = 11
        '
        'ThumbnailCombo1
        '
        Me.ThumbnailCombo1.Directory = Nothing
        Me.ThumbnailCombo1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.ThumbnailCombo1.Filter = "*.jpg"
        Me.ThumbnailCombo1.Flat = False
        Me.ThumbnailCombo1.Location = New System.Drawing.Point(312, 144)
        Me.ThumbnailCombo1.Name = "ThumbnailCombo1"
        Me.ThumbnailCombo1.Size = New System.Drawing.Size(121, 21)
        Me.ThumbnailCombo1.TabIndex = 12
        Me.ThumbnailCombo1.Text = "ThumbnailCombo1"
        Me.ThumbnailCombo1.ThumbHeight = 100
        '
        'ThumbnailList1
        '
        Me.ThumbnailList1.Directory = Nothing
        Me.ThumbnailList1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.ThumbnailList1.Filter = "*.jpg"
        Me.ThumbnailList1.ItemHeight = 200
        Me.ThumbnailList1.Location = New System.Drawing.Point(184, 144)
        Me.ThumbnailList1.Name = "ThumbnailList1"
        Me.ThumbnailList1.TabIndex = 13
        Me.ThumbnailList1.ThumbHeight = 100
        '
        'TitleList1
        '
        Me.TitleList1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.TitleList1.ItemHeight = 20
        Me.TitleList1.Location = New System.Drawing.Point(440, 48)
        Me.TitleList1.Name = "TitleList1"
        Me.TitleList1.TabIndex = 14
        '
        'VsListView1
        '
        Me.VsListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.VsListView1.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem1, ListViewItem2, ListViewItem3})
        Me.VsListView1.Location = New System.Drawing.Point(312, 168)
        Me.VsListView1.Name = "VsListView1"
        Me.VsListView1.TabIndex = 15
        '
        'VsmdiNav1
        '
        Me.VsmdiNav1.AutoActivate = True
        Me.VsmdiNav1.AutoHide = True
        Me.VsmdiNav1.Dock = System.Windows.Forms.DockStyle.Top
        Me.VsmdiNav1.Name = "VsmdiNav1"
        Me.VsmdiNav1.ShowIcon = False
        Me.VsmdiNav1.Size = New System.Drawing.Size(584, 25)
        Me.VsmdiNav1.TabIndex = 16
        '
        'MultiLineCombo1
        '
        Me.MultiLineCombo1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.MultiLineCombo1.Flat = False
        Me.MultiLineCombo1.ItemHeight = 16
        Me.MultiLineCombo1.Items.AddRange(New Object() {"eeee", "hg", "nb", "hjm"})
        Me.MultiLineCombo1.Location = New System.Drawing.Point(8, 432)
        Me.MultiLineCombo1.Name = "MultiLineCombo1"
        Me.MultiLineCombo1.Size = New System.Drawing.Size(424, 22)
        Me.MultiLineCombo1.TabIndex = 17
        Me.MultiLineCombo1.Text = "MultiLineCombo1"
        '
        'FlatCombo1
        '
        Me.FlatCombo1.Flat = False
        Me.FlatCombo1.Location = New System.Drawing.Point(8, 464)
        Me.FlatCombo1.Name = "FlatCombo1"
        Me.FlatCombo1.Size = New System.Drawing.Size(416, 21)
        Me.FlatCombo1.TabIndex = 18
        Me.FlatCombo1.Text = "FlatCombo1"
        '
        'Form2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 490)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.FlatCombo1, Me.MultiLineCombo1, Me.VsmdiNav1, Me.VsListView1, Me.TitleList1, Me.ThumbnailList1, Me.ThumbnailCombo1, Me.TabList1, Me.PowerListView1, Me.ImageToolBar1, Me.FileEdit1, Me.DirectoryTreeEdit1, Me.DirectoryTree1, Me.DirectoryEdit1, Me.CustomEdit1, Me.CoolList1, Me.ColorList1, Me.ColorCombo1})
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
