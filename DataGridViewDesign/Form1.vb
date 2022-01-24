Imports System.Data.OleDb

Public Class Form1
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;" & _
                            "Data source=DataGridView‬.accdb;")


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open()
        If con.State = ConnectionState.Open Then
            MsgBox("The database connection was successful ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Connection")
        End If
        Dim da As New OleDbDataAdapter("select * from table1", con)
        Dim dt As New DataTable
        da.Fill(dt)
        DataGridView1.Columns.Clear()
        DataGridView1.DataSource = dt
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text.Trim <> "" Then
            DataGridView_Design(DataGridView1, ComboBox1.Text, "")
            'DataGridView1      اسم من نوع داتا جريد فيو ليس نصياًّ
            'ComboBox1.Text     Index من0 إلى 8
            '""                 محاذات عناوين الأعمدة تأخذ احدى هذه القيم النصية (           ""              - "BC"            - "BL"       - "BR"          - "MC"         - "ML"                - "MR"            - "TC"             - "TL"              - "TR"   ) 
            '                                                                           أعلى يمين         أعلى يسار       أعلى في المركز  المنتصف الأيمن      المنتصف الأيسر   منتصف المركز   أسفل يمين    أسفل يسار       أسفل المركز         الافتراضي       
        End If
    End Sub

    Private Sub DataGridView_Design(ByVal DataGridView_1 As DataGridView, ByVal NumStyle As Integer, ByVal CellTextAlignment As String)
        Select Case NumStyle
            Case 0
                With DataGridView_1
                    'DataGridView خلفية
                    .BackgroundColor = Color.Gainsboro

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 4
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.MidnightBlue
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.White

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.MidnightBlue

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.White

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.MidnightBlue

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender
                End With
            Case 1

                With Me.DataGridView1
                    .BackgroundColor = Color.DarkGray

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    'لون الشبكة
                    .GridColor = Color.Silver

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.Black

                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.White

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.Black

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.White

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.Black

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    .MultiSelect = False
                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.Gainsboro
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke
                End With
            Case 2
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.Gainsboro

                    'لون الشبكة
                    .GridColor = Color.Silver

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    '\\\\\\\\\\\\\\\\\\\\\
                    'شكل إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    ' Me.FontHeight = 20

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGreen
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.White

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.DarkGreen

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.White

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.White
                End With
            Case 3
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.LightGoldenrodYellow

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    '\\\\\\\\\\\\\\\\\\\\\
                    'شكل إطار الخلية
                    .CellBorderStyle = 4
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    ' Me.FontHeight = 20

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.Maroon
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.LightGoldenrodYellow

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.Maroon

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.LightGoldenrodYellow

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.DarkSlateBlue

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                End With
            Case 4
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.Lavender

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    'لون الشبكة 
                    .GridColor = Color.RoyalBlue

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.MidnightBlue
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.White

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.MidnightBlue

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.White

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.MidnightBlue

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.White
                End With
            Case 5
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.Tan

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    'لون الشبكة 
                    .GridColor = Color.Tan

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.Wheat
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.SaddleBrown

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.Wheat

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.SaddleBrown

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.DarkSlateGray

                    'نوع وحجم خط الخلية
                    .RowsDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8)

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.OldLace
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.OldLace
                End With
            Case 6
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.Ivory

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 4

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 4

                    'لون الشبكة 
                    .GridColor = Color.Wheat

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.CadetBlue
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.Black

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.CadetBlue

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.Black

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.DarkSlateGray

                    'نوع وحجم خط الخلية
                    .RowsDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8)

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.White
                End With
            Case 7
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.DarkGray

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 2

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 2

                    'لون الشبكة 
                    .GridColor = Color.Silver

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.Silver
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.Black

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.Silver

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.Black

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.DarkSlateGray

                    'نوع وحجم خط الخلية
                    .RowsDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8)

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Silver
                End With
            Case 8
                With Me.DataGridView1
                    'DataGridView خلفية
                    .BackgroundColor = Color.DarkGray

                    'جعل Rows Header بدون إطار
                    .RowHeadersBorderStyle = 2

                    'جعل Column Header بدون إطار
                    .ColumnHeadersBorderStyle = 2

                    'لون الشبكة 
                    .GridColor = Color.WhiteSmoke

                    '\\\\\\\\\\\\\\\\\\\\\
                    'نوع إطار الخلية
                    .CellBorderStyle = 1
                    '1  Single
                    '2  Raised
                    '3  Sunken
                    '4  None
                    '5  SingleVertical
                    '6  RaisedVertical
                    '7  SunkenVertical
                    '8  SingleHorizontal
                    '9  RaisedHorizontal
                    '10 SunkenHorizontal
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'font
                    'نوع وحجم خط Column Header
                    .ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 12)

                    'Column Headerتلوين
                    .ColumnHeadersDefaultCellStyle.BackColor = Color.WhiteSmoke
                    'Column Headerتلوين خط 
                    .ColumnHeadersDefaultCellStyle.ForeColor = Color.Black

                    'Rows Headerتلوين
                    .RowHeadersDefaultCellStyle.BackColor = Color.WhiteSmoke

                    'Rows Headerتلوين خط 
                    .RowHeadersDefaultCellStyle.ForeColor = Color.Black

                    'تلوين خط الخلية
                    .RowsDefaultCellStyle.ForeColor = Color.DarkSlateGray

                    'نوع وحجم خط الخلية
                    .RowsDefaultCellStyle.Font = New Font("Tahoma", 8)

                    'محاذات النص في Column Header
                    Dim Al As DataGridViewContentAlignment
                    If CellTextAlignment.Trim = "" Then
                        Al = DataGridViewContentAlignment.NotSet
                    Else
                        Select Case CellTextAlignment
                            Case "BC"
                                Al = DataGridViewContentAlignment.BottomCenter
                            Case "BL"
                                Al = DataGridViewContentAlignment.BottomLeft
                            Case "BR"
                                Al = DataGridViewContentAlignment.BottomRight
                            Case "MC"
                                Al = DataGridViewContentAlignment.MiddleCenter
                            Case "ML"
                                Al = DataGridViewContentAlignment.MiddleLeft
                            Case "MR"
                                Al = DataGridViewContentAlignment.MiddleRight
                            Case "TC"
                                Al = DataGridViewContentAlignment.TopCenter
                            Case "TL"
                                Al = DataGridViewContentAlignment.TopLeft
                            Case "TR"
                                Al = DataGridViewContentAlignment.TopRight
                            Case Else
                                Al = DataGridViewContentAlignment.NotSet
                        End Select
                    End If
                    .ColumnHeadersDefaultCellStyle.Alignment = Al
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    'جعل Column Header مسطح
                    .EnableHeadersVisualStyles = False

                    'الاختيار المتعدد
                    .MultiSelect = False

                    'تلوين الصف الفردي
                    .RowsDefaultCellStyle.BackColor = Color.White
                    'تلوين الصف الزوجي
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.White
                End With
        End Select
    End Sub

End Class
