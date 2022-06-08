#Region "ABOUT"
' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: http://www.facebook.com/g2gnet (for Thailand)
' / Facebook: http://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com
' /
' / Purpose: Demonstrate search product from DataTable and calculate product sales results.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------
#End Region

Public Class frmDetailDataGridView

    ' / --------------------------------------------------------------------------------
    '/ Don't forget to set Form has KeyPreview = True
    Private Sub frmDetailDataGridView_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F8
                '/ Remove Row
                Call DeleteRow("btnDelRow")
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub frmDetailDataGridView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True  '/ สามารถกดปุ่มฟังค์ชั่นคีย์ลงในฟอร์มได้
        Call InitializeGrid()
        '/
        txtSumTotal.ReadOnly = True
        txtSumTotal.Text = "0.00"
        lblLastPrice.Text = ""
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub InitializeGrid()
        With dgvData
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .ReadOnly = False
            .RowTemplate.MinimumHeight = 27
            .RowTemplate.Height = 27
            '/ Columns Specified
            '/ Index = 0
            .Columns.Add("PK", "Primary Key")
            With .Columns("PK")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                .Visible = True 'False '/ ปกติหลัก Primary Key จะต้องถูกซ่อนไว้
            End With
            '/ Index = 1
            .Columns.Add("ProductID", "Product ID")
            .Columns("ProductID").ReadOnly = True
            '/ Index = 2
            .Columns.Add("ProductName", "Product Name")
            .Columns("ProductName").ReadOnly = True
            '/ Index = 3
            .Columns.Add("Quantity", "Quantity")
            .Columns("Quantity").ValueType = GetType(Integer)
            '/ Index = 4
            .Columns.Add("UnitPrice", "Unit Price")
            .Columns("UnitPrice").ValueType = GetType(Double)
            '/ Index = 5
            .Columns.Add("Total", "Total")
            .Columns("Total").ValueType = GetType(Double)
            .Font = New Font("Tahoma", 11)
            '/ Total Column
            With .Columns("Total")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                .DefaultCellStyle.ForeColor = Color.Blue
                .DefaultCellStyle.Font = New Font(dgvData.Font, FontStyle.Bold)
            End With
            '// เพิ่มปุ่มลบ (Index = 6)
            Dim btnDelRow As New DataGridViewButtonColumn
            dgvData.Columns.Add(btnDelRow)
            With btnDelRow
                .HeaderText = "Delete F8"
                .Text = "Delete"
                .UseColumnTextForButtonValue = True
                .Width = 60
                .ReadOnly = True
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .SortMode = DataGridViewColumnSortMode.NotSortable  '/ Not sort order but can click header for delete row.
            End With

            '/ Alignment MiddleRight only columns 3 to 5
            For i As Byte = 3 To 5
                '/ Header Alignment
                .Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                '/ Cell Alignment
                .Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Next
            '/ Auto size column width of each main by sorting the field.
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '/ Adjust Header Styles
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.RoyalBlue
                .ForeColor = Color.White
                .Font = New Font("Tahoma", 11, FontStyle.Bold)
            End With
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 36
            '/ กำหนดให้ EnableHeadersVisualStyles = False เพื่อให้ยอมรับการเปลี่ยนแปลงสีพื้นหลังของ Header
            .EnableHeadersVisualStyles = False
        End With

    End Sub

    ' / --------------------------------------------------------------------------------
    '// SAMPLE DATATABLE
    Private Function GetDataTable() As DataTable
        Dim DT As New DataTable
        '// Add Columns
        With DT.Columns
            .Add("PK", GetType(Integer))
            .Add("ProductID", GetType(String))
            .Add("ProductName", GetType(String))
            .Add("UnitPrice", GetType(Double))
        End With
        '// Add Rows
        With DT.Rows
            '// PK, ProductID, ProductName, UnitPrice
            .Add(1, "01", "Classic Checken", "45.00")
            .Add(2, "02", "Mexicana", "85.00")
            .Add(3, "03", "Lemon Shrimp", "110.00")
            .Add(4, "04", "Bacon", "90.00")
            .Add(5, "05", "Spicy Shrimp", "120.00")
            .Add(6, "06", "Tex Supreme", "80.00")
            .Add(7, "07", "Fish", "100.00")
            .Add(8, "08", "Pepsi Small", "20.00")
            .Add(9, "09", "Coke Small", "20.00")
            .Add(10, "10", "7Up Small", "20.00")
            .Add(11, "11", "มอคค่า", "85.00")
            .Add(12, "12", "อเมริกาโน่", "95.00")
            .Add(13, "13", "เอ็กซเพรสโซ่", "115.00")
            .Add(14, "14", "ชานมไข่มุก", "45.00")
            .Add(15, "15", "เหล้าเป็ก", "99.00")
        End With
        Return DT
    End Function

    ' / --------------------------------------------------------------------------------
    ' / การค้นหาข้อมูลในช่อง TextBox
    Private Sub txtSearch_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        '// เมื่อกดคีย์ ENTER เพื่อเริ่มต้นการค้นหาข้อมูล
        If e.KeyChar = Chr(13) Then
            '/ Replace some word for reserved in DataBase.
            txtSearch.Text = txtSearch.Text.Replace("'", "").Replace("*", "").Replace("%", "")
            e.Handled = True    '// ปิดเสียง
            '/ สร้าง DataTable สมมุติขึ้นมา
            Dim DT As DataTable = GetDataTable()
            '/ ค้นหาข้อมูลจาก DataTable แล้วรับค่ามาใส่ไว้ใน DataRow
            '/ การค้นหาข้อมูลแบบ String จะต้องใส่เครื่องหมาย Single Quote ครอบเอาไว้ เช่น ProductID = '01'
            Dim r() As DataRow = DT.Select(" ProductID = " & "'" & txtSearch.Text.Trim & "'")
            '// หากพบข้อมูลใน DataTable
            If r.Count > 0 Then
                '/ ตัวแปรบูลีน Flag แจ้งการค้นหาข้อมูลในตารางกริด (True = พบรายการในแถว, False = ไม่พบ)
                Dim blnExist As Boolean = False
                '/ ต้องค้นหาข้อมูลจากตารางกริดก่อน เพื่อค้นหาว่ามีรายการสินค้าเดิมหรือไม่?
                '/ หากในตารางกริดยังไม่มีแถวรายการ ก็จะข้าม For Loop นี้ไปเพิ่มรายการใหม่ทันที (ครั้งแรกที่ยังไม่มีข้อมูลในตารางกริด)
                For iRow As Integer = 0 To dgvData.RowCount - 1
                    '/ ทดสอบด้วย Primary Key r(0).Item(0) หรือ Product ID r(0).Item(1) ก็ได้
                    If r(0).Item(0) = dgvData.Rows(iRow).Cells(0).Value Then
                        '/ เมื่อพบรายการเดิม ก็ให้เพิ่มจำนวนขึ้น 1 
                        dgvData.Rows(iRow).Cells(3).Value += 1
                        lblLastPrice.Text = "Last Price: " & CDbl(dgvData.Rows(iRow).Cells(4).Value).ToString("0.00")
                        '/ Flag แจ้งว่าพบข้อมูลเดิมแล้ว
                        blnExist = True
                        '/ เมื่อเจอสินค้าเดิมในตารางกริดแล้ว ไม่ว่าจะอยู่แถวลำดับที่เท่าไหร่ ก็ให้ออกจาก For Loop การค้นหาได้เลย
                        '/ เพราะรายการสินค้าใดๆ จะต้องมีอยู่เพียงแค่รายการเดียว ไม่ต้องเสียเวลาวนรอบกลับไปทำให้จนครบจำนวนแถว
                        Exit For
                    End If
                Next
                '/ กรณีที่พบสินค้าในตารางกริด กำหนด blnExist = True ทำให้ Not True = False จะทำให้ข้ามเงื่อนไขนี้ออกไป
                '/ กรณีที่ไม่พบข้อมูลสินค้าเดิมในตารางกริด กำหนด blnExist = False ทำให้ Not False = True เพิ่มรายการสินค้าแถวใหม่เข้าไปในตารางกริดได้
                If Not blnExist Then
                    '/ Primary Key, Product ID, Product Name, Quantity, UnitPrice, Total
                    dgvData.Rows.Add(r(0).Item(0), r(0).Item(1), r(0).Item(2), "1", Format(CDbl(r(0).Item(3).ToString), "0.00"), "0.00")
                    lblLastPrice.Text = "Last Price: " & CDbl(r(0).Item(3)).ToString("0.00")
                End If
                '/ หากไม่ใช้ NOT ก็จะต้องเขียนโปรแกรมแบบนี้
                '/ If blnExist = True Then
                '/     ไม่ต้องทำอะไร
                '/ Else
                '/     ทำคำสั่งเพิ่มรายการ
                '/ End If
                '// คำนวณผลรวมใหม่
                Call CalSumTotal()
                DT.Dispose()
            End If
            txtSearch.Clear()
            txtSearch.Focus()
        End If
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Calcualte sum of Total (Column Index = 5)
    ' / ทำทุกครั้งที่มีการเพิ่มหรือลบแถวรายการ และมีการเปลี่ยนแปลงค่าในเซลล์ Quantity, UnitPrice
    Private Sub CalSumTotal()
        txtSumTotal.Text = "0.00"
        '/ วนรอบตามจำนวนแถวที่มีอยู่ปัจจุบัน
        For i As Integer = 0 To dgvData.RowCount - 1
            '/ หลักสุดท้ายของตารางกริด = [จำนวน x ราคา]
            dgvData.Rows(i).Cells(5).Value = Format(dgvData.Rows(i).Cells(3).Value * dgvData.Rows(i).Cells(4).Value, "#,##0.00")
            '/ นำค่าจาก Total มารวมกันเพื่อแสดงผลในสรุปผลรวม (x = x + y)
            txtSumTotal.Text = Format(CDbl(txtSumTotal.Text) + CDbl(dgvData.Rows(i).Cells(5).Value), "#,##0.00")
        Next
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / โปรแกรมย่อยในการลบแถวรายการที่เลือกออกไป
    Private Sub DeleteRow(ByVal ColName As String)
        If dgvData.RowCount = 0 Then Return
        '/ ColName เป็นชื่อของหลัก Index = 6 ของตารางกริด (ไปดูที่โปรแกรมย่อย InitializeGrid)
        If ColName = "btnDelRow" Then
            '// ลบรายการแถวที่เลือกออกไป
            dgvData.Rows.Remove(dgvData.CurrentRow)
            '/ เมื่อแถวรายการถูกลบออกไป ต้องไปคำนวณหาค่าผลรวมใหม่
            Call CalSumTotal()
        End If
        txtSearch.Focus()
    End Sub

    Private Sub dgvData_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellClick
        Select Case e.ColumnIndex
            '// Delete Button
            Case 6
                'MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)
                Call DeleteRow("btnDelRow")
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / After you press Enter
    Private Sub dgvData_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellEndEdit
        '/ เกิดการเปลี่ยนแปลงค่าในหลัก Index ที่ 3 หรือ 4
        Select Case e.ColumnIndex
            Case 3, 4 '/ Column Index = 3 (Quantity), Column Index = 4 (UnitPrice)
                '/ Quantity
                '/ การดัก Error กรณีมีค่า Null Value ให้ใส่ค่า 0 ลงไปแทน
                If IsDBNull(dgvData.Rows(e.RowIndex).Cells(3).Value) Then dgvData.Rows(e.RowIndex).Cells(3).Value = "0"
                Dim Quantity As Integer = dgvData.Rows(e.RowIndex).Cells(3).Value
                '/ UnitPrice
                '/ If Null Value
                If IsDBNull(dgvData.Rows(e.RowIndex).Cells(4).Value) Then dgvData.Rows(e.RowIndex).Cells(4).Value = "0.00"
                Dim UnitPrice As Double = dgvData.Rows(e.RowIndex).Cells(4).Value
                dgvData.Rows(e.RowIndex).Cells(4).Value = Format(CDbl(dgvData.Rows(e.RowIndex).Cells(4).Value), "0.00")

                '/ Quantity x UnitPrice
                dgvData.Rows(e.RowIndex).Cells(5).Value = CDbl((Quantity * UnitPrice).ToString("#,##0.00"))
                '/ Calculate Summary
                Call CalSumTotal()
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub dgvData_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvData.EditingControlShowing
        Select Case dgvData.Columns(dgvData.CurrentCell.ColumnIndex).Name
            ' / Can use both Colume Index or Field Name
            Case "Quantity", "UnitPrice"
                '/ Stop and Start event handler
                RemoveHandler e.Control.KeyPress, AddressOf ValidKeyPress
                AddHandler e.Control.KeyPress, AddressOf ValidKeyPress
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub ValidKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim tb As TextBox = sender
        Select Case dgvData.CurrentCell.ColumnIndex
            Case 3  ' Quantity is Integer
                Select Case e.KeyChar
                    Case "0" To "9"   ' digits 0 - 9 allowed
                    Case ChrW(Keys.Back)    ' backspace allowed for deleting (Delete key automatically overrides)

                    Case Else ' everything else ....
                        ' True = CPU cancel the KeyPress event
                        e.Handled = True ' and it's just like you never pressed a key at all
                End Select

            Case 4  ' UnitPrice is Double
                Select Case e.KeyChar
                    Case "0" To "9"
                        ' Allowed "."
                    Case "."
                        ' can present "." only one
                        If InStr(tb.Text, ".") Then e.Handled = True

                    Case ChrW(Keys.Back)
                        '/ Return False is Default value

                    Case Else
                        e.Handled = True

                End Select
        End Select
    End Sub

    Private Sub frmDetailDataGridView_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
