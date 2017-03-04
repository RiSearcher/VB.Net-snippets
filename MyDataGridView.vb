

Imports System.Text.RegularExpressions
Imports System.IO

Public Class MyDataGridView

    Inherits DataGridView

    '//////////////////////////////////////////////////////////////////////////////
    ' Check if the cell is empty
    ' Uses MS (column, row) notation !!!
    Public ReadOnly Property IsEmpty(i As Integer, j As Integer) As Boolean
        Get
            Return Me(i, j).Value Is Nothing OrElse Me(i, j).Value Is DBNull.Value OrElse String.IsNullOrWhiteSpace(Me(i, j).Value.ToString())
        End Get
    End Property
    Public ReadOnly Property IsNotEmpty(i As Integer, j As Integer) As Boolean
        Get
            Return Not (Me(i, j).Value Is Nothing OrElse Me(i, j).Value Is DBNull.Value OrElse String.IsNullOrWhiteSpace(Me(i, j).Value.ToString()))
        End Get
    End Property

    ' Process "Ctrl-?" pressed in DataGridView!!!!
    '//////////////////////////////////////////////////////////////////////////////
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

        If e.KeyCode = Keys.V And e.Control Then
            PasteData(Me)
        End If

        If e.KeyCode = Keys.C And e.Control Then
            Me.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        End If


        If e.KeyCode = Keys.X And e.Control Then
            Me.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            If Me.GetCellCount(DataGridViewElementStates.Selected) > 0 Then

                Try
                    ' Add the selection to the clipboard.
                    Clipboard.SetDataObject(Me.GetClipboardContent())
                    ' Clear selected cells
                    Dim cell As DataGridViewCell
                    For Each cell In Me.SelectedCells
                        cell.Value = ""
                    Next
                Catch ex As System.Runtime.InteropServices.ExternalException

                End Try
            End If
        End If


        If e.KeyCode = Keys.Delete Then
            Dim cell As DataGridViewCell
            For Each cell In Me.SelectedCells
                cell.Value = ""
            Next
        End If


        MyBase.OnKeyDown(e)
    End Sub

    ' Process "Ctrl-V" pressed in Cell Editing mode!!!!
    '//////////////////////////////////////////////////////////////////////////////
    Protected Overrides Sub OnEditingControlShowing(e As DataGridViewEditingControlShowingEventArgs)

        Dim EditingTxtBox As TextBox = CType(e.Control, TextBox)
        RemoveHandler EditingTxtBox.KeyDown, AddressOf MyDataGridView_EditingModeKeyDown
        AddHandler EditingTxtBox.KeyDown, AddressOf MyDataGridView_EditingModeKeyDown

        MyBase.OnEditingControlShowing(e)
    End Sub

    '//////////////////////////////////////////////////////////////////////////////
    ' Column select
    Protected Overrides Sub OnColumnHeaderMouseClick(e As DataGridViewCellMouseEventArgs)

        For i As Integer = 0 To Me.ColumnCount - 1
            Me.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next


        If e.Button = Windows.Forms.MouseButtons.Left Then
            If Me.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect Then
                Me.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect
                Me.Columns(e.ColumnIndex).Selected = True
            End If

        End If

        MyBase.OnColumnHeaderMouseClick(e)

    End Sub
    ' Row select
    Protected Overrides Sub OnRowHeaderMouseClick(e As DataGridViewCellMouseEventArgs)

        If e.Button = Windows.Forms.MouseButtons.Left Then
            If Me.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect Then
                Me.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
                Me.Rows(e.RowIndex).Selected = True
            End If
        End If


        MyBase.OnRowHeaderMouseClick(e)

    End Sub
    '//////////////////////////////////////////////////////////////////////////////




    '//////////////////////////////////////////////////////////////////////////////////////////////////////////
    '
    '  Private methods
    '
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub PasteData(ByRef dgv As DataGridView)
        Dim tmp() As String
        Dim row_split() As String
        Dim i, ii As Integer
        Dim c, cc, r As Integer
        Try

            Dim str As String = Clipboard.GetText()
            tmp = Regex.Split(str, vbCrLf)

            If Clipboard.ContainsData("Biff5") OrElse Clipboard.ContainsData("Biff8") OrElse Clipboard.ContainsData("Biff12") Then

                ' The data is Excel table
                ' Get the size of the selected area from clipboard 'Format129' field
                Dim iData As IDataObject = Windows.Forms.Clipboard.GetDataObject()
                ' Does not work without next line...???
                Dim allFormats As [String]() = iData.GetFormats()
                If iData.GetDataPresent("Format129") Then
                    Dim Stream As MemoryStream = iData.GetData("Format129")
                    Dim sr As StreamReader = New StreamReader(Stream)
                    Dim tstr As String = sr.ReadToEnd()

                    ' Match strings like: "Cut 2R x 2C" or "Copy 2R x 2C"
                    Dim m As Match = Regex.Match(tstr, "^(Copy|Cut) (\d+)R x (\d+)C")
                    If m.Success Then
                        Dim groups As GroupCollection = m.Groups
                        Dim nr As Integer = CInt(groups.Item("2").Value)
                        ' Correct the number of selected rows
                        If nr < tmp.Length Then
                            ReDim Preserve tmp(nr - 1)
                        End If
                    End If
                End If

            ElseIf Clipboard.ContainsData("ObjectLink") Then
                Dim Stream As MemoryStream = Windows.Forms.Clipboard.GetData("ObjectLink")
                Dim sr As StreamReader = New StreamReader(Stream)
                If Regex.IsMatch(sr.ReadToEnd(), "Word\.Document") Then
                    ' The data may come from Word table
                    ' Trim last 'vbCrLf'
                    str = Regex.Replace(str, "\r?\n$", "")
                    tmp = Regex.Split(str, vbCrLf)
                End If
            End If


            r = dgv.CurrentCellAddress.Y()
            c = dgv.CurrentCellAddress.X()
            If tmp.Length >= dgv.Rows.Count - r Then
                dgv.Rows.Add(tmp.Length - (dgv.Rows.Count - r) + 1)
            End If

            For i = 0 To tmp.Length - 1

                row_split = tmp(i).Split(vbTab)
                cc = c
                For ii = 0 To row_split.Length - 1
                    If cc > dgv.ColumnCount - 1 Then
                        ' Save and restore 'SelectionMode'
                        Dim tmp_selectionmode As DataGridViewSelectionMode = Me.SelectionMode
                        Me.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
                        dgv.ColumnCount += 1
                        dgv.Columns(cc).SortMode = DataGridViewColumnSortMode.NotSortable
                        Me.SelectionMode = tmp_selectionmode
                    End If

                    dgv(cc, r).Value = row_split(ii).TrimStart
                    cc = cc + 1
                Next
                r = r + 1
            Next


        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' Process "Ctrl-V" pressed in Cell Editing mode!!!!
    '//////////////////////////////////////////////////////////////////////////////
    Public Sub MyDataGridView_EditingModeKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        Dim dgv As DataGridView = sender.EditingControlDataGridView

        If (e.KeyCode = Keys.V AndAlso e.Control) Then
            If Regex.IsMatch(Clipboard.GetText(), "[\t\n]") Then
                ' Paste in table mode
                dgv.EndEdit()
                PasteData(dgv)
            End If
        End If
    End Sub


End Class
