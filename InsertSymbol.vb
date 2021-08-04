Friend Class frmInsertSymbol

    '   public components

    '      properties
    Friend Property SymbolFont As Font        'INPUT/OUTPUT: Font of symbol
    Friend Property SymbolCharacter As String 'OUTPUT: Character for symbol desired

    '      constructor

    Friend Sub New(ByVal SymbolFont As Font)
        '   set parameter
        Me.SymbolFont = SymbolFont : Me.SymbolCharacter = " "
        '   this call is required by the designer.
        InitializeComponent()
    End Sub



    '   private components

    '      private constants

    Private Const CharactersAtATime As Integer = 1024, _
     TotalCharacters As Integer = 65536, CharactersPerRow As Integer = 32


    '      private variables

    Private pvtPopulating As Boolean = False, pvtValidChar As Boolean = False 'status flags
    Private pvtSymbolList() As Char 'printable characters

    '      event procedures

    Private Sub InsertSymbol_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        '    get list of printable characters
        pvtSymbolList = GetPrintableCharacters()
        With cbxSymbolRange.Items
            .Clear()
            For StartCode As Integer = 0 To pvtSymbolList.Length - 1 _
              Step CharactersAtATime
                '    get character code range for this page
                Dim EndCode As Integer = _
                 Math.Min(pvtSymbolList.Length, StartCode + CharactersAtATime) - 1
                .Add(AscW(pvtSymbolList(StartCode)).ToString _
                   & "-" & AscW(pvtSymbolList(EndCode)).ToString)
            Next StartCode
        End With
        '   fill grid with characters and set font
        cbxSymbolRange.SelectedIndex = 0
        ComboBox1.Items.Clear()
        Dim FontSinifi As FontFamily
        For Each FontSinifi In FontFamily.Families
            ComboBox1.Items.Add(FontSinifi.Name)
        Next
        ComboBox1.Text = Me.SymbolFont.Name
        NewFont(Me.SymbolFont)

    End Sub

    Private Sub dgvSymbols_CurrentCellhanged(ByVal sender As Object, ByVal e As EventArgs) _
     Handles dgvSymbols.CurrentCellChanged
        '   cell clicked
        With dgvSymbols
            If Not (pvtPopulating OrElse .CurrentCell.Value Is Nothing) Then
                Dim character As String = .CurrentCell.Value.ToString
                pvtValidChar = Not String.IsNullOrEmpty(character)
                If pvtValidChar Then
                    '   character chosen
                    Me.SymbolCharacter = character : IndicateSymbol(Me.SymbolCharacter)
                End If
            End If
        End With
    End Sub

    Private Sub dgvSymbols_CellDoubleClick(ByVal sender As Object, _
     ByVal e As DataGridViewCellEventArgs) Handles dgvSymbols.CellDoubleClick
        '   cell double-clicked
        ' If pvtValidChar Then
        'Me.DialogResult = DialogResult.OK 'accept character
        'End If
        With dgvSymbols
            If Not (pvtPopulating OrElse .CurrentCell.Value Is Nothing) Then
                Me.DialogResult = DialogResult.OK 'accept character
            End If
        End With

    End Sub

    Private Sub btnInsert_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnInsert.Click
        With dgvSymbols
            If Not (pvtPopulating OrElse .CurrentCell.Value Is Nothing) Then
                Me.DialogResult = DialogResult.OK 'accept character
            End If
        End With

    End Sub
    Private Sub btnCnacel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCnacel.Click
        Me.DialogResult = DialogResult.Cancel 'ignore selected symbol
    End Sub
    Private Sub cbxSymbolRange_SelectedItemChanged(ByVal sender As Object, ByVal e As EventArgs) _
     Handles cbxSymbolRange.SelectedIndexChanged
        If cbxSymbolRange.SelectedIndex > -1 Then
            NewGridRange(cbxSymbolRange.SelectedIndex)
        End If
    End Sub
    '      NON-EVENT procedures
    '         change grid font and character range
    Private Sub NewGridRange(ByVal Page As Integer)
        '   change character range
        With dgvSymbols
            '   flag that we are updating the grid
            .SuspendLayout() : pvtPopulating = True
            '   get printable characters in this range
            '   fill grid
            Dim StartPos As Integer = Page * CharactersAtATime, _
             EndPos As Integer = Math.Min(pvtSymbolList.Length, StartPos + CharactersAtATime), _
             SymbolIndex As Integer = StartPos, _
             NumberOfRows As Integer = Math.Ceiling((EndPos - StartPos) / CharactersPerRow)
            Dim RowInfo(CharactersPerRow - 1) As String
            .ColumnCount = CharactersPerRow
            With .Rows
                .Clear()
                For row As Integer = 0 To NumberOfRows - 1
                    '   assign column values for row
                    For column As Integer = 0 To CharactersPerRow - 1
                        If SymbolIndex >= EndPos Then
                            '   past range of printable characters--use null string
                            RowInfo(column) = ""
                        Else
                            '   printable character--use it
                            RowInfo(column) = pvtSymbolList(SymbolIndex).ToString
                            SymbolIndex += 1
                        End If
                    Next column
                    '   add row
                    .Add(RowInfo)
                Next row
            End With
            '   flag that we are done updating grid
            pvtPopulating = False : .Invalidate() : .ResumeLayout()
            '   indicate that first character of grid is selected
            pvtValidChar = True
            Me.SymbolCharacter = pvtSymbolList(StartPos) : IndicateSymbol(Me.SymbolCharacter)
        End With
    End Sub
    '         configure grid for new font
    Private Sub NewFont(ByVal NewFont As Font)
        '   change font
        With dgvSymbols
            .SuspendLayout()
            '   set new font
            .Font = New Font(NewFont.Name, .DefaultCellStyle.Font.SizeInPoints, NewFont.Style)
            .ResumeLayout()
        End With
    End Sub

    '         display current symbol info
    Private Sub IndicateSymbol(ByVal NewSymbol As String)
        lblSymbol.Text = "Simge: """ & NewSymbol _
         & """ Kısayol Tuşları: Alt+" & AscW(NewSymbol(0)).ToString
    End Sub

    '         get string of printable characters
    Private Function GetPrintableCharacters() As Char()
        '   go through Unicode character list
        Dim SymbolList As String = ""
        For SymbolIndex As Integer = 0 To TotalCharacters - 1
            '   check this character
            Dim Character As Char = ChrW(SymbolIndex)
            Select Case Char.GetUnicodeCategory(Character)
                Case Globalization.UnicodeCategory.PrivateUse, _
                  Globalization.UnicodeCategory.Format, _
                  Globalization.UnicodeCategory.OtherNotAssigned
                    '   character unavailable
                    Continue For
            End Select
            Select Case True
                Case Character = " "c
                    '   space--printable
                    SymbolList &= Character
                Case Char.IsControl(Character), Char.IsSeparator(Character), _
                   Char.IsWhiteSpace(Character), Char.IsSurrogate(Character)
                    '   character not renderable
                    Continue For
                Case Else
                    '   printable character
                    SymbolList &= Character
            End Select
        Next SymbolIndex
        '   return with list of characters
        Return _
         SymbolList.ToCharArray
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim fnt As Font = New Font(ComboBox1.Text, dgvSymbols.DefaultCellStyle.Font.SizeInPoints, Me.SymbolFont.Style)
        NewFont(fnt)
        Me.SymbolFont = fnt
    End Sub

End Class