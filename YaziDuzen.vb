Option Explicit On
Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Drawing.Printing
Imports System.IO
Imports System.IO.File
Public Class YaziDuzen
    Private checkPrint As Integer
    Private sonkarakter As Integer
    Private pf As Font
    Private sayisi As Integer
    Private sirasi As Integer
    Private sayfasay As Integer
    Private sayfatop As Integer
    Private uzunlugu As Integer
    Private harf As String
    Private kelime As String
    Private satir As String
    Private sr As String
    Private logo As Image
    Private baslik1 As String
    Private baslik2 As String
    Private baslik3 As String
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Yazikutusu1.SelectionAlignment = yazikutusu.yazikutusu.TextAlign.Left
        tusayarla()
    End Sub
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Yazikutusu1.SelectionAlignment = yazikutusu.yazikutusu.TextAlign.Center
        tusayarla()
    End Sub
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Yazikutusu1.SelectionAlignment = yazikutusu.yazikutusu.TextAlign.Right
        tusayarla()
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Yazikutusu1.SelectionAlignment = yazikutusu.yazikutusu.TextAlign.Justify
        tusayarla()
    End Sub
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        GroupBox1.Location = New Point(ToolStripButton5.Owner.Location.X + 200, ToolStripButton5.Owner.Location.Y + 23)
        GroupBox1.Visible = True
        TextBox5.Focus()
    End Sub
    Private Sub GroupBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles GroupBox1.Leave
        GroupBox1.Hide()
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim sbTaRtf As New System.Text.StringBuilder
        Dim saysut As Integer = Val(Trim(TextBox5.Text))
        Dim saysat As Integer = Val(Trim(TextBox4.Text))
        Dim hucsat As String = ""
        Dim huckal As String = ""
        Dim oran As Integer = Val(Trim(TextBox6.Text))
        Dim gen As Integer = oran * 565
        Dim hizala As String = "trql"
        If ComboBox1.SelectedIndex = 0 Then
            hizala = "trql"
        End If
        If ComboBox1.SelectedIndex = 1 Then
            hizala = "trqc"
        End If
        If ComboBox1.SelectedIndex = 2 Then
            hizala = "trqr"
        End If
        sbTaRtf.Append("{\rtf1")
        If CheckBox1.Checked = True Then
            sbTaRtf.Append("\trowd\" & hizala & "\trgaph70\trleft0\trbrdrl\brdrs\brdrw10 \trbrdrt\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trpaddl70\trpaddr70\trpaddfl3\trpaddfr3")
            For i = 1 To saysut
                hucsat = hucsat & "\clbrdrl\brdrw10\brdrs\clbrdrt\brdrw10\brdrs\clbrdrr\brdrw10\brdrs\clbrdrb\brdrw10\brdrs \cellx" & gen * i 'bütün sütunlar aynı genişlikte
                huckal = huckal & "\cell\pard\intbl"
            Next
            sbTaRtf.Append(hucsat)
            For i = 1 To saysat
                sbTaRtf.Append(huckal & "\row")
            Next
        Else
            sbTaRtf.Append("\trowd\" & hizala & "\trgaph70\trleft0")
            For i = 1 To saysut
                hucsat = hucsat & "\cellx" & gen * i 'bütün sütunlar aynı genişlikte
                huckal = huckal & "\cell\pard\intbl"
            Next
            sbTaRtf.Append(hucsat)
            For i = 1 To saysat
                sbTaRtf.Append(huckal & "\row")
            Next
        End If
        sbTaRtf.Append("}")
        Clipboard.SetText(sbTaRtf.ToString, TextDataFormat.Rtf)
        Yazikutusu1.Paste()
        GroupBox1.Hide()
        Yazikutusu1.Focus()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        GroupBox1.Hide()
    End Sub
    Private Sub Yazikutusu1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Yazikutusu1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        If files.Count = 0 Then
        Else
            Dim uzanti As String = Strings.Right(files(0), 4)
            uzanti = LCase(uzanti)
            If uzanti = ".exe" Or uzanti = ".dll" Or uzanti = ".com" Or uzanti = ".ocx" Then
                MsgBox("Bu dosya açılamaz sakıncalıdır.")
                GoTo sonnn
            End If
            If uzanti = ".rtf" Then
                Yazikutusu1.LoadFile(files(0), RichTextBoxStreamType.RichText)
            Else
                Yazikutusu1.LoadFile(files(0), RichTextBoxStreamType.PlainText)
            End If
            Me.Text = files(0)
            ToolStripStatusLabel21.Text = ""
            ToolStripStatusLabel20.Text = files(0)
        End If
sonnn:
        tusayarla()
    End Sub

    Private Sub Yazikutusu1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Yazikutusu1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub Yazikutusu1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Yazikutusu1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.P Then
            sayfasayisibul(PrintDocument2)
            PrintDialog1.ShowDialog()
        End If
        tusayarla()
    End Sub

    Private Sub Yazikutusu1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Yazikutusu1.KeyPress
        ToolStripStatusLabel10.Text = Yazikutusu1.TextLength
        ToolStripStatusLabel3.Text = Yazikutusu1.GetLineFromCharIndex(Yazikutusu1.SelectionStart) + 1
        ToolStripStatusLabel7.Text = Yazikutusu1.SelectionStart - Yazikutusu1.GetFirstCharIndexOfCurrentLine + 1
        ToolStripStatusLabel14.Text = Yazikutusu1.SelectionStart + 1
        If IsNothing(Yazikutusu1.SelectionFont) Then
            ToolStripComboBox1.Text = ""
            ToolStripComboBox2.Text = ""
        Else
            ToolStripComboBox1.Text = Yazikutusu1.SelectionFont.Name
            ToolStripComboBox2.Text = Yazikutusu1.SelectionFont.Size
        End If
    End Sub
    Private Sub Yazikutusu1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Yazikutusu1.KeyUp
        tusayarla()
    End Sub
    Private Sub Yazikutusu1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Yazikutusu1.MouseClick
        tusayarla()
    End Sub
    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        If PrintDialog1.PrinterSettings.PrintRange = 1 Then
            checkPrint = Yazikutusu1.SelectionStart
            sonkarakter = Yazikutusu1.SelectionStart + Yazikutusu1.SelectionLength
        Else
            checkPrint = 0
            sonkarakter = Yazikutusu1.TextLength
        End If
        sayfasay = 1
        logo = PictureBox1.Image
        baslik1 = TextBox1.Text
        baslik2 = TextBox2.Text
        baslik3 = TextBox3.Text
    End Sub
    Private Sub PrintDocument1_EndPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.EndPrint
        Yazikutusu1.HideSelection = False
    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Print the content of the RichTextBox. Store the last character printed.
        Dim yPos As Single = 0 'y position of the next line to print 
        Dim xPos As Single = 0 'x position of the next token to print 
        pf = New Font("Calibri", 12)
        If IsNothing(PictureBox1.Image) Then
        Else
            e.Graphics.DrawImage(logo, e.MarginBounds.Left, 10, 85, 85)
        End If
        Dim ww As SizeF = e.Graphics.MeasureString(baslik1, pf)
        yPos = 30
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(baslik1, pf, Brushes.Black, New Point(xPos, yPos))
        ww = e.Graphics.MeasureString(baslik2, pf)
        yPos = yPos + pf.GetHeight(e.Graphics)
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(baslik2, pf, Brushes.Black, New Point(xPos, yPos))
        ww = e.Graphics.MeasureString(baslik3, pf)
        yPos = yPos + pf.GetHeight(e.Graphics)
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(baslik3, pf, Brushes.Black, New Point(xPos, yPos))
        checkPrint = Yazikutusu1.Print(checkPrint, sonkarakter, e)
        ' Look for more pages
        yPos = e.MarginBounds.Bottom + 20
        ww = e.Graphics.MeasureString(sayfasay, pf)
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(sayfasay & "/" & sayfatop, pf, Brushes.Black, New Point(xPos, yPos))
        If checkPrint < sonkarakter Then
            e.HasMorePages = True
            sayfasay = sayfasay + 1
        Else
            e.HasMorePages = False
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        On Error GoTo bit
        Dim opf As New OpenFileDialog
        opf.Filter = "Resim Dosyaları|*.png; *.bmp; *.jpg; *.jpeg; *.gif|Tüm Dosyalar|*.*"
        If opf.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim img As Image = Image.FromFile(opf.FileName)
            If img.Size.Height > 5000 Then
                MsgBox("Resim boyutu çok büyük")
                Exit Sub
            End If
            If img.Size.Width > 5000 Then
                MsgBox("Resim boyutu çok büyük")
                Exit Sub
            End If
            PictureBox1.Image = img
            PictureBox1.ImageLocation = opf.FileName
        End If
        Exit Sub
bit:
        MsgBox("Resim dosyası geçersiz veya çok büyük", vbOKOnly, "Logo Seç")
    End Sub
    Private Sub GroupBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles GroupBox2.Leave
        GroupBox2.Visible = False
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        GroupBox2.Visible = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        PictureBox1.Image = My.Resources.Yeni_Logo_Küçük
    End Sub
    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        GroupBox2.Location = New Point(ToolStripButton6.Owner.Location.X + 265, ToolStripButton6.Owner.Location.Y + 23)
        GroupBox2.Visible = True
        TextBox1.Focus()
    End Sub
    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        sayfasayisibul(PrintDocument2)
        PrintPreviewDialog1.ShowDialog()
    End Sub
    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        PageSetupDialog1.ShowDialog()
        PageSetupDialog2.PageSettings.Margins.Left = PageSetupDialog1.PageSettings.Margins.Left
        PageSetupDialog2.PageSettings.Margins.Top = PageSetupDialog1.PageSettings.Margins.Top
        PageSetupDialog2.PageSettings.Margins.Right = PageSetupDialog1.PageSettings.Margins.Right
        PageSetupDialog2.PageSettings.Margins.Bottom = PageSetupDialog1.PageSettings.Margins.Bottom
        PageSetupDialog2.PageSettings.PaperSize = PageSetupDialog1.PageSettings.PaperSize
        PageSetupDialog2.PageSettings.PaperSource = PageSetupDialog1.PageSettings.PaperSource
        PageSetupDialog2.PageSettings.Landscape = PageSetupDialog1.PageSettings.Landscape
        zoomayarla()
    End Sub
    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        On Error GoTo bit
        Dim opf As New OpenFileDialog
        opf.Title = "Resim dosyası seç"
        opf.Filter = "Resim Dosyaları|*.png; *.bmp; *.jpg; *.jpeg; *.gif|Tüm Dosyalar|*.*"
        If opf.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim img As Image = Image.FromFile(opf.FileName)
            If img.Size.Height > 2500 Then
                MsgBox("Resim boyutu çok büyük")
                Exit Sub
            End If
            If img.Size.Width > 2500 Then
                MsgBox("Resim boyutu çok büyük")
                Exit Sub
            End If
            Clipboard.SetImage(img)
            Yazikutusu1.Paste()
            Clipboard.Clear()
        End If
        Exit Sub
bit:
        MsgBox("Resim dosyası geçersiz veya çok büyük", vbOKOnly, "Resim Ekle")
    End Sub

    Private Sub ToolStripComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ToolStripComboBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Yazikutusu1.Focus()
        End If
    End Sub
    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim ilkfont As Font
        Dim yenifont As Font
        Dim uzunluk As Integer
        If IsNothing(Yazikutusu1.SelectionFont) Then
            uzunluk = Yazikutusu1.SelectionLength
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = 1
            ilkfont = Yazikutusu1.SelectionFont
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = uzunluk
        Else
            ilkfont = New Font(Yazikutusu1.SelectionFont.Name, Yazikutusu1.SelectionFont.Size, Yazikutusu1.SelectionFont.Style)
        End If
        yenifont = New Font(ToolStripComboBox1.Text, ilkfont.Size, ilkfont.Style)
        Yazikutusu1.SelectionFont = yenifont
    End Sub

    Private Sub ToolStripComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ToolStripComboBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Dim ilkfont As Font
            Dim yenifont As Font
            Dim uzunluk As Integer
            If IsNothing(Yazikutusu1.SelectionFont) Then
                uzunluk = Yazikutusu1.SelectionLength
                Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
                Yazikutusu1.SelectionLength = 1
                ilkfont = Yazikutusu1.SelectionFont
                Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
                Yazikutusu1.SelectionLength = uzunluk
            Else
                ilkfont = New Font(Yazikutusu1.SelectionFont.Name, Yazikutusu1.SelectionFont.Size, Yazikutusu1.SelectionFont.Style)
            End If
            Dim stil As FontStyle = ilkfont.Style
            Dim szz As Single
            szz = ToolStripComboBox2.Text
            yenifont = New Font(ilkfont.Name, szz, stil)
            Yazikutusu1.SelectionFont = yenifont
            Yazikutusu1.Focus()
        End If
    End Sub
    Private Sub ToolStripComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        Dim ilkfont As Font
        Dim yenifont As Font
        Dim uzunluk As Integer
        If IsNothing(Yazikutusu1.SelectionFont) Then
            uzunluk = Yazikutusu1.SelectionLength
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = 1
            ilkfont = Yazikutusu1.SelectionFont
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = uzunluk
        Else
            ilkfont = New Font(Yazikutusu1.SelectionFont.Name, Yazikutusu1.SelectionFont.Size, Yazikutusu1.SelectionFont.Style)
        End If
        Dim stil As FontStyle = ilkfont.Style
        Dim szz As Single
        szz = ToolStripComboBox2.Text
        yenifont = New Font(ilkfont.Name, szz, stil)
        Yazikutusu1.SelectionFont = yenifont
    End Sub
    Private Sub ToolStripButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton10.Click
        Dim fnt As Font = New Font(Yazikutusu1.SelectionFont.Name, Yazikutusu1.SelectionFont.Size)
        Dim InsertSymbolForm As frmInsertSymbol = New frmInsertSymbol(fnt)
        With InsertSymbolForm
            If .ShowDialog() = DialogResult.OK Then
                Yazikutusu1.SelectedText = .SymbolCharacter
                Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart - 1
                Yazikutusu1.SelectionLength = 1
                fnt = New Font(InsertSymbolForm.Font.Name, InsertSymbolForm.Font.Size)
                Yazikutusu1.SelectionFont = fnt
                Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart + 1
                Yazikutusu1.SelectionLength = 0
            End If
        End With
    End Sub
    Private Sub ToolStripMenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click
        Yazikutusu1.ZoomFactor = 1
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click
        Yazikutusu1.ZoomFactor = 0.5
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        Yazikutusu1.ZoomFactor = 0.25
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        Yazikutusu1.ZoomFactor = 1.5
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click
        Yazikutusu1.ZoomFactor = 2
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click
        Yazikutusu1.ZoomFactor = 0.75
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        Yazikutusu1.ZoomFactor = 1.25
        zoomayarla()
    End Sub
    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        Yazikutusu1.ZoomFactor = 4
        zoomayarla()
    End Sub
    Private Sub ToolStripButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton11.Click
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        If Yazikutusu1.SelectionType = 2 Then Exit Sub
        Dim renk As New ColorDialog
        renk.FullOpen = True
        If renk.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim rengi As Color = renk.Color
            Yazikutusu1.SelectionColor = rengi
            ToolStripButton11.ForeColor = rengi
        End If
    End Sub
    Private Sub kelimesec()
        Dim basla As Integer = Yazikutusu1.SelectionStart
        Dim kelimebasla As Integer
        Dim kelimebit As Integer
        Dim onceki As Integer
        Dim ek As Integer = 0
        Do
            If Yazikutusu1.SelectionStart = 0 Then onceki = 0 : Exit Do
            onceki = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart - 1
            If onceki - Yazikutusu1.SelectionStart <> 1 Then
                kelimebasla = onceki
                Exit Do
            Else
            End If
            Yazikutusu1.SelectionLength = 1
            If Yazikutusu1.SelectionLength <> 1 Then
                Exit Do
            End If
            If Asc(Yazikutusu1.SelectedText) = 13 Then
                Exit Do
            End If
            If Asc(Yazikutusu1.SelectedText) = 32 Then
                Exit Do
            End If
            If Asc(Yazikutusu1.SelectedText) = 10 Then
                Exit Do
            End If
            'Dim koddd As Integer = Asc(Yazikutusu1.SelectedText)

        Loop
        Yazikutusu1.Select(onceki, 1)
        If basla > onceki Then
            kelimebasla = onceki
        Else
            kelimebasla = basla
        End If
        Yazikutusu1.Select(basla, 1)
        If Yazikutusu1.SelectionLength > 1 Then
            kelimebit = basla
            GoTo sonrasiyok
        End If
        onceki = basla
        Do
            If Yazikutusu1.SelectionStart = Yazikutusu1.TextLength - 1 Then ek = 1 : onceki = Yazikutusu1.SelectionStart : Exit Do
            Yazikutusu1.SelectionLength = 1
            If onceki - Yazikutusu1.SelectionStart > 1 Then
                ek = 0
                Exit Do
            End If
            onceki = Yazikutusu1.SelectionStart
            If Yazikutusu1.SelectionLength <> 1 Then
                Exit Do
            End If

            If Yazikutusu1.SelectionLength = 0 Then
                Exit Do
            End If
            If Asc(Yazikutusu1.SelectedText) = 13 Then
                Exit Do
            End If
            If Asc(Yazikutusu1.SelectedText) = 32 Then
                Exit Do
            End If
            If Asc(Yazikutusu1.SelectedText) = 10 Then
                Exit Do
            End If
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart + 1
            If onceki = Yazikutusu1.SelectionStart Then
                Exit Do
            End If
        Loop
        If kelimebasla >= Yazikutusu1.SelectionStart Then
            kelimebit = onceki
            ek = 1
        Else
            If onceki = Yazikutusu1.SelectionStart Then
                kelimebit = onceki
            Else
                kelimebit = Yazikutusu1.SelectionStart - 1
            End If
        End If
sonrasiyok:
        Yazikutusu1.Select(kelimebasla, kelimebit - kelimebasla + ek)
    End Sub
    Private Sub ToolStripButton12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton12.Click
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        If Yazikutusu1.SelectionType = 2 Then Exit Sub
        Dim renk As New ColorDialog
        renk.FullOpen = True
        If renk.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim rengi As Color = renk.Color
            Yazikutusu1.SelectionBackColor = rengi
            ToolStripButton12.ForeColor = rengi
        End If
    End Sub
    Private Sub Yazikutusu1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Yazikutusu1.MouseUp
        tusayarla()
    End Sub
    Private Sub tusayarla()
        ToolStripStatusLabel10.Text = Yazikutusu1.TextLength
        ToolStripStatusLabel3.Text = Yazikutusu1.GetLineFromCharIndex(Yazikutusu1.SelectionStart) + 1
        ToolStripStatusLabel7.Text = Yazikutusu1.SelectionStart - Yazikutusu1.GetFirstCharIndexOfCurrentLine + 1
        ToolStripStatusLabel14.Text = Yazikutusu1.SelectionStart + 1
        If IsNothing(Yazikutusu1.SelectionFont) Then
            ToolStripComboBox1.Text = ""
            ToolStripComboBox2.Text = ""
            ToolStripButton26.Checked = False
            ToolStripButton27.Checked = False
            ToolStripButton28.Checked = False
            ToolStripButton29.Checked = False
        Else
            ToolStripComboBox1.Text = Yazikutusu1.SelectionFont.Name
            ToolStripComboBox2.Text = Yazikutusu1.SelectionFont.Size
            If Yazikutusu1.SelectionFont.Bold = True Then
                ToolStripButton26.Checked = True
            Else
                ToolStripButton26.Checked = False
            End If
            If Yazikutusu1.SelectionFont.Italic = True Then
                ToolStripButton27.Checked = True
            Else
                ToolStripButton27.Checked = False
            End If
            If Yazikutusu1.SelectionFont.Underline = True Then
                ToolStripButton28.Checked = True
            Else
                ToolStripButton28.Checked = False
            End If
            If Yazikutusu1.SelectionFont.Strikeout = True Then
                ToolStripButton29.Checked = True
            Else
                ToolStripButton29.Checked = False
            End If
        End If
        If Yazikutusu1.SelectionBackColor.Name = "White" Or Yazikutusu1.SelectionBackColor = Color.Transparent Or Yazikutusu1.SelectionBackColor.Name = "Window" Then
            ToolStripButton12.ForeColor = Color.Black
        Else
            ToolStripButton12.ForeColor = Yazikutusu1.SelectionBackColor
        End If
        Select Case Yazikutusu1.SelectionAlignment
            Case 1
                ToolStripButton1.Checked = True
                ToolStripButton2.Checked = False
                ToolStripButton3.Checked = False
                ToolStripButton4.Checked = False
            Case 2
                ToolStripButton1.Checked = False
                ToolStripButton2.Checked = False
                ToolStripButton3.Checked = True
                ToolStripButton4.Checked = False
            Case 3
                ToolStripButton1.Checked = False
                ToolStripButton2.Checked = True
                ToolStripButton3.Checked = False
                ToolStripButton4.Checked = False
            Case 4
                ToolStripButton1.Checked = False
                ToolStripButton2.Checked = False
                ToolStripButton3.Checked = False
                ToolStripButton4.Checked = True
        End Select
        ToolStripButton11.ForeColor = Yazikutusu1.SelectionColor
    End Sub
    Private Sub Yazikutusu1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Yazikutusu1.MouseWheel
        If (ModifierKeys And Keys.Control) <> 0 Then
            CType(e, HandledMouseEventArgs).Handled = True
            If e.Delta > 0 Then
                If Yazikutusu1.ZoomFactor >= 4 Then Exit Sub
                Yazikutusu1.ZoomFactor = Yazikutusu1.ZoomFactor + 0.1
                zoomayarla()
            Else
                If Yazikutusu1.ZoomFactor <= 0.25 Then Exit Sub
                Yazikutusu1.ZoomFactor = Yazikutusu1.ZoomFactor - 0.1
                zoomayarla()
            End If
        Else
        End If
    End Sub
    Private Sub ToolStripButton13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton13.Click
        Dim gecerli As Integer
        gecerli = Yazikutusu1.SelectionIndent
        If gecerli >= Yazikutusu1.Width - 50 Then
            Exit Sub
        End If
        Yazikutusu1.SelectionIndent = gecerli + 10
        Yazikutusu1.SelectionHangingIndent = 0 - gecerli - 10
    End Sub
    Private Sub ToolStripButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton14.Click
        Dim gecerli As Integer
        gecerli = Yazikutusu1.SelectionIndent
        If gecerli <= 10 Then
            Yazikutusu1.SelectionIndent = 0
            Yazikutusu1.SelectionHangingIndent = 0
            Exit Sub
        End If
        Yazikutusu1.SelectionIndent = gecerli - 10
        Yazikutusu1.SelectionHangingIndent = 0 - gecerli + 10
    End Sub
    Private Sub ToolStripButton16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton16.Click
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        Yazikutusu1.SelectionBackColor = Color.Transparent
        ToolStripButton12.ForeColor = Color.Black
    End Sub
    Private Sub ToolStripButton15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton15.Click
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        Yazikutusu1.SelectionColor = Color.Black
        ToolStripButton11.ForeColor = Color.Black
    End Sub
    Private Sub ToolStripButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton17.Click
        If Yazikutusu1.SelectionBullet = False Then
            Yazikutusu1.SelectionBullet = True
        Else
            Yazikutusu1.SelectionBullet = False
        End If
    End Sub
    Private Sub ToolStripButton18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton18.Click
        Dim ilkfont As Font
        Dim uzunluk As Integer = Yazikutusu1.SelectionLength
        Dim basla As Integer = Yazikutusu1.SelectionStart
        Dim carp As Double = 1
        If basla = 0 Then
            Yazikutusu1.SelectionStart = basla
        Else
            Yazikutusu1.SelectionStart = basla - 1
        End If
        Yazikutusu1.SelectionLength = 1
        If Yazikutusu1.SelectionCharOffset <> 0 Then
            carp = 1
        Else
            carp = 0.7
        End If
        ilkfont = Yazikutusu1.SelectionFont
        Yazikutusu1.SelectionStart = basla
        Yazikutusu1.SelectionLength = uzunluk
        Yazikutusu1.SelectionFont = New Font(ilkfont.Name, ilkfont.Size * carp)
        If carp = 1 Then
            Yazikutusu1.SelectionCharOffset = ilkfont.Size * 10 / 7 * 0.3
        Else
            Yazikutusu1.SelectionCharOffset = ilkfont.Size * 0.3
        End If
    End Sub
    Private Sub ToolStripButton19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton19.Click
        Dim ilkfont As Font
       Dim uzunluk As Integer = Yazikutusu1.SelectionLength
        Dim basla As Integer = Yazikutusu1.SelectionStart
        Dim carp As Double = 1
        If basla = 0 Then
            Yazikutusu1.SelectionStart = basla
        Else
            Yazikutusu1.SelectionStart = basla - 1
        End If
        Yazikutusu1.SelectionLength = 1
        If Yazikutusu1.SelectionCharOffset <> 0 Then
            carp = 1
        Else
            carp = 0.7
        End If
        ilkfont = Yazikutusu1.SelectionFont
        Yazikutusu1.SelectionStart = basla
        Yazikutusu1.SelectionLength = uzunluk
        Yazikutusu1.SelectionFont = New Font(ilkfont.Name, ilkfont.Size * carp)
        If carp = 1 Then
            Yazikutusu1.SelectionCharOffset = 0 - (ilkfont.Size * 10 / 7 * 0.3)
        Else
            Yazikutusu1.SelectionCharOffset = 0 - (ilkfont.Size * 0.3)
        End If
    End Sub
    Private Sub ToolStripButton20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton20.Click
        Dim ilkfont As Font
        Dim uzunluk As Integer = Yazikutusu1.SelectionLength
        Dim basla As Integer = Yazikutusu1.SelectionStart
        Dim carp As Double = 1
        If basla = 0 Then
            Yazikutusu1.SelectionStart = basla
        Else
            Yazikutusu1.SelectionStart = basla - 1
        End If
        Yazikutusu1.SelectionLength = 1
        If Yazikutusu1.SelectionCharOffset = 0 Then
            carp = 1
        Else
            carp = 0.7
        End If
        ilkfont = Yazikutusu1.SelectionFont
        Yazikutusu1.SelectionStart = basla
        Yazikutusu1.SelectionLength = uzunluk
        Yazikutusu1.SelectionFont = New Font(ilkfont.Name, ilkfont.Size / carp)
        Yazikutusu1.SelectionCharOffset = 0
    End Sub
    Private Sub TabloEkleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabloEkleToolStripMenuItem.Click
        GroupBox1.Location = New Point(MousePosition.X - PointToScreen(Yazikutusu1.Location).X + Yazikutusu1.Left, MousePosition.Y - PointToScreen(Yazikutusu1.Location).Y + Yazikutusu1.Top)
        GroupBox1.Visible = True
        TextBox5.Focus()
    End Sub
    Private Sub GroupBox1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GroupBox1.VisibleChanged
        If GroupBox1.Visible = True Then
            TextBox5.Focus()
            ComboBox1.SelectedIndex = 0
        End If
    End Sub
    Private Sub ResimEkleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResimEkleToolStripMenuItem.Click
        ToolStripButton9.PerformClick()
    End Sub
    Private Sub KopyalaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KopyalaToolStripMenuItem.Click
        Yazikutusu1.Copy()
    End Sub
    Private Sub YapıştırToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YapıştırToolStripMenuItem.Click
        Yazikutusu1.Paste()
    End Sub

    Private Sub KesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KesToolStripMenuItem.Click
        Yazikutusu1.Cut()
    End Sub
    Private Sub ToolStripButton21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton21.Click
        If ToolStripStatusLabel21.Text = "" Then
            GoTo yap
        End If
        Dim sonuc As MsgBoxResult
        sonuc = MsgBox("Dosya kaydedilmedi, kaydedilsin mi?", vbYesNoCancel + vbQuestion, "Yeni Dosya Oluştur")
        If sonuc = MsgBoxResult.Ok Then
            GoTo yap
        Else
            If sonuc = MsgBoxResult.Cancel Then
                Exit Sub
            End If
            If sonuc = MsgBoxResult.Yes Then
                ToolStripButton23.PerformClick()
                If ToolStripStatusLabel21.Text = "*" Then
                    Exit Sub
                Else
                    GoTo yap
                End If
            End If
        End If
        Exit Sub
yap:
        Yazikutusu1.Clear()
        ToolStripStatusLabel20.Text = "Yeni Dosya"
        Me.Text = "Yeni Dosya"
        ToolStripStatusLabel21.Text = ""
    End Sub

    Private Sub ToolStripButton22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton22.Click
        If ToolStripStatusLabel21.Text = "" Then
            GoTo yap
        End If
        Dim sonuc As MsgBoxResult
        sonuc = MsgBox("Yeni dosya açılacak. Dosya kaydedilmedi, kaydedilsin mi?", vbYesNoCancel + vbQuestion, "Yeni Dosya Aç")
        If sonuc = MsgBoxResult.No Then
            GoTo yap
        End If
        If sonuc = MsgBoxResult.Cancel Then
            Exit Sub
        End If
        If sonuc = MsgBoxResult.Yes Then
            ToolStripButton23.PerformClick()
            If ToolStripStatusLabel21.Text = "*" Then
                Exit Sub
            Else
                GoTo yap
            End If
        End If
        Exit Sub
yap:
        On Error GoTo bit
        Dim opf As New OpenFileDialog
        opf.Filter = "Zengin Metin Belgesi|*.rtf|Salt Metin Dosyası|*.txt|Tüm Dosyalar|*.*"
        If opf.ShowDialog = Windows.Forms.DialogResult.OK Then
            If opf.FilterIndex = 1 Then
                Yazikutusu1.LoadFile(opf.FileName, RichTextBoxStreamType.RichText)
            Else
                Yazikutusu1.LoadFile(opf.FileName, RichTextBoxStreamType.PlainText)
            End If
            ToolStripStatusLabel20.Text = opf.FileName
            Me.Text = opf.FileName
            ToolStripStatusLabel21.Text = ""
        End If
        Exit Sub
bit:
        MsgBox("Dosya geçersiz veya çok büyük", vbOKOnly, "Aç")
    End Sub
    Private Sub ToolStripButton23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton23.Click
        If ToolStripStatusLabel20.Text = "Yeni Dosya" Then
            Dim dsykay As New SaveFileDialog
            dsykay.Filter = "Zengin Metin Belgesi|*.rtf|Salt Metin Dosyası|*.txt|Tüm Dosyalar|*.*"
            If dsykay.ShowDialog = Windows.Forms.DialogResult.OK Then
                Yazikutusu1.SaveFile(dsykay.FileName)
                ToolStripStatusLabel20.Text = dsykay.FileName
                Me.Text = dsykay.FileName
                ToolStripStatusLabel21.Text = ""
                Exit Sub
            Else
                Exit Sub
            End If
        End If
        Dim uzan As String
        uzan = Mid(ToolStripStatusLabel20.Text, Len(ToolStripStatusLabel20.Text) - 2, 3)
        If uzan = "txt" Or uzan = "TXT" Then
            Yazikutusu1.SaveFile(ToolStripStatusLabel20.Text, RichTextBoxStreamType.PlainText)
        Else
            Yazikutusu1.SaveFile(ToolStripStatusLabel20.Text, RichTextBoxStreamType.RichText)
        End If
        ToolStripStatusLabel21.Text = ""
    End Sub
    Private Sub ToolStripButton24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton24.Click
        Dim dsykay As New SaveFileDialog
        dsykay.Filter = "Zengin Metin Belgesi|*.rtf|Salt Metin Dosyası|*.txt|Tüm Dosyalar|*.*"
        If dsykay.ShowDialog = Windows.Forms.DialogResult.OK Then
            If dsykay.FilterIndex = 2 Then
                Yazikutusu1.SaveFile(dsykay.FileName, RichTextBoxStreamType.PlainText)
            Else
                Yazikutusu1.SaveFile(dsykay.FileName, RichTextBoxStreamType.RichText)
            End If
            ToolStripStatusLabel20.Text = dsykay.FileName
            Me.Text = dsykay.FileName
            ToolStripStatusLabel21.Text = ""
        End If
    End Sub
    Private Sub Yazikutusu1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Yazikutusu1.TextChanged
        ToolStripStatusLabel21.Text = "*"
    End Sub
    Private Sub YaziDuzen_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        If files.Count = 0 Then
        Else
            Dim uzanti As String = Strings.Right(files(0), 4)
            uzanti = LCase(uzanti)
            If uzanti = ".exe" Or uzanti = ".dll" Or uzanti = ".com" Or uzanti = ".ocx" Then
                MsgBox("Bu dosya açılamaz sakıncalıdır.")
                GoTo sonnn
            End If
            If uzanti = ".rtf" Then
                Yazikutusu1.LoadFile(files(0), RichTextBoxStreamType.RichText)
            Else
                Yazikutusu1.LoadFile(files(0), RichTextBoxStreamType.PlainText)
            End If
            Me.Text = files(0)
            ToolStripStatusLabel21.Text = ""
            ToolStripStatusLabel20.Text = files(0)
        End If
sonnn:
        tusayarla()
    End Sub
    Private Sub YaziDuzen_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If ToolStripStatusLabel21.Text = "*" Then
            Dim sonuc As MsgBoxResult
            sonuc = MsgBox("Çıkış yapılacak. Dosya kaydedilmedi, kaydedilsin mi?", vbYesNoCancel + vbCritical, "Çıkış")
            If sonuc = MsgBoxResult.No Then
            Else
                If sonuc = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
                If sonuc = MsgBoxResult.Yes Then
                    ToolStripButton23.PerformClick()
                    If ToolStripStatusLabel21.Text = "*" Then
                        e.Cancel = True
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub YaziDuzen_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Yazikutusu1.Focus()
    End Sub
    Private Sub YaziDuzen_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.Width < Yazikutusu1.Width Then
            Yazikutusu1.Left = 0
            Yazikutusu1.Height = Me.Height - 110
            Me.AutoScroll = True
        Else
            Me.AutoScroll = False
            Yazikutusu1.Height = Me.Height - 90
            Yazikutusu1.Left = (Me.Width - Yazikutusu1.Width) / 2
        End If
    End Sub
    Private Sub zoomayarla()
        Dim genislik As Integer
        Dim msg(6) As Integer
        msg(0) = PageSetupDialog1.PageSettings.PaperSize.Height
        msg(1) = PageSetupDialog1.PageSettings.PaperSize.Width
        msg(2) = PageSetupDialog1.PageSettings.Margins.Top
        msg(3) = PageSetupDialog1.PageSettings.Margins.Bottom
        msg(4) = PageSetupDialog1.PageSettings.Margins.Left
        msg(5) = PageSetupDialog1.PageSettings.Margins.Right
        If PageSetupDialog1.PageSettings.Landscape = True Then
            genislik = Math.Round(((msg(0) - msg(2) - msg(3)) * 0.9474), 0)
        Else
            genislik = Math.Round(((msg(1) - msg(4) - msg(5)) * 0.9474), 0)
        End If
        genislik = (genislik * Yazikutusu1.ZoomFactor) + 17
        Yazikutusu1.Width = genislik
        If Me.Width < Yazikutusu1.Width Then
            Yazikutusu1.Left = 0
            Dim ekran As Rectangle = Screen.PrimaryScreen.WorkingArea
            If Yazikutusu1.Width > ekran.Width Then
                Me.Width = ekran.Width
            Else
                Me.Width = Yazikutusu1.Width + 17
            End If
        Else
            Yazikutusu1.Left = (Me.Width - Yazikutusu1.Width) / 2
        End If
        ToolStripStatusLabel18.Text = Yazikutusu1.ZoomFactor * 100
    End Sub
    Private Sub YaziDuzen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GroupBox1.Hide()
        GroupBox2.Hide()
        ToolStripComboBox1.Items.Clear()
        PictureBox1.Image = My.Resources.Yeni_Logo_Küçük
        TextBox1.Text = "Başlık 1"
        TextBox2.Text = "Başlık 2"
        TextBox3.Text = "Başlık 3"
        Dim FontSinifi As FontFamily
        For Each FontSinifi In FontFamily.Families
            ToolStripComboBox1.Items.Add(FontSinifi.Name)
        Next
        ToolStripStatusLabel20.Text = "Yeni Dosya"
        Me.Text = "Yeni Dosya"
        If Exists("logobilgi.txt") Then
            Dim fs As FileStream = New FileStream("logobilgi.txt", FileMode.Open)
            Dim Degisken As StreamReader = New StreamReader(fs, System.Text.Encoding.GetEncoding("UTF-8"), False)
            Dim satir1
            For i = 1 To 4
                satir1 = Degisken.ReadLine()
                If satir1 Is Nothing Then Exit For
                If i = 1 Then
                    If satir1 = "Yok" Then
                        PictureBox1.Image = Nothing
                    Else
                        PictureBox1.Image = Image.FromFile(satir1)
                    End If
                End If
                If i = 2 Then
                    TextBox1.Text = satir1
                End If
                If i = 3 Then
                    TextBox2.Text = satir1
                End If
                If i = 4 Then
                    TextBox3.Text = satir1
                End If
            Next
            fs.Close()
        Else
        End If
        Yazikutusu1.SelectAll()
        Yazikutusu1.SelectionBackColor = Color.Transparent
        ToolStripComboBox2.Items.Clear()
        ToolStripComboBox2.Items.Add("8")
        ToolStripComboBox2.Items.Add("9")
        ToolStripComboBox2.Items.Add("10")
        ToolStripComboBox2.Items.Add("11")
        ToolStripComboBox2.Items.Add("12")
        ToolStripComboBox2.Items.Add("14")
        ToolStripComboBox2.Items.Add("16")
        ToolStripComboBox2.Items.Add("18")
        ToolStripComboBox2.Items.Add("20")
        ToolStripComboBox2.Items.Add("25")
        ToolStripComboBox2.Items.Add("30")
        ToolStripComboBox2.Items.Add("35")
        ToolStripComboBox2.Items.Add("40")
        ToolStripComboBox2.Items.Add("50")
        If Me.Width < 611 Then
            Me.Width = 611
        End If
        PageSetupDialog1.PageSettings.Margins.Left = 70
        PageSetupDialog1.PageSettings.Margins.Top = 100
        PageSetupDialog1.PageSettings.Margins.Right = 30
        PageSetupDialog1.PageSettings.Margins.Bottom = 60
        PageSetupDialog2.PageSettings.Margins.Left = 70
        PageSetupDialog2.PageSettings.Margins.Top = 100
        PageSetupDialog2.PageSettings.Margins.Right = 30
        PageSetupDialog2.PageSettings.Margins.Bottom = 60
        zoomayarla()
        Yazikutusu1.Left = (Me.Width - Yazikutusu1.Width) / 2
        Yazikutusu1.Height = Me.Height - 90
        ToolStripStatusLabel10.Text = Yazikutusu1.TextLength
        ToolStripStatusLabel3.Text = Yazikutusu1.GetLineFromCharIndex(Yazikutusu1.SelectionStart) + 1
        ToolStripStatusLabel7.Text = Yazikutusu1.SelectionStart - Yazikutusu1.GetFirstCharIndexOfCurrentLine + 1
        ToolStripStatusLabel14.Text = Yazikutusu1.SelectionStart + 1
        If IsNothing(Yazikutusu1.SelectionFont) Then
            ToolStripComboBox1.Text = ""
            ToolStripComboBox2.Text = ""
        Else
            ToolStripComboBox1.Text = Yazikutusu1.SelectionFont.Name
            ToolStripComboBox2.Text = Yazikutusu1.SelectionFont.Size
        End If
        ToolStripStatusLabel21.Text = ""
        If My.Application.CommandLineArgs.Count = 0 Then
        Else
            If My.Computer.FileSystem.FileExists(My.Application.CommandLineArgs.Item(0)) Then
                Dim uzanti As String = Strings.Right(My.Application.CommandLineArgs.Item(0), 4)
                uzanti = LCase(uzanti)
                If uzanti = ".exe" Or uzanti = ".dll" Or uzanti = ".com" Or uzanti = ".ocx" Then
                    MsgBox("Bu dosya açılamaz sakıncalıdır.")
                    GoTo sonnn
                End If
                If uzanti = ".rtf" Then
                    Yazikutusu1.LoadFile(My.Application.CommandLineArgs.Item(0), RichTextBoxStreamType.RichText)
                Else
                    Yazikutusu1.LoadFile(My.Application.CommandLineArgs.Item(0), RichTextBoxStreamType.PlainText)
                End If
                Me.Text = My.Application.CommandLineArgs.Item(0)
                ToolStripStatusLabel21.Text = ""
                ToolStripStatusLabel20.Text = My.Application.CommandLineArgs.Item(0)
            Else
                Me.Text = My.Application.CommandLineArgs.Item(0)
                ToolStripStatusLabel21.Text = ""
                ToolStripStatusLabel20.Text = My.Application.CommandLineArgs.Item(0)
            End If
        End If
sonnn:
        tusayarla()
        Yazikutusu1.AllowDrop = True
    End Sub
    Private Sub PrintDocument2_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument2.BeginPrint
        If PrintDialog1.PrinterSettings.PrintRange = 1 Then
            checkPrint = Yazikutusu1.SelectionStart
            sonkarakter = Yazikutusu1.SelectionStart + Yazikutusu1.SelectionLength
        Else
            checkPrint = 0
            sonkarakter = Yazikutusu1.TextLength
        End If
        sayfasay = 1
        logo = PictureBox1.Image
        baslik1 = TextBox1.Text
        baslik2 = TextBox2.Text
        baslik3 = TextBox3.Text
    End Sub
    Private Sub PrintDocument2_EndPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument2.EndPrint
        sayfatop = sayfasay
        PrintPreviewDialog2.Hide()
    End Sub
    Private Sub PrintDocument2_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage
        ' Print the content of the RichTextBox. Store the last character printed.
        Dim yPos As Single = 0 'y position of the next line to print 
        Dim xPos As Single = 0 'x position of the next token to print 
        pf = New Font("Calibri", 12)
        If IsNothing(PictureBox1.Image) Then
        Else
            e.Graphics.DrawImage(logo, e.MarginBounds.Left, 10, 85, 85)
        End If
        Dim ww As SizeF = e.Graphics.MeasureString(baslik1, pf)
        yPos = 30
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(baslik1, pf, Brushes.Black, New Point(xPos, yPos))
        ww = e.Graphics.MeasureString(baslik2, pf)
        yPos = yPos + pf.GetHeight(e.Graphics)
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(baslik2, pf, Brushes.Black, New Point(xPos, yPos))
        ww = e.Graphics.MeasureString(baslik3, pf)
        yPos = yPos + pf.GetHeight(e.Graphics)
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(baslik3, pf, Brushes.Black, New Point(xPos, yPos))
        checkPrint = Yazikutusu1.Print(checkPrint, sonkarakter, e)
        ' Look for more pages
        yPos = e.MarginBounds.Bottom + 20
        ww = e.Graphics.MeasureString(sayfasay, pf)
        xPos = (e.MarginBounds.Right - ww.Width + e.MarginBounds.Left) / 2
        e.Graphics.DrawString(sayfasay, pf, Brushes.Black, New Point(xPos, yPos))
        If checkPrint < sonkarakter Then
            e.HasMorePages = True
            sayfasay = sayfasay + 1
        Else
            e.HasMorePages = False
        End If
    End Sub
    Private Sub ToolStripButton26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton26.Click
        Dim ilkfont As Font
        Dim yenifont As Font
        Dim uzunluk As Integer
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        If IsNothing(Yazikutusu1.SelectionFont) Then
            uzunluk = Yazikutusu1.SelectionLength
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = 1
            ilkfont = Yazikutusu1.SelectionFont
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = uzunluk
        Else
            ilkfont = Yazikutusu1.SelectionFont
        End If
        Dim stil As FontStyle
        Select Case ilkfont.Style
            Case 0
                stil = 1
            Case 1
                stil = 0
            Case 2
                stil = 3
            Case 3
                stil = 2
            Case 4
                stil = 5
            Case 5
                stil = 4
            Case 6
                stil = 7
            Case 7
                stil = 6
            Case 8
                stil = 9
            Case 9
                stil = 8
            Case 10
                stil = 11
            Case 11
                stil = 10
            Case 12
                stil = 13
            Case 13
                stil = 12
            Case 14
                stil = 15
            Case 15
                stil = 14
        End Select
        yenifont = New Font(ilkfont.Name, ilkfont.Size, stil)
        Yazikutusu1.SelectionFont = yenifont
    End Sub
    Private Sub ToolStripButton27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton27.Click
        Dim ilkfont As Font
        Dim yenifont As Font
        Dim uzunluk As Integer
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        If IsNothing(Yazikutusu1.SelectionFont) Then
            uzunluk = Yazikutusu1.SelectionLength
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = 1
            ilkfont = Yazikutusu1.SelectionFont
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = uzunluk
        Else
            ilkfont = Yazikutusu1.SelectionFont
        End If
        Dim stil As FontStyle
        Select Case ilkfont.Style
            Case 0
                stil = 2
            Case 1
                stil = 3
            Case 2
                stil = 0
            Case 3
                stil = 1
            Case 4
                stil = 6
            Case 5
                stil = 7
            Case 6
                stil = 4
            Case 7
                stil = 5
            Case 8
                stil = 10
            Case 9
                stil = 11
            Case 10
                stil = 8
            Case 11
                stil = 9
            Case 12
                stil = 14
            Case 13
                stil = 15
            Case 14
                stil = 12
            Case 15
                stil = 13
        End Select
        yenifont = New Font(ilkfont.Name, ilkfont.Size, stil)
        Yazikutusu1.SelectionFont = yenifont
    End Sub
    Private Sub ToolStripButton28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton28.Click
        Dim ilkfont As Font
        Dim yenifont As Font
        Dim uzunluk As Integer
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        If IsNothing(Yazikutusu1.SelectionFont) Then
            uzunluk = Yazikutusu1.SelectionLength
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = 1
            ilkfont = Yazikutusu1.SelectionFont
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = uzunluk
        Else
            ilkfont = Yazikutusu1.SelectionFont
        End If
        Dim stil As FontStyle
        Select Case ilkfont.Style
            Case 0
                stil = 4
            Case 1
                stil = 5
            Case 2
                stil = 6
            Case 3
                stil = 7
            Case 4
                stil = 0
            Case 5
                stil = 1
            Case 6
                stil = 2
            Case 7
                stil = 3
            Case 8
                stil = 12
            Case 9
                stil = 13
            Case 10
                stil = 14
            Case 11
                stil = 15
            Case 12
                stil = 8
            Case 13
                stil = 9
            Case 14
                stil = 10
            Case 15
                stil = 11
        End Select
        yenifont = New Font(ilkfont.Name, ilkfont.Size, stil)
        Yazikutusu1.SelectionFont = yenifont
    End Sub
    Private Sub ToolStripButton29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton29.Click
        Dim ilkfont As Font
        Dim yenifont As Font
        Dim uzunluk As Integer
        If Yazikutusu1.SelectionLength = 0 Then
            kelimesec()
        End If
        If IsNothing(Yazikutusu1.SelectionFont) Then
            uzunluk = Yazikutusu1.SelectionLength
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = 1
            ilkfont = Yazikutusu1.SelectionFont
            Yazikutusu1.SelectionStart = Yazikutusu1.SelectionStart
            Yazikutusu1.SelectionLength = uzunluk
        Else
            ilkfont = Yazikutusu1.SelectionFont
        End If
        Dim stil As FontStyle
        Select Case ilkfont.Style
            Case 0
                stil = 8
            Case 1
                stil = 9
            Case 2
                stil = 10
            Case 3
                stil = 11
            Case 4
                stil = 12
            Case 5
                stil = 13
            Case 6
                stil = 14
            Case 7
                stil = 15
            Case 8
                stil = 0
            Case 9
                stil = 1
            Case 10
                stil = 2
            Case 11
                stil = 3
            Case 12
                stil = 4
            Case 13
                stil = 5
            Case 14
                stil = 6
            Case 15
                stil = 7
        End Select
        yenifont = New Font(ilkfont.Name, ilkfont.Size, stil)
        Yazikutusu1.SelectionFont = yenifont
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        PictureBox1.Image = Nothing
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    Private Sub YaziDuzen_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Public Shared Function sayfasayisibul(ByVal printDocument As PrintDocument) As Integer
        printDocument.PrinterSettings.PrintFileName = Path.GetTempFileName()
        printDocument.PrinterSettings.PrintToFile = True
        Dim count As Integer = 0
        printDocument.PrintController = New StandardPrintController()
        printDocument.Print()
        File.Delete(printDocument.PrinterSettings.PrintFileName)
        Return count
    End Function
    Private Sub ToolStripButton25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton25.Click
        PrintDialog1.AllowSelection = True
        Dim sonuc As DialogResult = PrintDialog1.ShowDialog()
        If sonuc = vbOK Then
            PrintDocument1.PrinterSettings.PrinterName = PrintDialog1.PrinterSettings.PrinterName
            PrintPreviewDialog2.ShowDialog()
            PrintDocument1.Print()
        End If
        'Dim sdsa As Size = New Size(GetPreferredSize(Yazikutusu1.Size))
        'MsgBox(sdsa.Height & " " & Yazikutusu1.Height & "-" & Yazikutusu1.PreferredSize.Width & " " & Yazikutusu1.Width)
    End Sub
    Private Sub ToolStripButton30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton30.Click
        PrintPreviewDialog2.ShowDialog()
        'MsgBox("yazıldı")
        PrintDocument1.Print()
    End Sub
    Private Sub ToolStripButton30_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton30.MouseEnter
        ToolStripButton30.ToolTipText = "Hızlı yazdır (" & PrintDialog1.PrinterSettings.PrinterName & ")"
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Exists("logobilgi.txt") Then
            File.Delete("logobilgi.txt")
        Else
        End If
        Dim fs As FileStream = New FileStream("logobilgi.txt", FileMode.Create)
        fs.Close()
        Dim sw As StreamWriter = New StreamWriter("logobilgi.txt", True)
        If IsNothing(PictureBox1.ImageLocation) Then
            If IsNothing(PictureBox1.Image) Then
                sw.WriteLine("Yok")
            Else
                PictureBox1.Image.Save("logo.png")
                PictureBox1.ImageLocation = "logo.png"
                sw.WriteLine(PictureBox1.ImageLocation)
            End If
        End If
        sw.WriteLine(TextBox1.Text)
        sw.WriteLine(TextBox2.Text)
        sw.WriteLine(TextBox3.Text)
        sw.Close()
    End Sub
End Class
