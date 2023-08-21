Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class frmPantalla
    Public Sub New()

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()
        btnWord.Image = Image.FromFile("..\\..\\Resources\\word.png")
        btnExcel.Image = Image.FromFile("..\\..\\Resources\\excel.png")
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles btnWord.Click

        Dim m_Word As Word.Application = New Word.Application()
        Dim f As DateTime = DateTime.Now
        Dim ruta As String = Path.GetFullPath("..\\..\\Resources\\") + "HelloWorldWord" + f.Day.ToString() + f.Month.ToString() + f.Year.ToString() + "_" + f.Hour.ToString() + "-" +
                    f.Minute.ToString() + ".docx"
        '_Word.Application.Visible = True
        Dim d_word As Word.Document = m_Word.Documents.Add()
        Dim ps_word As Word.Paragraphs = d_word.Paragraphs
        Dim p_word As Paragraph = ps_word.Add()

        p_word.Range.Text = "Hello World!"
        ps_word.Add()
        p_word.Range.Text = "Hello World!"
        p_word.Range.Bold = 1
        ps_word.Add()
        p_word.Range.Text = "Hello World!"
        p_word.Range.Bold = 0
        p_word.Range.Italic = 1
        m_Word.ActiveDocument.SaveAs2(ruta)
        m_Word.Quit(False)
        MessageBox.Show("Archivo creado: " + ruta)
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Dim aplicacion As Microsoft.Office.Interop.Excel.Application
        Dim libros_trabajo As Microsoft.Office.Interop.Excel.Workbook
        Dim hoja_trabajo As Microsoft.Office.Interop.Excel.Worksheet
        aplicacion = New Microsoft.Office.Interop.Excel.Application()
        libros_trabajo = aplicacion.Workbooks.Add()
        hoja_trabajo = libros_trabajo.Worksheets(1)
        ''As (Microsoft.Office.Interop.Excel.Worksheet)
        hoja_trabajo.Cells(1, 1) = "Hello World!"
        hoja_trabajo.Cells(2, 1) = "Hello"
        hoja_trabajo.Cells(3, 1) = "World"
        hoja_trabajo.Cells(4, 1) = "!!!"
        Dim f As DateTime = DateTime.Now
        Dim archivo As String = Path.GetFullPath("..\\..\\Resources\\") + "HelloWorldExcel" + f.Day.ToString() + f.Month.ToString() + f.Year.ToString() + "_" + f.Hour.ToString() + "-" +
                    f.Minute.ToString() + ".xlsx"
        libros_trabajo.SaveAs(archivo)
        libros_trabajo.Close(True)
        aplicacion.Quit()
        MessageBox.Show("Archivo creado: " + archivo)

    End Sub


    Private Sub frmPantalla_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        frmPantalla.ActiveForm.Icon = New Icon("..\\..\\Resources\\world.ico")
    End Sub
End Class
