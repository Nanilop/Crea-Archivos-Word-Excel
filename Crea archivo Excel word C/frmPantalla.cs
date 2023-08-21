using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace Crea_archivo_Excel_word_C
{
    public partial class frmPantalla : Form
    {
        public frmPantalla()
        {
            InitializeComponent();
            btnWord.Image = Image.FromFile("..\\..\\Resources\\word.png");
            btnExcel.Image = Image.FromFile("..\\..\\Resources\\excel.png");
        }

        private void frmPantalla_Activated(object sender, EventArgs e)
        {
            frmPantalla.ActiveForm.Icon = new Icon("..\\..\\Resources\\world.ico");
        }

        private void btnWord_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Word.Application m_Word = new Microsoft.Office.Interop.Word.Application();
            DateTime f = DateTime.Now;
            String ruta = Path.GetFullPath("..\\..\\Resources\\") + "HelloWorldWord" + f.Day.ToString() + f.Month.ToString() + f.Year.ToString() + "_" + f.Hour.ToString() + "-" +
                    f.Minute.ToString() + ".docx";
            //m_Word.Application.Visible = true;
            Document d_word = m_Word.Documents.Add();
            Paragraphs ps_word = d_word.Paragraphs;
            Paragraph p_word = ps_word.Add();

            p_word.Range.Text = "Hello World!";
            ps_word.Add();
            p_word.Range.Text = "Hello World!";
            p_word.Range.Bold = 1;
            ps_word.Add();
            p_word.Range.Text = "Hello World!";
            p_word.Range.Bold = 0;
            p_word.Range.Italic = 1;
            m_Word.ActiveDocument.SaveAs2(ruta);
            m_Word.Quit(false);
            MessageBox.Show("Archivo creado: "+ruta);
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application aplicacion;
            Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
            Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
            aplicacion = new Microsoft.Office.Interop.Excel.Application();
            libros_trabajo = aplicacion.Workbooks.Add();
            hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
            //Recorremos el DataGridView rellenando la hoja de trabajo
            hoja_trabajo.Cells[1, 1] = "Hello World!";
            hoja_trabajo.Cells[2, 1] = "Hello";
            hoja_trabajo.Cells[3, 1] = "World";
            hoja_trabajo.Cells[4, 1] = "!!!";
            DateTime f = DateTime.Now;
            String archivo = Path.GetFullPath("..\\..\\Resources\\") + "HelloWorldExcel" + f.Day.ToString() + f.Month.ToString() + f.Year.ToString() + "_" + f.Hour.ToString() + "-" +
                    f.Minute.ToString() + ".xlsx";
            libros_trabajo.SaveAs(archivo);
            libros_trabajo.Close(true);
            aplicacion.Quit();
            MessageBox.Show("Archivo creado: " + archivo);

        }
    }
}
