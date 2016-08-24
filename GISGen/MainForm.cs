using System;
using System.Windows.Forms;
using GISGen.Excel;

namespace GISGen
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            var drso = new DRSO("test.xlsx");
            Console.WriteLine(drso.Document.Worksheets[2].Cells["C3"].Text);
            Console.WriteLine(drso.Document.Worksheets[2].Cells["B3"].Text);
            drso.Close();
            
        }


    }
}
