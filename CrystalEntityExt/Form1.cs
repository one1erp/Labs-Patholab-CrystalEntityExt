using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace CrystalEntityExt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ReportDocument crystalReport = new ReportDocument();
            crystalReport.Load(@"P:\ziv\Crystal\Plate1.rpt");

            crystalReportViewer1.ReportSource = crystalReport;

            crystalReportViewer1.Refresh();
        }
    }
}
