using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;
using ClosedXML.Excel;

namespace calprice
{
    public partial class MainForm : Form
    {

        private DataTable   dtItems    = new DataTable();
        private decimal     ultimatasa;
        private bool        recalcular;

        public MainForm()
        {
            InitializeComponent();

            this.ultimatasa = 0;
            this.recalcular = false;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            ConfigForm();
        }

        private void MainForm_Activated(object sender, EventArgs e)
        {

        }

        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
            dataGridView1.Width     = (this.Width - 30);
            dataGridView1.Height    = (this.Height - 340);
            pnlTotales.Top          = (this.Height - 190);
            btnSeleccionarTodos.Top = pnlTotales.Top;
            pnlOpciones.Left        = (this.Width - pnlOpciones.Width - 16);
            pnlOpciones.Top         = (this.Height - 110);   
        }

        private void cboGrupo_Validating(object sender, CancelEventArgs e)
        {
            ActGrilla1();
        }

        private void btnMostrarTodos_Click(object sender, EventArgs e)
        {
            cboGrupo.SelectedValue = 0;
            texBuscapor.Text       = "";

            ActGrilla1();
        }

        private void btnMostrarSeleccionados_Click(object sender, EventArgs e)
        {
            ActGrilla1(true);
        }

        private void chkConStock_CheckedChanged(object sender, EventArgs e)
        {
            if (chkConStock.Checked)
            {
                chkSinStock.Checked = false;
            }
        }

        private void chkSinStock_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSinStock.Checked)
            {
                chkConStock.Checked = false;
            }
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            ActGrilla1();
        }

        private void texTasaCambio_Enter(object sender, EventArgs e)
        {
            texTasaCambio.Tag = texTasaCambio.Text;
        }

        private void texTasaCambio_Validating(object sender, CancelEventArgs e)
        {
            if (texTasaCambio.Text != texTasaCambio.Tag.ToString())
            {
                btnActualizarPrecios.Enabled = false;
            }               
        }

        private void btnRecalcularPreciosBs_Click(object sender, EventArgs e)
        {
            RecalcularPreciosBs();
        }

        private void btnRecalcularCostoDolar_Click(object sender, EventArgs e)
        {
            RecalcularCostoDolar();
        }

        private void chkRedondeaPrecioDolar_CheckedChanged(object sender, EventArgs e)
        {
            btnActualizarPrecios.Enabled = false;
        }

        private void btnListaPrecios_Click(object sender, EventArgs e)
        {
            ListaPrecios();
        }

        private void btnListaCostos_Click(object sender, EventArgs e)
        {
            ListaCostos();
        }

        private void btnActualizarPrecios_Click(object sender, EventArgs e)
        {
            ActualizarPrecios();
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSeleccionarTodos_Click(object sender, EventArgs e)
        {
            bool selecall = Convert.ToBoolean(btnSeleccionarTodos.Tag);

            selecall = !selecall;

            foreach (DataRow row in dtItems.Rows)
            {
                row["sj_selec"]  = selecall;
                row["actualiza"] = true;
            }

            btnSeleccionarTodos.Tag = selecall.ToString();

            CalTot();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 || (e.ColumnIndex >= 9 && e.ColumnIndex <= 12))
            {
                dataGridView1[17, e.RowIndex].Value = true;

                btnActualizarPrecios.Enabled        = true;

                CalTot();
            }
        }



        private void ConfigForm()
        {
            this.dtItems.Columns.Add("codprod",         typeof(string));
            this.dtItems.Columns.Add("descrip",         typeof(string));
            this.dtItems.Columns.Add("refere",          typeof(string));
            this.dtItems.Columns.Add("marca",           typeof(string));
            this.dtItems.Columns.Add("costact",         typeof(decimal));
            this.dtItems.Columns.Add("costant",         typeof(decimal));
            this.dtItems.Columns.Add("precio1",         typeof(decimal));
            this.dtItems.Columns.Add("precio2",         typeof(decimal));
            this.dtItems.Columns.Add("precio3",         typeof(decimal));
            this.dtItems.Columns.Add("existen",         typeof(decimal));
            this.dtItems.Columns.Add("fechauv",         typeof(DateTime));
            this.dtItems.Columns.Add("fechauc",         typeof(DateTime));
            this.dtItems.Columns.Add("sj_selec",        typeof(bool));
            this.dtItems.Columns.Add("sj_tasacambio",   typeof(decimal));
            this.dtItems.Columns.Add("sj_costodolar",   typeof(decimal));
            this.dtItems.Columns.Add("sj_p1dolar",      typeof(decimal));
            this.dtItems.Columns.Add("sj_p2dolar",      typeof(decimal));
            this.dtItems.Columns.Add("sj_p3dolar",      typeof(decimal));
            this.dtItems.Columns.Add("sj_putilidad1",   typeof(decimal));
            this.dtItems.Columns.Add("sj_putilidad2",   typeof(decimal));
            this.dtItems.Columns.Add("sj_putilidad3",   typeof(decimal));
            this.dtItems.Columns.Add("actualiza",       typeof(bool));

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource          = this.dtItems;

            misclases.namecompany     = ConfigurationManager.AppSettings["namecompany"].ToUpper();

            misclases.conexLite       = misclases.Conexion_Sqlite();
            misclases.conexSql        = misclases.Conexion_Sql();

            cboGrupo.DataSource       = misclases.CursorTable("select * from sainsta");

            cboBuscapor.SelectedIndex = 0;
            cboGrupo.SelectedValue    = 0;

            ActualizaEstructuraBD();

            //
            l_tasascambiodolar tasascambio = new l_tasascambiodolar();
            DataTable          dt1         = new DataTable();

            dt1 = tasascambio.select("1 order by id desc limit 1");
            
            if (dt1.Rows.Count > 0)
            {
                this.ultimatasa = Convert.ToDecimal(dt1.Rows[0]["tasacambio"]);
            }     

            texTasaCambio.Text = this.ultimatasa.ToString("F0");

            ActcboTasasCambio();

            ValidaInicio();
        }

        private void ActGrilla1(bool seleccionados = false)
        {
            t_saprod  saprod = new t_saprod();
            DataTable dt1    = new DataTable();
            string    condi  = "activo = 1";
            string[]  campos = { "codprod", "descrip", "refere", "marca" };

            if (cboGrupo.SelectedIndex >= 0)
            {
                condi = condi + " and codinst = " + cboGrupo.SelectedValue.ToString();
            }

            if (chkConStock.Checked)
            {
                condi = condi + " and existen > 0";
            }
            else if (chkSinStock.Checked)
            {
                condi = condi + " and existen <= 0";
            }

            if (chkSinCostoDolar.Checked)
            {
                condi = condi + " and sj_costodolar = 0";
            }

            if (texBuscapor.Text != "")
            {
                condi = condi + " and " + campos[cboBuscapor.SelectedIndex] + " like '%" + texBuscapor.Text.Trim() + "%'";
            }

            if (seleccionados)
            {
                condi = condi + " and sj_selec = 1";
            }

            this.dtItems.Rows.Clear();

            dt1 = saprod.select(condi);

            foreach (DataRow row in dt1.Rows)
            {
                this.dtItems.Rows.Add(new object[] {row["codprod"].ToString(),
                                                    row["descrip"].ToString(),
                                                    row["refere"].ToString(),
                                                    row["marca"].ToString(),
                                                    Convert.ToDecimal(row["costact"]),
                                                    Convert.ToDecimal(row["costant"]),
                                                    Convert.ToDecimal(row["precio1"]),
                                                    Convert.ToDecimal(row["precio2"]),
                                                    Convert.ToDecimal(row["precio3"]),
                                                    Convert.ToDecimal(row["existen"]),
                                                    DateTime.Now,
                                                    DateTime.Now,
                                                    Convert.ToBoolean(row["sj_selec"]),
                                                    Convert.ToDecimal(row["sj_tasacambio"]),
                                                    Convert.ToDecimal(row["sj_costodolar"]),
                                                    Convert.ToDecimal(row["sj_p1dolar"]),
                                                    Convert.ToDecimal(row["sj_p2dolar"]),
                                                    Convert.ToDecimal(row["sj_p3dolar"]),
                                                    Convert.ToDecimal(row["sj_putilidad1"]),
                                                    Convert.ToDecimal(row["sj_putilidad2"]),
                                                    Convert.ToDecimal(row["sj_putilidad3"]),
                                                    false});
            }

            btnRecalcularPreciosBs.Enabled   = (this.dtItems.Rows.Count > 0);
            btnRecalcularCostoDolar.Enabled  = btnRecalcularPreciosBs.Enabled;
            btnListaPrecios.Enabled          = btnRecalcularPreciosBs.Enabled;
            btnListaCostos.Enabled           = btnRecalcularPreciosBs.Enabled;
            btnActualizarPrecios.Enabled     = false;
            this.recalcular                  = false;

            CalTot();
        }

        private void CalTot()
        {
            decimal toproductos, tosincostodolar, toseleccionados, toporactualizar;

            toproductos     = this.dtItems.Rows.Count;
            tosincostodolar = this.dtItems.AsEnumerable().Where(r => r.Field<decimal>("sj_costodolar") == 0).Count();
            toseleccionados = this.dtItems.AsEnumerable().Where(r => r.Field<bool>("sj_selec")).Count();
            toporactualizar = this.dtItems.AsEnumerable().Where(r => r.Field<bool>("actualiza")).Count();

            texToproductos.Text     = toproductos.ToString("###,##0");
            texTosincostodolar.Text = tosincostodolar.ToString("###,##0");
            texToseleccionados.Text = toseleccionados.ToString("###,##0");
            texToporactualizar.Text = toporactualizar.ToString("###,##0");
        }

        private void RecalcularPreciosBs()
        {
            decimal tasacambio, costact, costant, precio1, precio2, precio3, sj_costodolar, 
                    sj_p1dolar, sj_p2dolar, sj_p3dolar, sj_putilidad1, sj_putilidad2, sj_putilidad3;

            tasacambio = Convert.ToDecimal(texTasaCambio.Text);

            if (tasacambio > 0)
            {
                if (tasacambio < this.ultimatasa)
                {
                    MessageBox.Show("Esta tasa de cambio es menor a la anterior, si actualiza los " +
                                    "precios de los productos bajaran", "Calprice",
                                     MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                DataRow[] rowFound = this.dtItems.Select("sj_selec");

                foreach (DataRow row in rowFound)
                {
                    costact       = Convert.ToDecimal(row["costact"]);
                    sj_costodolar = Convert.ToDecimal(row["sj_costodolar"]);
                    sj_putilidad1 = Convert.ToDecimal(row["sj_putilidad1"]);
                    sj_putilidad2 = Convert.ToDecimal(row["sj_putilidad2"]);
                    sj_putilidad3 = Convert.ToDecimal(row["sj_putilidad3"]);

                    if (sj_costodolar > 0)
                    {
                        costant              = costact;
                        costact              = (sj_costodolar * tasacambio);
                        precio1              = (costact / (100 - sj_putilidad1) * 100);
                        precio2              = (costact / (100 - sj_putilidad2) * 100);
                        precio3              = (costact / (100 - sj_putilidad3) * 100);
                        sj_p1dolar           = (precio1 / tasacambio);
                        sj_p2dolar           = (precio2 / tasacambio);
                        sj_p3dolar           = (precio3 / tasacambio);

                        row["costact"]       = costact;
                        row["costant"]       = costant;
                        row["precio1"]       = (precio1 > 0 ? precio1 : 0);
                        row["precio2"]       = (precio2 > 0 ? precio2 : 0);
                        row["precio3"]       = (precio3 > 0 ? precio3 : 0);
                        row["sj_tasacambio"] = tasacambio;
                        row["sj_p1dolar"]    = sj_p1dolar;
                        row["sj_p2dolar"]    = sj_p2dolar;
                        row["sj_p3dolar"]    = sj_p3dolar;

                        row["actualiza"]     = true;
                    }
                }

                btnActualizarPrecios.Enabled = true;
                this.recalcular              = true;

                CalTot();
            }
        }

        private void RecalcularCostoDolar()
        {
            decimal tasacambio, costact, precio1, precio2, precio3, sj_costodolar, 
                    sj_putilidad1, sj_putilidad2, sj_putilidad3;

            tasacambio = Convert.ToDecimal(texTasaCambio.Text);

            if (tasacambio > 0)
            {
                foreach (DataRow row in this.dtItems.Rows)
                {
                    costact       = Convert.ToDecimal(row["costact"]);
                    precio1       = Convert.ToDecimal(row["precio1"]);
                    precio2       = Convert.ToDecimal(row["precio2"]);
                    precio3       = Convert.ToDecimal(row["precio3"]);
                    sj_costodolar = Convert.ToDecimal(row["sj_costodolar"]);

                    if (sj_costodolar == 0)
                    {
                        sj_costodolar = (costact / tasacambio);

                        if (chkRedondeaPrecioDolar.Checked)
                        {
                            sj_costodolar    = Math.Round(sj_costodolar, 0);
                        }

                        row["sj_costodolar"] = sj_costodolar;
                    }

                    sj_putilidad1        = (precio1 > 0 ? ((precio1 - costact) / precio1 * 100) : 0);
                    sj_putilidad2        = (precio2 > 0 ? ((precio2 - costact) / precio2 * 100) : 0);
                    sj_putilidad3        = (precio3 > 0 ? ((precio3 - costact) / precio3 * 100) : 0);

                    sj_putilidad1        = Math.Round(sj_putilidad1, 0);
                    sj_putilidad2        = Math.Round(sj_putilidad2, 0);
                    sj_putilidad3        = Math.Round(sj_putilidad3, 0);

                    row["sj_putilidad1"] = (sj_putilidad1 > 0 ? sj_putilidad1 : 0);
                    row["sj_putilidad2"] = (sj_putilidad2 > 0 ? sj_putilidad2 : 0);
                    row["sj_putilidad3"] = (sj_putilidad3 > 0 ? sj_putilidad3 : 0);

                    row["actualiza"]     = true;
                }

                btnActualizarPrecios.Enabled = true;
                this.recalcular              = true;

                CalTot();
            }
        }

        private void ActualizarPrecios()
        {
            DialogResult SiNo;

            SiNo = MessageBox.Show("¿ Desea actualizar los precios de los productos ?",
                                   "Conforme", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                   MessageBoxDefaultButton.Button2);

            if (SiNo == DialogResult.Yes)
            {
                t_saprod           saprod      = new t_saprod();
                l_tasascambiodolar tasascambio = new l_tasascambiodolar();
                DataRow[]          rowFound    = this.dtItems.Select("actualiza");
                DateTime           fecha       = DateTime.Now;
                int                i           = 0;

                toolStripProgressBar1.Maximum = rowFound.Length;

                foreach (DataRow row in rowFound)
                {
                    toolStripStatusLabel1.Text = row["codprod"].ToString();

                    saprod.CostAct          = Convert.ToDecimal(row["costact"]);
                    saprod.CostPro          = saprod.CostAct;
                    saprod.CostAnt          = Convert.ToDecimal(row["costant"]);
                    saprod.Precio1          = Convert.ToDecimal(row["precio1"]);
                    saprod.Precio2          = Convert.ToDecimal(row["precio2"]);
                    saprod.Precio3          = Convert.ToDecimal(row["precio3"]);
                    saprod.sj_selec         = Convert.ToBoolean(row["sj_selec"]);
                    saprod.sj_tasacambio    = Convert.ToDecimal(row["sj_tasacambio"]);
                    saprod.sj_costodolar    = Convert.ToDecimal(row["sj_costodolar"]);
                    saprod.sj_p1dolar       = Convert.ToDecimal(row["sj_p1dolar"]);
                    saprod.sj_p2dolar       = Convert.ToDecimal(row["sj_p2dolar"]);
                    saprod.sj_p3dolar       = Convert.ToDecimal(row["sj_p3dolar"]);
                    saprod.sj_putilidad1    = Convert.ToDecimal(row["sj_putilidad1"]);
                    saprod.sj_putilidad2    = Convert.ToDecimal(row["sj_putilidad2"]);
                    saprod.sj_putilidad3    = Convert.ToDecimal(row["sj_putilidad3"]);
                    saprod.sj_feulactualiza = fecha;

                    saprod.update("codprod = '" + row["codprod"].ToString() + "'");

                    row["actualiza"]        = false;

                    toolStripProgressBar1.Value = i;
                    i++;
                }

                toolStripStatusLabel1.Text  = "Listo";
                toolStripProgressBar1.Value = 0;
                statusStrip1.Refresh();

                if (this.recalcular)
                {
                    tasascambio.tasacambio = Convert.ToDecimal(texTasaCambio.Text);
                    tasascambio.insert();

                    this.ultimatasa    = tasascambio.tasacambio;
                    this.recalcular    = false;

                    ActcboTasasCambio();
                }

                btnActualizarPrecios.Enabled = false;

                CalTot();
            }
        }

        private void ListaPrecios()
        {
            DataRow[] rowFound = dtItems.Select("existen > 0");
            string    reporte  = misclases.FileNameTemp("xlsx"), p;
            int       i        = 0, y = 6;

            toolStripStatusLabel1.Text = "Espere";
            statusStrip1.Refresh();

            toolStripProgressBar1.Maximum = rowFound.Length;

            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Hoja1");

            worksheet.PageSetup.PaperSize = XLPaperSize.LetterPaper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            worksheet.PageSetup.SetRowsToRepeatAtTop("$1:$5");

            worksheet.Cell(5, 01).Value = "Código";
            worksheet.Cell(5, 02).Value = "Producto";
            worksheet.Cell(5, 03).Value = "Referencia";
            worksheet.Cell(5, 04).Value = "Marca";
            worksheet.Cell(5, 05).Value = "Stock";
            worksheet.Cell(5, 06).Value = "Precio 1";
            worksheet.Cell(5, 07).Value = "Precio 1 dólar";
            worksheet.Cell(5, 08).Value = "Precio 2";
            worksheet.Cell(5, 09).Value = "Precio 2 dólar";
            worksheet.Cell(5, 10).Value = "Precio 3";
            worksheet.Cell(5, 11).Value = "Precio 3 dólar";

            worksheet.Range("A5:L5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A5:L5").Style.Font.Bold            = true;
            worksheet.Range("A5:L5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("A5:L5").Style.Border.RightBorder = XLBorderStyleValues.Thin;

            decimal tasacambio, precio1, precio2, precio3, sj_p1dolar, sj_p2dolar, sj_p3dolar;

            tasacambio = Convert.ToDecimal(texTasaCambio.Text);

            foreach (DataRow row in rowFound)
            {
                precio1    = Convert.ToDecimal(row["precio1"]);
                precio2    = Convert.ToDecimal(row["precio2"]);
                precio3    = Convert.ToDecimal(row["precio3"]);

                sj_p1dolar = sj_p2dolar = sj_p3dolar = 0;

                if (tasacambio > 0)
                {
                    sj_p1dolar = (precio1 / tasacambio);
                    sj_p2dolar = (precio2 / tasacambio);
                    sj_p3dolar = (precio3 / tasacambio);
                }

                worksheet.Cell(y, 01).Value = "'" + row["codprod"].ToString();
                worksheet.Cell(y, 02).Value = "'" + row["descrip"].ToString();
                worksheet.Cell(y, 03).Value = "'" + row["refere"].ToString();
                worksheet.Cell(y, 04).Value = "'" + row["marca"].ToString();
                worksheet.Cell(y, 05).Value = Convert.ToDecimal(row["existen"]);
                worksheet.Cell(y, 06).Value = precio1;
                worksheet.Cell(y, 07).Value = sj_p1dolar;
                worksheet.Cell(y, 08).Value = precio2;
                worksheet.Cell(y, 09).Value = sj_p2dolar;
                worksheet.Cell(y, 10).Value = precio3;
                worksheet.Cell(y, 11).Value = sj_p3dolar;

                toolStripProgressBar1.Value = i;
                i++;
                y++;
            }

            p = (y - 1).ToString();

            worksheet.Range("E6:L" + p).Style.NumberFormat.Format = "#,###.00";

            worksheet.Range("A1:L" + p).Style.Font.FontSize = 9;

            worksheet.Columns(1, 12).AdjustToContents();

            worksheet.Cell(1, 01).Value = misclases.namecompany;
            worksheet.Cell(3, 04).Value = "LISTA DE PRECIOS";
            worksheet.Range("A1:Lk3").Style.Font.FontSize = 14;

            workbook.SaveAs(reporte);

            toolStripProgressBar1.Value = 0;
            toolStripStatusLabel1.Text = "";
            statusStrip1.Refresh();

            Process.Start(reporte);
        }

        private void ListaCostos()
        {
            DataRow[] rowFound = this.dtItems.Select("sj_selec");
            string    reporte  = misclases.FileNameTemp("xlsx"), p;
            int       i        = 0, y = 6;

            toolStripStatusLabel1.Text = "Espere";
            statusStrip1.Refresh();

            toolStripProgressBar1.Maximum = dtItems.Rows.Count;

            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Hoja1");

            worksheet.PageSetup.PaperSize = XLPaperSize.LetterPaper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            worksheet.PageSetup.SetRowsToRepeatAtTop("$1:$5");

            worksheet.Cell(5, 01).Value = "Código";
            worksheet.Cell(5, 02).Value = "Producto";
            worksheet.Cell(5, 03).Value = "Referencia";
            worksheet.Cell(5, 04).Value = "Marca";
            worksheet.Cell(5, 05).Value = "Costo anterior";
            worksheet.Cell(5, 06).Value = "Costo actual";
            worksheet.Cell(5, 07).Value = "Diferencia";
            worksheet.Cell(5, 08).Value = "Stock";
            worksheet.Cell(5, 09).Value = "Costo total";
            worksheet.Cell(5, 10).Value = "Tasa cambio";
            worksheet.Cell(5, 11).Value = "Costo dólar";

            worksheet.Range("A5:L5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A5:L5").Style.Font.Bold = true;
            worksheet.Range("A5:L5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range("A5:L5").Style.Border.RightBorder = XLBorderStyleValues.Thin;

            decimal costact, costant, diferencia, stock, tocosto;

            foreach (DataRow row in rowFound)
            {
                costact    = Convert.ToDecimal(row["costact"]);
                costant    = Convert.ToDecimal(row["costant"]);
                diferencia = (costact - costant);
                stock      = Convert.ToDecimal(row["existen"]);
                costact    = Convert.ToDecimal(row["costact"]);
                tocosto    = (costact * stock);

                worksheet.Cell(y, 01).Value = "'" + row["codprod"].ToString();
                worksheet.Cell(y, 02).Value = "'" + row["descrip"].ToString();
                worksheet.Cell(y, 03).Value = "'" + row["refere"].ToString();
                worksheet.Cell(y, 04).Value = "'" + row["marca"].ToString();
                worksheet.Cell(y, 05).Value = costant;
                worksheet.Cell(y, 06).Value = costact;
                worksheet.Cell(y, 07).Value = diferencia;
                worksheet.Cell(y, 08).Value = stock;
                worksheet.Cell(y, 09).Value = tocosto;
                worksheet.Cell(y, 10).Value = Convert.ToDecimal(row["sj_tasacambio"]);
                worksheet.Cell(y, 11).Value = Convert.ToDecimal(row["sj_costodolar"]);

                toolStripProgressBar1.Value = i;
                i++;
                y++;
            }

            p = (y - 1).ToString();

            worksheet.Range("E6:L" + p).Style.NumberFormat.Format = "#,###.00";

            worksheet.Range("A1:L" + p).Style.Font.FontSize = 9;

            worksheet.Columns(1, 12).AdjustToContents();

            worksheet.Cell(1, 01).Value = misclases.namecompany;
            worksheet.Cell(3, 04).Value = "LISTA DE COSTOS";
            worksheet.Range("A1:Lk3").Style.Font.FontSize = 14;

            workbook.SaveAs(reporte);

            toolStripProgressBar1.Value = 0;
            toolStripStatusLabel1.Text = "";
            statusStrip1.Refresh();

            Process.Start(reporte);
        }

        private void ActualizaEstructuraBD()
        {
            string[,] matriz = {
                                { "saprod",     "sj_selec",         "SMALLINT NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_tasacambio",    "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_costodolar",    "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_p1dolar",       "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_p2dolar",       "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_p3dolar",       "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_putilidad1",    "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_putilidad2",    "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_putilidad3",    "DECIMAL(28,4) NOT NULL DEFAULT 0" },
                                { "saprod",     "sj_feulactualiza", "DATETIME NULL DEFAULT NULL" },
                                };
            DataTable dt1 = new DataTable();
            string    sql;

            for (int i = 0; i < (matriz.Length / 3); i++)
            {
                sql = "select * from information_schema.columns " +
                      "where table_name = '" + matriz[i, 0] + "' and column_name = '" + matriz[i, 1] + "'";
                dt1 = misclases.CursorTable(sql);
                if (dt1.Rows.Count == 0)
                {
                    sql = "alter table " + matriz[i, 0] + " add " + matriz[i, 1] + " " + matriz[i, 2];
                    misclases.CursorTable(sql);
                }
            }
        }

        private void ActcboTasasCambio()
        {
            l_tasascambiodolar tasascambio = new l_tasascambiodolar();
            DataTable dt1 = new DataTable();

            dt1 = tasascambio.select("1 order by id desc");

            cboTasasCambio.DataSource = dt1;
        }

        private void ValidaInicio()
        {
            DateTime appExpira = new DateTime(2023, 06, 30);

            if (DateTime.Now >= appExpira)
            {
                MessageBox.Show("Su licencia calprice expiro", "sysjoma",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }

    }
}
