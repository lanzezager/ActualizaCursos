using ClosedXML.Excel;
using CsvHelper;
using DocumentFormat.OpenXml.Bibliography;
using System.Data;
using System.Formats.Asn1;
using System.Globalization;
using System.Reflection.Emit;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace ActualizaCursos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Tablas Globales
        DataTable dt_datos_nuevos = new DataTable();
        DataTable dt_calificaciones = new DataTable();
        DataTable dt_actividades = new DataTable();


        public string seleccionar_archivo(string tipo_arch)
        {
            string ruta = "";

            OpenFileDialog ofd = new OpenFileDialog();

            if (tipo_arch == "csv")
            {
                ofd.Filter = "Archivos de CSV (*.csv)|*.csv";
            }

            if (tipo_arch == "xls")
            {
                ofd.Filter = "Archivos de Excel (*.xls *.xlsx)|*.xls;*.xlsx";
            }

            if (tipo_arch == "pdf")
            {
                ofd.Filter = "Archivos de Excel (*.pdf)|*.pdf";
            }

            ofd.Title = "Seleccionar el archivo";
            ofd.FileName = string.Empty;

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ruta = ofd.FileName;
            }
            return ruta;
        }

        public DataTable cargar_csv(string archivo)
        {
            DataTable dt = new DataTable();


            using (var reader = new StreamReader(archivo))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    using (var dr = new CsvDataReader(csv))
                    {
                        dt.Load(dr);
                    }
                }
            }


            return dt;
        }

        public DataTable cargar_excel(string archivo)
        {
            DataTable dt = new DataTable();


            using var workbook = new XLWorkbook(archivo);

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                comboBox1.Items.Add(workbook.Worksheet(i + 1).Name);
            }

            var worksheet = workbook.Worksheet(1);

            var rows = worksheet.RowsUsed();

            bool primeraFila = true;

            foreach (var row in rows)
            {
                if (primeraFila)
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }

                    primeraFila = false;
                }
                else
                {
                    DataRow dataRow = dt.NewRow();

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        dataRow[i] = row.Cell(i + 1).Value.ToString();
                    }

                    dt.Rows.Add(dataRow);
                }
            }

            return dt;
        }

        public Object[] exportar_fila_dgv(DataGridViewCellCollection fila_dgv)
        {            
            Object[] valores = new object[fila_dgv.Count];

            for (int i = 0; i < fila_dgv.Count; i++)
            {                
                valores[i]= fila_dgv[i].Value;                
            }

            return valores;
        }

        public DataTable exportar_dgv(int numDataGrid)
        {
            DataTable dt_export = new DataTable();

            switch (numDataGrid)
            {
                case 1:

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        dt_export.Columns.Add(dataGridView1.Columns[i].Name);
                    }

                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView1.Rows[i].Cells));
                    }
                    break;

                case 2:

                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        dt_export.Columns.Add(dataGridView2.Columns[i].Name);
                    }

                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView2.Rows[i].Cells));
                    }
                    break;

                case 3:

                    for (int i = 0; i < dataGridView3.ColumnCount; i++)
                    {
                        dt_export.Columns.Add(dataGridView3.Columns[i].Name);
                    }

                    for (int i = 0; i < dataGridView3.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView3.Rows[i].Cells));
                    }
                    break;

                case 4:

                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                    {
                        dt_export.Columns.Add(dataGridView4.Columns[i].Name);
                    }

                    for (int i = 0; i < dataGridView4.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView4.Rows[i].Cells));
                    }
                    break;

            }

            return dt_export;
        }

        private string copiarTablaPortapapeles(int numDataGrid)
        {
            string textoRespuesta = "", linea = "";

            DataTable dataTable = new DataTable();

            dataTable = exportar_dgv(numDataGrid);

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                linea += dataTable.Columns[i].ColumnName + "\t";
            }

            textoRespuesta += linea + "\r\n";

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                linea = "";

                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    linea += dataTable.Rows[i][j].ToString() + "\t";
                }

                if (i+1 < dataTable.Rows.Count) {
                    textoRespuesta += linea + "\r\n";
                }
                else
                {
                    textoRespuesta += linea;
                }
            }

            return textoRespuesta;
        }

        private System.Data.DataTable adaptarDataTable(string datos, int num_grid)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            int totfilas = 0, totColumnas = 0;
            string letra = "";

            for (int i = 0; i < datos.Length; i++)
            {
                letra = datos[i].ToString();

                if (letra == "\n")
                {
                    totfilas++;
                }

                if (letra == "\t")
                {
                    totColumnas++;
                }
            }

            string[] listLineas = new string[totfilas + 1];
            string[] linea = new string[totColumnas + 1];

            listLineas = datos.Split("\r\n");
            linea = listLineas[0].Split("\t");

            if (num_grid == 1)
            {
                linea = listLineas[0].Split("\t");

                for (int i = 0; i < linea.Length; i++)
                {
                    dt.Columns.Add(linea[i]);
                }
            }

            for (int i = 0; i < listLineas.Length; i++)
            {
                if (i > 0)
                {
                    linea = listLineas[i].Split("\t");
                    dt.Rows.Add(linea);
                }
            }

            return dt;
        }

        private System.Data.DataTable numerarDataTable(System.Data.DataTable original)
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            int num = 0;

            dt.Columns.Add("#", num.GetType());

            for (int i = 0; i < original.Columns.Count; i++)
            {
                dt.Columns.Add(original.Columns[i].ColumnName);
            }

            for (int i = 0; i < original.Rows.Count; i++)
            {
                dt.Rows.Add();
                dt.Rows[i][0] = i + 1;

                for (int j = 0; j < original.Columns.Count; j++)
                {
                    dt.Rows[i][j + 1] = original.Rows[i][j];

                }
            }

            return dt;

        }

        public void pegar_dgv1()
        {
            string textoCopiado = "", ultimo_vacio = "";
            int i = 0;

            if (Clipboard.ContainsData(DataFormats.Text))
            {
                textoCopiado = Clipboard.GetText();

                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 1));
                dataGridView1.Columns["#"].Width = 50;

                i = dataGridView1.Rows.Count - 1;
                /* while (ultimo_vacio.Length == 0)
                 {
                     ultimo_vacio = dataGridView1[1, i].Value.ToString();

                     if (ultimo_vacio.Length == 0)
                     {
                         dataGridView1.Rows.RemoveAt(i);
                     }

                     i--;
                 }*/

                label12.Text = "Registros: " + dataGridView1.Rows.Count;

                if (dataGridView1.Rows.Count > 0)
                {

                }
            }
        }

        public void pegar_dgv2()
        {
            string textoCopiado = "", ultimo_vacio = "";
            int i = 0;

            if (Clipboard.ContainsData(DataFormats.Text))
            {
                textoCopiado = Clipboard.GetText();

                dataGridView2.DataSource = null;
                dataGridView2.Columns.Clear();
                dataGridView2.Rows.Clear();
                dataGridView2.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 2));
                dataGridView2.Columns["#"].Width = 50;

                i = dataGridView2.Rows.Count - 1;

                /* while (ultimo_vacio.Length == 0)
                 {
                     ultimo_vacio = dataGridView2[1, i].Value.ToString();

                     if (ultimo_vacio.Length == 0)
                     {
                         dataGridView2.Rows.RemoveAt(i);
                     }

                     i--;
                 }*/

                label1.Text = "Registros: " + dataGridView2.Rows.Count;

                if (dataGridView2.Rows.Count > 0)
                {

                }
            }

        }

        public void copiar_datos()
        {
            if (dataGridView1.ContainsFocus == true)
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    Clipboard.SetText(copiarTablaPortapapeles(1));
                }
            }

            if (dataGridView2.ContainsFocus == true)
            {
                if (dataGridView2.Rows.Count > 0)
                {
                    Clipboard.SetText(copiarTablaPortapapeles(2));
                }
            }

            if (dataGridView3.ContainsFocus == true)
            {
                if (dataGridView3.Rows.Count > 0)
                {
                    Clipboard.SetText(copiarTablaPortapapeles(3));
                }
            }

            if (dataGridView4.ContainsFocus == true)
            {
                if (dataGridView4.Rows.Count > 0)
                {
                    Clipboard.SetText(copiarTablaPortapapeles(4));
                }
            }

        }

        public void obtener_nuevos()
        {
            string email = "",repe="";
            int coincidencias = 0, repetido=0;

            for (int i=0; i<dt_datos_nuevos.Columns.Count;i++)
            {
                dataGridView2.Columns.Add(dt_datos_nuevos.Columns[i].ColumnName, dt_datos_nuevos.Columns[i].ColumnName);
            }
            
            for (int i = 0; i < dt_datos_nuevos.Rows.Count; i++) 
            {
                email = dt_datos_nuevos.Rows[i][5].ToString();
                coincidencias = 0;

                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    if (email== dataGridView1[5+1,j].Value.ToString())
                    {
                        coincidencias++;
                    }
                }

                if (coincidencias==0)
                {
                    repetido = 0;
                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                    {
                        if (dataGridView2.Rows.Count > 1)
                        {
                            repe = dataGridView2[5,j].Value.ToString();
                            if (email == repe)
                            {
                                repetido++;
                            }
                        }
                    }

                    if (repetido == 0)
                    {
                        //MessageBox.Show(dt_datos_nuevos.Rows[i].ItemArray.ToString());
                        dataGridView2.Rows.Add(dt_datos_nuevos.Rows[i].ItemArray);
                    }                    
                }
            }
            
            label2.Text = "Registros: "+dataGridView2.Rows.Count;

        }

        public void procesar()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            dataGridView4.Rows.Clear();
            dataGridView4.Columns.Clear();



            obtener_nuevos();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            textBox1.Text = seleccionar_archivo("csv");
            textBox1.SelectionStart = textBox1.Text.Length;

            if (textBox1.Text.Length > 0)
            {
                dt_datos_nuevos.Rows.Clear();
                dt_datos_nuevos.Columns.Clear();
                dt_datos_nuevos = cargar_csv(textBox1.Text);
            }


            //dataGridView1.DataSource = dt_datos_nuevos;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = string.Empty;

            if (radioButton1.Checked)
            {
                textBox2.Text = seleccionar_archivo("xls");
                textBox2.SelectionStart = textBox2.Text.Length;

                if (textBox2.Text.Length > 0)
                {
                    dt_calificaciones.Rows.Clear();
                    dt_calificaciones.Columns.Clear();
                    dt_calificaciones = cargar_excel(textBox2.Text);
                }
            }

            if (radioButton2.Checked)
            {
                textBox2.Text = seleccionar_archivo("csv");
                textBox2.SelectionStart = textBox2.Text.Length;

                if (textBox2.Text.Length > 0)
                {
                    dt_calificaciones.Rows.Clear();
                    dt_calificaciones.Columns.Clear();
                    dt_calificaciones = cargar_csv(textBox2.Text);
                }
            }
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            textBox3.Text = string.Empty;

            if (radioButton1.Checked)
            {
                textBox3.Text = seleccionar_archivo("csv");
                textBox3.SelectionStart = textBox3.Text.Length;

                if (textBox3.Text.Length > 0)
                {
                    dt_actividades.Rows.Clear();
                    dt_actividades.Columns.Clear();
                    dt_actividades = cargar_csv(textBox3.Text);
                }

            }

            if (radioButton2.Checked)
            {
                textBox3.Text = seleccionar_archivo("pdf");
                textBox3.SelectionStart = textBox3.Text.Length;

                if (textBox3.Text.Length > 0)
                {
                    dt_actividades.Rows.Clear();
                    dt_actividades.Columns.Clear();
                    //dt_actividades = cargar_excel(textBox3.Text);
                }
            }

        }
        
        private void button4_Click(object sender, EventArgs e)
        {
            procesar();
            tabControl1.SelectedIndex = 2;
            dataGridView2.Focus();
            MessageBox.Show("El Análisis ha terminado Correctamente", "Analisis Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void copiarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            copiar_datos();
        }

        private void pegarToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (dataGridView1.ContainsFocus == true)
            {
                pegar_dgv1();
            }
            else
            {

                if (dataGridView2.ContainsFocus == true)
                {
                    pegar_dgv2();
                }
            }
        }

        private void copiarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.ContainsFocus == true)
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    Clipboard.SetText(copiarTablaPortapapeles(1));
                }
            }
        }

        private void pegarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            pegar_dgv1();
        }
        
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            copiar_datos();
        }
        
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            label6.Text = "Archivo CSV  Reporte de Actividades";
            label4.Text = "Archivo XLSX  Reporte de Calificaciones";
            groupBox1.BackColor = System.Drawing.Color.RosyBrown;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            label6.Text = "Archivo PDF Reporte de Actividades";
            label4.Text = "Archivo CSV  Reporte de Calificaciones";
            groupBox1.BackColor = System.Drawing.Color.CadetBlue;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            if (textBox1.Text.Length>0 && textBox2.Text.Length>0 && textBox3.Text.Length>0)
            {
                button4.Enabled = true;
            }
            else
            {
                button4.Enabled = false;
            }
        }
    
    }
}
