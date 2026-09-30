using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Bibliography;
using System.Data;
using System.Formats.Asn1;
using System.Globalization;
using System.Reflection.Emit;
using System.Security.Cryptography;
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

        //Tot Repetidos
        int tot_repetidos = 0;

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

        public string corrigeTel(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            // Si contiene notación científica (E+), lo convertimos numéricamente primero
            if (text.Contains("E+", StringComparison.OrdinalIgnoreCase))
            {
                // Reemplazamos la coma por punto si es necesario para la cultura invariable
                string limpio = text.Replace(',', '.');
                if (double.TryParse(limpio, NumberStyles.Any, CultureInfo.InvariantCulture, out double numero))
                {
                    // Convertimos el double a un string sin decimales (formato largo)
                    return numero.ToString("F0", CultureInfo.InvariantCulture);
                }
            }

            return text; // Si ya es texto normal, lo devuelve tal cual
        }

        public DataTable cargar_csv(string archivo)
        {
            DataTable dt = new DataTable();
            FileStreamOptions opcionesArchivo = new FileStreamOptions();

            opcionesArchivo.Share = FileShare.ReadWrite;
            opcionesArchivo.Access = FileAccess.Read;
            opcionesArchivo.Mode = FileMode.Open;

            CsvConfiguration config = new CsvConfiguration(CultureInfo.InvariantCulture);
            config.Delimiter = ",";

            try
            {
                using (var reader = new StreamReader(archivo, opcionesArchivo))
                {
                    using (var csv = new CsvReader(reader, config))
                    {
                        using (var dr = new CsvDataReader(csv))
                        {
                            dt.Load(dr);
                        }
                    }
                }

                if (dt.Columns.Count == 1)
                {
                    dt.Rows.Clear();
                    dt.Columns.Clear();                   

                    config.Delimiter = "\t";

                    using (var reader = new StreamReader(archivo, opcionesArchivo))
                    {
                        using (var csv = new CsvReader(reader, config))
                        {
                            using (var dr = new CsvDataReader(csv))
                            {
                                dt.Load(dr);
                            }
                        }
                    }
                }

            }
            catch(BadDataException ex)
            {
                try
                {
                    config.Mode = CsvMode.RFC4180;

                    using (var reader = new StreamReader(archivo, opcionesArchivo))
                    {
                        using (var csv = new CsvReader(reader, config))
                        {
                            using (var dr = new CsvDataReader(csv))
                            {
                                dt.Load(dr);
                            }
                        }
                    }

                    if (dt.Columns.Count == 1)
                    {
                        dt.Rows.Clear();
                        dt.Columns.Clear();

                        config.Delimiter = "\t";

                        using (var reader = new StreamReader(archivo, opcionesArchivo))
                        {
                            using (var csv = new CsvReader(reader, config))
                            {
                                using (var dr = new CsvDataReader(csv))
                                {
                                    dt.Load(dr);
                                }
                            }
                        }
                    }
                }
                catch(BadDataException ex2) 
                { 

                }
            }

            return dt;
        }

        public DataTable cargar_excel(string archivo)
        {
            DataTable dt = new DataTable();
            FileStreamOptions opcionesArchivo = new FileStreamOptions();

            opcionesArchivo.Share = FileShare.ReadWrite;
            opcionesArchivo.Access = FileAccess.Read;
            opcionesArchivo.Mode = FileMode.Open;

            FileStream fs = new FileStream(archivo, opcionesArchivo);

            using var workbook = new XLWorkbook(fs);

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
                valores[i] = fila_dgv[i].Value;
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
                        dt_export.Columns.Add(dataGridView1.Columns[i].HeaderText);
                    }

                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView1.Rows[i].Cells));
                    }
                    break;

                case 2:

                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        dt_export.Columns.Add(dataGridView2.Columns[i].HeaderText);
                    }

                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView2.Rows[i].Cells));
                    }
                    break;

                case 3:

                    for (int i = 0; i < dataGridView3.ColumnCount; i++)
                    {
                        dt_export.Columns.Add(dataGridView3.Columns[i].HeaderText);
                    }

                    for (int i = 0; i < dataGridView3.RowCount; i++)
                    {
                        dt_export.Rows.Add(exportar_fila_dgv(dataGridView3.Rows[i].Cells));
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

                if (i + 1 < dataTable.Rows.Count)
                {
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
                //dataGridView1.DataSource = numerarDataTable(adaptarDataTable(textoCopiado, 1));
                dataGridView1.DataSource = adaptarDataTable(textoCopiado, 1);
                //dataGridView1.Columns["#"].Width = 50;

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

        }

        public void obtener_nuevos()
        {
            string email = "", repe = "";
            int coincidencias = 0, repetido = 0;

            for (int i = 0; i < dt_datos_nuevos.Columns.Count; i++)
            {
                dataGridView2.Columns.Add(dt_datos_nuevos.Columns[i].ColumnName, dt_datos_nuevos.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt_datos_nuevos.Rows.Count; i++)
            {
                email = dt_datos_nuevos.Rows[i][5].ToString();
                coincidencias = 0;

                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    if (email == dataGridView1[5, j].Value.ToString())
                    {
                        coincidencias++;
                    }
                }

                if (coincidencias == 0)
                {
                    repetido = 0;
                    for (int j = 0; j < dataGridView2.Rows.Count; j++)
                    {
                        if (dataGridView2.Rows.Count > 1)
                        {
                            repe = dataGridView2[5, j].Value.ToString();
                            if (email == repe)
                            {
                                repetido++;
                            }
                        }
                    }

                    if (repetido == 0)
                    {
                        //MessageBox.Show(dt_datos_nuevos.Rows[i].ItemArray.ToString());
                        //dt_datos_nuevos.Rows[i][7] = corrigeTel(dt_datos_nuevos.Rows[i][7].ToString());
                        dataGridView2.Rows.Add(dt_datos_nuevos.Rows[i].ItemArray);
                    }
                    else
                    {
                        //tot_repetidos++;
                        //MessageBox.Show(dt_datos_nuevos.Rows[i].ItemArray[5].ToString());
                    }
                }
            }

            label2.Text = "Registros: " + dataGridView2.Rows.Count;

        }

        public Decimal calificaciones(string correo_buscar)
        {
            Decimal calificacion = 0;
            int col_cal = 0, col_correo = 0;

            if (radioButton1.Checked)
            {
                col_correo = 5;
                col_cal = dt_calificaciones.Columns.Count - 2;
            }
            else
            {
                col_correo = 1;
                col_cal = dt_calificaciones.Columns.Count - 1;
            }

            for (int i = 0; i < dt_calificaciones.Rows.Count; i++)
            {
                if (correo_buscar == dt_calificaciones.Rows[i][col_correo].ToString())
                {
                    if (Decimal.TryParse(dt_calificaciones.Rows[i][col_cal].ToString(), out calificacion))
                    {

                    }
                }
            }

            return calificacion;

        }

        public void actividades()
        {
            string email = "", formato_fecha = "dd/MM/yyyy HH:mm";
            int coincidencia = 0,repetido=0, col_cont=1;
            int[] asistencia = new int[12];
            DateTime fecha_posible = DateTime.Now;
            Decimal calificacion = 0, min_apro = 0;

            if (Decimal.TryParse(numericUpDown1.Value.ToString(), out min_apro))
            {
                //no va nada
            }

            for (int i = 0; i < dt_datos_nuevos.Rows.Count; i++)
            {
                email = dt_datos_nuevos.Rows[i][5].ToString();
                coincidencia = 0;

                for (int z = 0; z < asistencia.Length; z++)
                {
                    asistencia[z] = 0;
                }

                for (int j = 0; j < dt_actividades.Rows.Count; j++)
                {
                    if (email == dt_actividades.Rows[j][1].ToString())
                    {
                        coincidencia = 1;

                        for (int k = 0; k < dt_actividades.Columns.Count; k++)
                        {
                            if (DateTime.TryParse(dt_actividades.Rows[j][k].ToString(), out fecha_posible))
                            {
                                fecha_posible = DateTime.ParseExact(dt_actividades.Rows[j][k].ToString(), formato_fecha, CultureInfo.InvariantCulture);

                                if (fecha_posible.Year != 1969)
                                {
                                    asistencia[fecha_posible.Month - 1] = 1;
                                }

                            }
                        }
                    }
                }


                Object[] valores = new object[dataGridView3.Columns.Count];
                int tot_asistencias = 0;
                repetido = 0;

                valores[0] = col_cont++;
                valores[1] = dt_datos_nuevos.Rows[i][1].ToString();//matricula
                valores[2] = dt_datos_nuevos.Rows[i][2].ToString();//nombre (s)
                valores[3] = dt_datos_nuevos.Rows[i][3].ToString();//primer apellido
                valores[4] = dt_datos_nuevos.Rows[i][4].ToString();//segundo apellido
                valores[5] = dt_datos_nuevos.Rows[i][5].ToString();//email


                for (int a = 0; a < asistencia.Length; a++)
                {
                    tot_asistencias += asistencia[a];
                }

                
                valores[6] = tot_asistencias;//tot_asistencias
                valores[7] = tot_asistencias;//tot_modulos_activos

                valores[8] = asistencia[0];//enero
                valores[9] = asistencia[1];//febrero
                valores[10] = asistencia[2];//marzo
                valores[11] = asistencia[3];//abril
                valores[12] = asistencia[4];//mayo
                valores[13] = asistencia[5];//junio
                valores[14] = asistencia[6];//julio
                valores[15] = asistencia[7];//agosto
                valores[16] = asistencia[8];//septiembre
                valores[17] = asistencia[9];//octubre
                valores[18] = asistencia[10];//noviembre
                valores[19] = asistencia[11];//diciembre
                
                valores[20] = " ";//deserción

                calificacion = calificaciones(dt_datos_nuevos.Rows[i][5].ToString());
                valores[21] = calificacion;//calificacion


                if (calificacion >= min_apro)//certificado
                {
                    valores[22] = "Certificado";
                }
                else
                {
                    valores[22] = "No Aprobado";
                }

                for(int j = 0; j < dataGridView3.Rows.Count; j++)
                {
                    if (email == dataGridView3[5,j].Value.ToString())
                    {
                        repetido++;
                    }
                }

                if (repetido==0) 
                {
                    dataGridView3.Rows.Add(valores);
                }
                else
                {
                    tot_repetidos++;
                    col_cont--;
                }

            }

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                dataGridView3[20, i].Style.BackColor = SystemColors.ControlDark;
            }

            label3.Text = "Registros: " + dataGridView3.Rows.Count.ToString();

            

        }

        public void procesar()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            dataGridView3.Rows.Clear();

            obtener_nuevos();
            /*
            for(int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2[7, i].Value = corrigeTel(dataGridView2[7, i].Value.ToString());
            }
            */
            actividades();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            for (int i = 0; i < dataGridView3.Columns.Count; i++)
            {
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            //MessageBox.Show(DateTime.Today.Month.ToString());
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
                    //dataGridView2.DataSource = dt_actividades;
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
            tot_repetidos = 0;
            procesar();
            tabControl1.SelectedIndex = 2;
            dataGridView2.Focus();
            MessageBox.Show("El Análisis ha terminado Correctamente\nSe quitaron: "+tot_repetidos+" registros repetidos.", "Analisis Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

            label5.Text = "Repetidos Ignorados: " + tot_repetidos;
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

            if (textBox1.Text.Length > 0 && textBox2.Text.Length > 0 && textBox3.Text.Length > 0)
            {
                button4.Enabled = true;
            }
            else
            {
                button4.Enabled = false;
            }

            if (dataGridView2.Rows.Count > 0 && dataGridView3.Rows.Count > 0)
            {
                button5.Enabled = true;
                button6.Enabled = true;
            }
            else
            {
                button5.Enabled = false;
                button6.Enabled = false;
            }
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // Crear el texto del número (índice base 1)
            string rowNumber = (e.RowIndex + 1).ToString();

            // Configurar la fuente y el color del texto
            Font font = new Font("Arial", 9, FontStyle.Bold);
            Brush brush = Brushes.Black;

            // Calcular la posición del texto para que quede centrado
            StringFormat centerFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            // Definir los límites del área del RowHeader
            Rectangle headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dataGridView1.RowHeadersWidth, e.RowBounds.Height);

            // Dibujar el número en el RowHeader
            e.Graphics.DrawString(rowNumber, font, brush, headerBounds, centerFormat);
        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // Crear el texto del número (índice base 1)
            string rowNumber = (e.RowIndex + 1).ToString();

            // Configurar la fuente y el color del texto
            Font font = new Font("Arial", 9, FontStyle.Bold);
            Brush brush = Brushes.Black;

            // Calcular la posición del texto para que quede centrado
            StringFormat centerFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            // Definir los límites del área del RowHeader
            Rectangle headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dataGridView2.RowHeadersWidth, e.RowBounds.Height);

            // Dibujar el número en el RowHeader
            e.Graphics.DrawString(rowNumber, font, brush, headerBounds, centerFormat);
        }

        private void dataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // Crear el texto del número (índice base 1)
            string rowNumber = (e.RowIndex + 1).ToString();

            // Configurar la fuente y el color del texto
            Font font = new Font("Arial", 9, FontStyle.Bold);
            Brush brush = Brushes.Black;

            // Calcular la posición del texto para que quede centrado
            StringFormat centerFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            // Definir los límites del área del RowHeader
            Rectangle headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dataGridView3.RowHeadersWidth, e.RowBounds.Height);

            // Dibujar el número en el RowHeader
            e.Graphics.DrawString(rowNumber, font, brush, headerBounds, centerFormat);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                dataGridView3.Columns[0].Frozen = true;
                dataGridView3.Columns[1].Frozen = true;
                dataGridView3.Columns[2].Frozen = true;
                dataGridView3.Columns[3].Frozen = true;
                dataGridView3.Columns[4].Frozen = true;

            }
            else
            {
                dataGridView3.Columns[0].Frozen = false;
                dataGridView3.Columns[1].Frozen = false;
                dataGridView3.Columns[2].Frozen = false;
                dataGridView3.Columns[3].Frozen = false;
                dataGridView3.Columns[4].Frozen = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count > 0)
            {
                SaveFileDialog guardar = new SaveFileDialog();

                guardar.Filter = "Archivos de Excel (*.xlsx)|*.xlsx";
                guardar.Title = "Guardar archivo";

                if (guardar.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(exportar_dgv(2), "Alumnos Nuevos");
                    wb.Worksheets.Add(exportar_dgv(3), "Actividades-Calificaciones");
                    wb.SaveAs(@"" + guardar.FileName);
                    MessageBox.Show("Archivo guardado correctamente", "Exito", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            //dataGridView1.Rows.Clear();
            //dataGridView1.Columns.Clear();

            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            dataGridView3.Rows.Clear();    

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";

            label12.Text = "Registros:";
            label2.Text = "Registros:";
            label3.Text = "Registros:";
            numericUpDown1.Value = 70;

            tabControl1.SelectedIndex = 0;
        }
    
    }
}
