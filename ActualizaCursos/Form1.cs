using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Bibliography;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
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
        int tot_repetidos=0,tot_repetidos2 = 0;

        //Tablas Globales
        DataTable dt_datos_base = new DataTable();
        DataTable dt_datos_nuevos = new DataTable();
        DataTable dt_calificaciones = new DataTable();
        DataTable dt_actividades = new DataTable();        

        public string[] seleccionar_archivo(string tipo_arch)
        {
            //string ruta = "";
            string[] rutas;
            
            rutas = new string[1];

            OpenFileDialog ofd = new OpenFileDialog();

            if (tipo_arch == "csv")
            {
                ofd.Filter = "Archivos de CSV (*.csv)|*.csv";
                ofd.Multiselect = false;
            }

            if (tipo_arch == "xls")
            {
                ofd.Filter = "Archivos de Excel (*.xls *.xlsx)|*.xls;*.xlsx";
                ofd.Multiselect = false;
            }

            if (tipo_arch == "pdf")
            {
                ofd.Filter = "Archivos de Excel (*.pdf)|*.pdf";
                ofd.Multiselect = true;
            }

            ofd.Title = "Seleccionar el archivo";
            ofd.FileName = string.Empty;

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (tipo_arch != "pdf") {
                    rutas[0] = ofd.FileName;
                }
                else
                {
                    rutas = new string[ofd.FileNames.Length];

                    for (int i = 0; i < ofd.FileNames.Length; i++)
                    {
                        rutas[i] = ofd.FileNames[i];
                    }
                }
            }
            return rutas;
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

        public CsvConfiguration verificar_csv(string archivo)
        {
            CsvConfiguration config = new CsvConfiguration(CultureInfo.InvariantCulture);

            if (System.IO.File.Exists(archivo))
            {
                FileStreamOptions opcionesArchivo = new FileStreamOptions();

                opcionesArchivo.Share = FileShare.ReadWrite;
                opcionesArchivo.Access = FileAccess.Read;
                opcionesArchivo.Mode = FileMode.Open;

                // El 'using' asegura que el archivo se cierre automáticamente al terminar
                using (StreamReader sr = new StreamReader(archivo,opcionesArchivo))
                {
                    string linea;
                    // ReadLine() devuelve null cuando llega al final del archivo
                    linea = sr.ReadLine();                    

                    if (linea.Contains("\t"))
                    {
                        config.Delimiter = "\t";
                    }
                    else
                    {
                        if (linea.Contains(","))
                        {
                            config.Delimiter = ",";
                        }
                        else
                        {
                            if (linea.Contains(";"))
                            {
                                config.Delimiter = ";";
                            }
                        }
                    }

                    if (linea.Contains("\""))
                    {
                        config.Mode = CsvMode.RFC4180;
                    }


                   /* while ((linea = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(linea); // Procesa la línea aquí
                    }*/
                }
            }
            else
            {
                Console.WriteLine("El archivo no existe.");
            }


            return config;
        }

        public DataTable cargar_csv(string archivo)
        {
            DataTable dt = new DataTable();
            FileStreamOptions opcionesArchivo = new FileStreamOptions();

            opcionesArchivo.Share = FileShare.ReadWrite;
            opcionesArchivo.Access = FileAccess.Read;
            opcionesArchivo.Mode = FileMode.Open;

            CsvConfiguration config = new CsvConfiguration(CultureInfo.InvariantCulture);

            config = verificar_csv(archivo);

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
               
            }
            catch(BadDataException ex)
            {
                MessageBox.Show(ex.Message);                
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

        public string leer_pdf(string archivo)
        {
            FileStreamOptions opcionesArchivo = new FileStreamOptions();

            opcionesArchivo.Share = FileShare.ReadWrite;
            opcionesArchivo.Access = FileAccess.Read;
            opcionesArchivo.Mode = FileMode.Open;

            FileStream fs = new FileStream(archivo, opcionesArchivo);

            using (PdfReader lector = new PdfReader(fs))
            {
                using (PdfDocument pdfDoc = new PdfDocument(lector))
                {
                    System.Text.StringBuilder textoTotal = new System.Text.StringBuilder();

                    // Recorrer cada una de las páginas del PDF
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        // Extraer el texto de la página actual
                        
                        string textoPagina = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i));
                        textoTotal.AppendLine(textoPagina);
                        /*
                        string[] lineas = textoPagina.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                        // Imprimir o procesar cada línea
                        foreach (string linea in lineas)
                        {
                            textBox4.AppendText(linea);
                        }*/
                    }

                    return textoTotal.ToString();
                }
            }
        }

        public DataTable cargar_pdf(string[] archivo)
        {
            DataTable dt = new DataTable();
            DateTime fecha_hoy= DateTime.Now;
            string texto_bruto = "",linea="", fecha="", alumno="", alumn_n="", alumn_a="";
            int check = 0,it=-1,cal_global=0,tot_cols=0,ical=-1, verif_fecha=0, pos_pri_mes=0, 
                paquete=0, apell=0, fin_nom=0;
            string[] meses = new string[] {"enero","febrero","marzo","abril","mayo","junio",
                                           "julio","agosto","septiembre","octubre","noviembre","diciembre"};

            for (int i = 0; i < archivo.Length; i++)
            {
               texto_bruto+=leer_pdf(archivo[i]);
            }

            texto_bruto = texto_bruto.ToLower();

            string[] lineas = texto_bruto.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            dt.Columns.Add("Alumno");
            dt.Columns.Add("Email");
            dt.Columns.Add("Class_Grade");
            dt.Columns.Add("Final Exam");
            dt.Columns.Add("Promedio");

            textBox4.Text = "";
            it = -1;
            tot_cols = 0;

            // Imprimir o procesar cada línea
            //foreach (string linea in lineas)
            for(int i = 0;i < lineas.Length;i++)
            {
                linea=lineas[i];

                //asignar columnas totales
                if (linea.Contains("due") && tot_cols==0)
                {
                    string[] palabras = linea.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    
                    for(int j = 0; j < palabras.Length; j++)
                    {
                        if (palabras[j] == "due")
                        {
                            tot_cols++;
                        }
                    }

                    for(int j = 0;j < tot_cols; j++)
                    {
                        dt.Columns.Add();
                        dt.Columns.Add();
                    }
                }

                
                //buscar bloque
                if (linea.Contains('%'))
                {
                    //inicio bloque
                    if (linea.Substring(linea.IndexOf('%') - 1, 1) != "(") {
                        check = 1;                      
                    }

                }

                //crear bloque
                if (check == 1)
                {
                    if (linea.Contains('@'))
                    {
                        check = 2;
                    }

                    if (check != 2) {
                        textBox4.AppendText(linea + "\n");
                    }
                    else
                    {
                        textBox4.AppendText(linea);
                    }
                }

                //analizar bloque
                if (check==2)
                {
                    if (lineas[i+1].Contains('%')==false)
                    {
                        if (lineas[i + 1].Contains("final exam")==false) {
                            textBox4.Text += lineas[i + 1];
                        }
                       
                    }

                    if (textBox4.Text.Length>0)
                    {
                        dt.Rows.Add();
                        it++;
                        cal_global = 0;
                        ical = 5;

                        for (int j = 0; j < textBox4.Lines.Count(); j++)
                        {
                            string[] palabras = textBox4.Lines[j].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                            //calificaciones globales
                            if (palabras[0] == "0.00%")
                            {
                                dt.Rows[it][2] = "0.00%";
                                dt.Rows[it][3] = "--";
                                dt.Rows[it][4] = "0.00%";
                                cal_global = 1;
                            }
                            else
                            {
                                if (cal_global == 0)
                                {
                                    dt.Rows[it][2] = palabras[0];
                                    dt.Rows[it][3] = palabras[1];
                                    dt.Rows[it][4] = palabras[2];
                                    cal_global = 1;
                                }

                                //buscar calificaciones
                                if (textBox4.Lines[j].Contains("/"))
                                {
                                    for (int k = 0; k < palabras.Length; k++)
                                    {
                                        if (palabras[k].Contains("/"))
                                        {
                                            dt.Rows[it][ical] = palabras[k];
                                            ical++;
                                        }
                                    }
                                }

                                //buscar fechas
                                for (int k = 0; k < meses.Length; k++)
                                {
                                    if (textBox4.Lines[j].Contains(meses[k]))
                                    {
                                        if ((textBox4.Lines[j].Contains('@')==false) && (textBox4.Lines[j].Contains(',')==false)) {
                                            verif_fecha = 1;
                                            pos_pri_mes = j;
                                            break;
                                        }
                                    }
                                }

                                //escribir fecha
                                if (verif_fecha == 1)
                                {
                                    pos_pri_mes = -1;

                                    for (int k = 0; k < palabras.Length; k++)
                                    {
                                        for (int l = 0; l < meses.Length; l++)
                                        {
                                            if (palabras[k].Contains(meses[l]))
                                            {
                                                if (pos_pri_mes==-1) {
                                                    pos_pri_mes = k; 
                                                }
                                            }
                                        }
                                    }

                                    for (int k = (pos_pri_mes-1); k < palabras.Length; k++)
                                    {
                                        if (paquete < 3)
                                        {
                                            fecha += palabras[k] + " ";
                                            paquete++;
                                         }

                                        if (paquete == 3)
                                        {
                                            fecha = fecha + "00:00";

                                            if (DateTime.TryParse(fecha, out fecha_hoy))
                                            {
                                                for (int l = 0; l < tot_cols; l++) {

                                                    if (dt.Rows[it][(ical - tot_cols)].ToString().Contains("--"))
                                                    {
                                                        ical++;
                                                    }
                                                    else
                                                    {
                                                        break;
                                                    }
                                                }

                                                dt.Rows[it][ical] = fecha_hoy;
                                            }

                                            fecha = "";
                                            paquete = 0;
                                            ical++;
                                        }
                                        

                                    }

                                    verif_fecha = 0;
                                }

                                //buscar nombre
                                if (textBox4.Lines[j].Contains(','))
                                {
                                    alumno = "";
                                    alumn_n = "";
                                    alumn_a = "";

                                    if (textBox4.Lines[j].Contains('/'))
                                    {
                                        apell = 1;
                                        fin_nom = 0;

                                        for (int k = 0; k < palabras.Length; k++)
                                        {
                                            if (palabras[k].Contains('/'))
                                            {
                                                fin_nom = 1;
                                            }

                                            if (fin_nom == 0) {

                                                if (apell == 1)
                                                {
                                                    alumn_a += palabras[k] + " ";
                                                }
                                                else
                                                {
                                                    alumn_n += palabras[k] + " ";
                                                }

                                                if (palabras[k].Contains(","))
                                                {
                                                    apell = 0;
                                                }

                                            }

                                        }
                                        alumn_a = alumn_a.Substring(0, alumn_a.Length - 1);
                                        alumno = alumn_n + alumn_a.Substring(0, alumn_a.Length - 1);
                                        dt.Rows[it][0] = alumno;
                                    }
                                    else
                                    {
                                        apell = 1;
                                        fin_nom = 0;

                                        for (int k = 0; k < palabras.Length; k++)
                                        {
                                            if (fin_nom == 0)
                                            {
                                                if (apell == 1)
                                                {
                                                    alumn_a += palabras[k] + " ";
                                                }
                                                else
                                                {
                                                    alumn_n += palabras[k] + " ";
                                                }

                                                if (palabras[k].Contains(','))
                                                {
                                                    apell = 0;
                                                }

                                            }

                                            if (palabras[k].Contains(''))
                                            {
                                                fin_nom = 1;
                                            }
                                        }

                                        while (alumn_n.Contains(''))
                                        {
                                            alumn_n = alumn_n.Substring(0, alumn_n.Length - 1);
                                        }

                                        alumn_a = alumn_a.Substring(0, alumn_a.Length - 1);

                                        alumno = alumn_n + alumn_a.Substring(0, alumn_a.Length - 1);
                                        dt.Rows[it][0] = alumno;

                                    }
                                }


                                //buscar correo

                                if (textBox4.Lines[j].Contains('@'))
                                {
                                    dt.Rows[it][1] = textBox4.Lines[j];
                                }
                            }
                        }

                    }

                    textBox4.Text = "";
                    check = 0;
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

                if (dataGridView1.Columns.Count>0) 
                {
                    dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }

                //dataGridView1.Columns["#"].Width = 50;
                /* 
                i = dataGridView1.Rows.Count - 1;
                while (ultimo_vacio.Length == 0)
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
                    if (email == dataGridView1[0, j].Value.ToString())
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
                        dt_datos_base.Rows.Add(dt_datos_nuevos.Rows[i][5].ToString());
                    }
                    else
                    {
                        tot_repetidos++;
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

        public int[] busca_fechas_plai(string email, int indice_email)
        {
            string formato_fecha = "dd/MM/yyyy HH:mm";
            int[] asistencia = new int[12];
            DateTime fecha_posible = DateTime.Now;

            for (int z = 0; z < asistencia.Length; z++)
            {
                asistencia[z] = 0;
            }

            for (int j = 0; j < dt_actividades.Rows.Count; j++)
            {
                if (email == dt_actividades.Rows[j][indice_email].ToString())
                {
                    for (int k = 0; k < dt_actividades.Columns.Count; k++)
                    {
                        if (DateTime.TryParse(dt_actividades.Rows[j][k].ToString(), out fecha_posible))
                        {
                            //fecha_posible = DateTime.ParseExact(dt_actividades.Rows[j][k].ToString(), formato_fecha, CultureInfo.InvariantCulture);
                            //if (DateTime.TryParseExact(dt_actividades.Rows[j][k].ToString(),formato_fecha, CultureInfo.InvariantCulture,DateTimeStyles.None, out fecha_posible_exacta)) {

                            if (fecha_posible.Year != 1969)
                            {
                                asistencia[fecha_posible.Month - 1] = 1;
                            }
                            //}
                        }
                    }
                }
            }

            return asistencia;
        }

        public int[] busca_fechas_netacad(string email, int indice_email)
        {
            string formato_fecha = "dd/MM/yyyy HH:mm";
            int[] asistencia = new int[12];
            DateTime fecha_posible = DateTime.Now;

            for (int z = 0; z < asistencia.Length; z++)
            {
                asistencia[z] = 0;
            }

            for (int j = 0; j < dt_actividades.Rows.Count; j++)
            {
                if (email == dt_actividades.Rows[j][indice_email].ToString())
                {
                    for (int k = 0; k < dt_actividades.Columns.Count; k++)
                    {
                        if (DateTime.TryParse(dt_actividades.Rows[j][k].ToString(), out fecha_posible))
                        {
                            //fecha_posible = DateTime.ParseExact(dt_actividades.Rows[j][k].ToString(), formato_fecha, CultureInfo.InvariantCulture);
                            //if (DateTime.TryParseExact(dt_actividades.Rows[j][k].ToString(),formato_fecha, CultureInfo.InvariantCulture,DateTimeStyles.None, out fecha_posible_exacta)) {

                            if (fecha_posible.Year != 1969)
                            {
                                asistencia[fecha_posible.Month - 1] = 1;
                            }
                            //}

                        }
                    }
                }
            }

            return asistencia;
        }

        public void actividades()
        {
            string email = "";
            int repetido=0, col_cont=1,indice=-1;
            int[] asistencia = new int[12];
            
            Decimal calificacion = 0, min_apro = 0;

            if (Decimal.TryParse(numericUpDown1.Value.ToString(), out min_apro))
            {
                //no va nada
            }

            for (int i = 0; i < dt_datos_base.Rows.Count; i++)
            {
                email = dt_datos_base.Rows[i][0].ToString();
                indice = -1;                
                              
                if (radioButton1.Checked)
                {
                    asistencia = busca_fechas_plai(email,1);
                }

                if (radioButton2.Checked)
                {
                    asistencia = busca_fechas_netacad(email,1);
                }
                 
                for (int l = 0; l < dt_datos_nuevos.Rows.Count; l++)
                {
                    if (email == dt_datos_nuevos.Rows[l][5].ToString())
                    {
                        indice = l;
                        break;
                    }
                }

                if (indice>=0) { 
                    Object[] valores = new object[dataGridView3.Columns.Count];
                    int tot_asistencias = 0;
                    repetido = 0;

                    valores[0] = col_cont++;
                    valores[1] = dt_datos_nuevos.Rows[indice][1].ToString();//matricula
                    valores[2] = dt_datos_nuevos.Rows[indice][2].ToString();//nombre (s)
                    valores[3] = dt_datos_nuevos.Rows[indice][3].ToString();//primer apellido
                    valores[4] = dt_datos_nuevos.Rows[indice][4].ToString();//segundo apellido
                    valores[5] = dt_datos_nuevos.Rows[indice][5].ToString();//email


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

                    calificacion = calificaciones(dt_datos_nuevos.Rows[indice][5].ToString());
                    valores[21] = calificacion;//calificacion


                    if (calificacion >= min_apro)//certificado
                    {
                        valores[22] = "Certificado";
                    }
                    else
                    {
                        valores[22] = "No Aprobado";
                    }

                    for (int j = 0; j < dataGridView3.Rows.Count; j++)
                    {
                        if (email == dataGridView3[5, j].Value.ToString())
                        {
                            repetido++;
                        }
                    }

                    if (repetido == 0)
                    {
                        dataGridView3.Rows.Add(valores);
                    }
                    else
                    {
                        tot_repetidos2++;
                        col_cont--;
                    }
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

            dt_datos_base.Rows.Clear();
            dt_datos_base.Columns.Clear();
            dt_datos_base = exportar_dgv(1);

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
            textBox1.Text = seleccionar_archivo("csv")[0];
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
                textBox2.Text = seleccionar_archivo("xls")[0];
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
                textBox2.Text = seleccionar_archivo("csv")[0];
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
                textBox3.Text = seleccionar_archivo("csv")[0];
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
                string[] rutas;
                int ulti_diag = 0;
                rutas = seleccionar_archivo("pdf");

                comboBox2.Items.Clear();

                for (int i = 0; i < rutas.Length; i++)
                {
                    ulti_diag = rutas[i].LastIndexOf("\\");
                    comboBox2.Items.Add(rutas[i].Substring(ulti_diag+1, (rutas[i].Length - (ulti_diag + 1))));
                }

                if (comboBox2.Items.Count>0)
                {
                    textBox3.Text = comboBox2.Items[0].ToString();
                    comboBox2.SelectedIndex = 0;
                }

                
                if (textBox3.Text.Length > 0)
                {
                    dt_actividades.Rows.Clear();
                    dt_actividades.Columns.Clear();
                    dt_actividades = cargar_pdf(rutas);
                    //dataGridView2.DataSource= cargar_pdf(rutas);
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            tot_repetidos = 0;
            tot_repetidos2 = 0;
            procesar();
            tabControl1.SelectedIndex = 2;
            dataGridView2.Focus();

            if (tot_repetidos2==0) 
            {
                MessageBox.Show("El Análisis ha terminado Correctamente.\nSe quitaron: " + tot_repetidos + " registros repetidos.", "Analisis Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label5.Text = "Repetidos Descartados: " + tot_repetidos;
            }
            else
            {
                MessageBox.Show("El Análisis ha terminado Correctamente.\nSe quitaron: " + tot_repetidos + " registros repetidos.\nSe ignoraron: " + tot_repetidos2 + " registros duplicados.", "Analisis Terminado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label5.Text = "Repetidos Descartados: " + tot_repetidos+" | Repetidos Ignorados: "+tot_repetidos2 ;
            }

            
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
            comboBox2.Visible = false;
            textBox3.Visible = true;
            label4.Text = "Archivo XLSX  Reporte de Calificaciones";
            groupBox1.BackColor = System.Drawing.Color.RosyBrown;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            label6.Text = "Archivo PDF Reporte de Actividades";
            comboBox2.Visible = true;
            textBox3.Visible = false;
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
