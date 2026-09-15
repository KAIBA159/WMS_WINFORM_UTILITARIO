using ExcelDataReader;
using System.Collections.Concurrent;
using System.Data;
using System.IO;
using System.Reflection.Emit;
using System.Text;
using System.Text.Json;

namespace WinFormsApp1
{
    public partial class lblBaseDatos : Form
    {

        // MUEVE ESTO AQUÍ ARRIBA (Nivel de clase)
        private int itemsProcesados = 0;
        private int transferenciasExitosas = 0;
        private int errores = 0;
        private int totalLineasExitosas = 0;
        private int totalLineasFallidas = 0;

        //string miBaseDatos = "SBO_MAKITA_PE"; // La BD de la sociedad
        //string miUsuario = "manager";

        // 1. Creamos el manejador que le dice a .NET que acepte el certificado de HANA
        private static readonly HttpClientHandler handler = new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (message, cert, chain, sslPolicyErrors) => true
        };


        //// 2. Instanciamos el cliente pasándole nuestro manejador modificado
        //private static readonly HttpClient client = new HttpClient(handler)
        //{
        //    Timeout = TimeSpan.FromSeconds(300) // <-- SOLUCIÓN DEFINITIVA AL TIMEOUT
        //};

        private static readonly HttpClient client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(300) // <-- SOLUCIÓN DEFINITIVA AL TIMEOUT
        };


        // ==========================================
        // VARIABLES GLOBALES PARA FASE 4 (Sinceramiento de Series)
        // ==========================================
        private DataTable dtExcelFase4Completo;


        private string sessionToken = "";

        // ==========================================
        // VARIABLES GLOBALES DE CONEXIÓN ODBC (HANA)
        // ==========================================
        private readonly string dbServerGlobal = "192.168.1.17:30015";
        private readonly string dbUserGlobal = "SAPINST";
        private readonly string dbPassGlobal = "Ki$ta66mpe";
        private readonly string dbSchemaGlobal = "SBO_MAKITA_20260717";
        //private readonly string dbSchemaGlobal = "SBO_MAKITA_PE";


        // Cadena de conexión pre-ensamblada lista para usarse
        private string GetHanaConnectionString()
        {
            return $"Driver={{HDBODBC}};ServerNode={dbServerGlobal};UID={dbUserGlobal};PWD={dbPassGlobal};CS={dbSchemaGlobal};";
        }


        // ==========================================
        // VARIABLES GLOBALES DE CONEXIÓN SERVICE LAYER (SAP B1)
        // ==========================================
        //private readonly string slBaseDatosGlobal = "SBO_MAKITA_PE";
        private readonly string slBaseDatosGlobal = "SBO_MAKITA_20260717";
        private readonly string slUsuarioGlobal = "manager";
        private readonly string slClaveGlobal = "m1r1";
        private readonly string slUrlBaseGlobal = "https://192.168.1.17:50000/b1s/v1/";



        public lblBaseDatos()
        {
            InitializeComponent();
        }


        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
                button1.Enabled = false;
                button1.Text = "Conectando...";

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (senderCert, cert, chain, sslPolicyErrors) => true;

                // Configuramos la URL base usando la variable global
                if (client.BaseAddress == null)
                {
                    client.BaseAddress = new Uri(slUrlBaseGlobal);
                }

                // Armamos el JSON inyectando las variables globales directamente
                string loginJson = $@"{{
            ""CompanyDB"": ""{slBaseDatosGlobal}"",
            ""UserName"": ""{slUsuarioGlobal}"",
            ""Password"": ""{slClaveGlobal}""
        }}";

                var content = new StringContent(loginJson, System.Text.Encoding.UTF8, "application/json");

                // Ejecutamos el POST al endpoint de Login
                HttpResponseMessage response = await client.PostAsync("Login", content);

                if (response.IsSuccessStatusCode)
                {
                    // Extraer las cookies (donde viene el B1SESSION)
                    var cookies = response.Headers.GetValues("Set-Cookie");
                    foreach (var cookie in cookies)
                    {
                        if (cookie.StartsWith("B1SESSION"))
                        {
                            sessionToken = cookie.Split(';')[0];
                            client.DefaultRequestHeaders.Add("Cookie", sessionToken);
                            break;
                        }
                    }

                    textBox1.Text = "Conexión exitosa - " + sessionToken;

                    // ACTUALIZAMOS TUS LABELS CON LAS GLOBALES
                    lblbdcon.Text = $"Base Datos: {slBaseDatosGlobal}";
                    lblUsuario.Text = $"Usuario     : {slUsuarioGlobal}";

                    MessageBox.Show("Token generado correctamente. Ya podemos interactuar con la BD.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    MessageBox.Show("Error de credenciales o SL: " + error, "Error de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    lblbdcon.Text = "BD: Desconectado";
                    lblUsuario.Text = "Usuario: -";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de red/timeout: " + ex.Message, "Fallo Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                button1.Text = "Generar Token";
            }
        }




        private async void button2_Click(object sender, EventArgs e)
        {
            try
            {
                button2.Enabled = false;
                button2.Text = "Consultando HANA, espere...";

                label1.Text = "Se encontró : Buscando...";

                // 1. Configuramos las columnas del DataGridView 1 
                dataGridView1.Columns.Clear();
                dataGridView1.Columns.Add("Nro", "#");

                // ¡AGREGAMOS EL ABSENTRY (Fundamental para el Botón 4)!
                dataGridView1.Columns.Add("AbsEntry", "ID Interno");

                dataGridView1.Columns.Add("BinCode", "Código Ubicación");
                dataGridView1.Columns.Add("WhsCode", "Almacén");
                dataGridView1.Columns.Add("Sub1", "Fila");
                dataGridView1.Columns.Add("Sub2", "Cuerpo");
                dataGridView1.Columns.Add("Sub3", "Piso");
                dataGridView1.Columns.Add("Sub4", "Fondo");
                dataGridView1.Columns.Add("Inactive", "Inactivo");

                // Ajustamos diseño visual
                dataGridView1.Columns["Nro"].Width = 40;

                // ¡LA MAGIA AQUÍ! Ocultamos la columna para no ensuciar tu diseño visual en pantalla
                dataGridView1.Columns["AbsEntry"].Visible = false;

                // 2. Endpoint actualizado: Pedimos 'AbsEntry' en el $select a la API
                string nextLink = "BinLocations?$select=AbsEntry,BinCode,Warehouse,Sublevel1,Sublevel2,Sublevel3,Sublevel4,Inactive";
                int totalRecords = 0;

                // 3. Ciclo de paginación OData (recorrerá todas las ubicaciones)
                while (!string.IsNullOrEmpty(nextLink))
                {
                    HttpResponseMessage response = await client.GetAsync(nextLink);
                    response.EnsureSuccessStatusCode();

                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                    {
                        JsonElement root = doc.RootElement;
                        JsonElement values = root.GetProperty("value");

                        foreach (JsonElement item in values.EnumerateArray())
                        {
                            // Capturamos el AbsEntry y todos los campos nativos
                            string absEntry = item.GetProperty("AbsEntry").ToString();
                            string binCode = item.GetProperty("BinCode").ToString();
                            string whsCode = item.GetProperty("Warehouse").ToString();
                            string fila = item.GetProperty("Sublevel1").ToString();
                            string cuerpo = item.GetProperty("Sublevel2").ToString();
                            string piso = item.GetProperty("Sublevel3").ToString();
                            string fondo = item.GetProperty("Sublevel4").ToString();
                            string inactive = item.GetProperty("Inactive").ToString();

                            totalRecords++;

                            // Llenamos la fila incluyendo el absEntry (que estará oculto)
                            dataGridView1.Rows.Add(totalRecords, absEntry, binCode, whsCode, fila, cuerpo, piso, fondo, inactive);
                        }

                        // 4. Verificamos si HANA nos manda una siguiente "página" de resultados
                        if (root.TryGetProperty("odata.nextLink", out JsonElement nextLinkElement))
                        {
                            string link = nextLinkElement.GetString();
                            nextLink = link.Contains("/b1s/v1/") ? link.Substring(link.IndexOf("/b1s/v1/") + 8) : link;
                        }
                        else
                        {
                            nextLink = null; // Terminamos de leer toda la base
                        }
                    }
                }

                // 5. Actualizamos el Label con el total exacto de registros encontrados
                label1.Text = $"Se encontró : {totalRecords}";

                MessageBox.Show($"¡Éxito! Se cargaron {totalRecords} ubicaciones actuales desde la base de datos.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al consultar ubicaciones: " + ex.Message, "Error Service Layer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                label1.Text = "Se encontró : Error";
            }
            finally
            {
                button2.Enabled = true;
                button2.Text = "Consulta Ubicaciones en BD";
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Seleccione el archivo DATAMASTER";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    button3.Enabled = false;
                    button3.Text = "Cargando Excel...";
                    label2.Text = "Total Excel : Cargando...";

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            DataTable dtExcel = result.Tables["DATAMASTER"];

                            if (dtExcel != null)
                            {
                                int totalOriginal = dtExcel.Rows.Count;

                                // Validamos y filtramos registros únicos basados en las columnas B, C, D, E y F (índices 1 al 5)
                                var filasUnicas = dtExcel.AsEnumerable()
                                    .GroupBy(row => new
                                    {
                                        Almacen = row[1]?.ToString().Trim(),
                                        Fila = row[2]?.ToString().Trim(),
                                        Cuerpo = row[3]?.ToString().Trim(),
                                        Piso = row[4]?.ToString().Trim(),
                                        Fondo = row[5]?.ToString().Trim()
                                    })
                                    .Select(grupo => grupo.First())
                                    .ToList();

                                int totalUnicos = filasUnicas.Count;
                                int totalDuplicados = totalOriginal - totalUnicos;

                                // 1. DGV2 estructurado idénticamente a DGV1 (sin campos extra)
                                dataGridView2.Columns.Clear();
                                dataGridView2.Columns.Add("Nro", "#");
                                dataGridView2.Columns.Add("BinCodeRef", "Ubicación Referencial");
                                dataGridView2.Columns.Add("Almacen", "Almacén");
                                dataGridView2.Columns.Add("Sub1", "Fila");
                                dataGridView2.Columns.Add("Sub2", "Cuerpo");
                                dataGridView2.Columns.Add("Sub3", "Piso");
                                dataGridView2.Columns.Add("Sub4", "Fondo");
                                dataGridView2.Columns.Add("Inactive", "Inactivo");

                                dataGridView2.Columns["Nro"].Width = 40;

                                // 2. Volcamos EXCLUSIVAMENTE los registros únicos a la grilla
                                int rowCount = 0;
                                foreach (var row in filasUnicas)
                                {
                                    rowCount++;

                                    string binRef = row[0].ToString();
                                    string whs = row[1].ToString();
                                    string fila = row[2].ToString();
                                    string cuerpo = row[3].ToString();
                                    string piso = row[4].ToString();
                                    string fondo = row[5].ToString();

                                    dataGridView2.Rows.Add(rowCount, binRef, whs, fila, cuerpo, piso, fondo, "tNO");
                                }

                                label2.Text = $"Total Excel : {totalUnicos}";

                                // 3. Notificación con el desglose exacto de duplicados omitidos y registros únicos
                                MessageBox.Show(
                                    $"¡Importación procesada con éxito!\n\n" +
                                    $"• Total de filas leídas en Excel: {totalOriginal}\n" +
                                    $"• Registros duplicados omitidos: {totalDuplicados}\n" +
                                    $"• Registros únicos cargados: {totalUnicos}",
                                    "WMS Makita - Control de Duplicados",
                                    MessageBoxButtons.OK,
                                    totalDuplicados > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information
                                );
                            }
                            else
                            {
                                MessageBox.Show("No se encontró una hoja llamada 'DATAMASTER' en el Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                label2.Text = "Total Excel : Error";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Fallo de Importación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    label2.Text = "Total Excel : Error";
                }
                finally
                {
                    button3.Enabled = true;
                    button3.Text = "Importar Ubicaciones Finales/Reales";
                }
            }
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            button4.Text = "Procesando HANA...";
            dataGridView1.Enabled = false;
            dataGridView2.Enabled = false;

            // Contadores Thread-Safe
            int creadas = 0;
            int actualizadas = 0;
            int errores = 0;
            int filasProcesadas = 0;
            int inactivadasFantasmas = 0; // NUEVO CONTADOR PARA ESCENARIO 3

            // Colección segura para el log de errores
            System.Collections.Concurrent.ConcurrentBag<string> logErrores = new System.Collections.Concurrent.ConcurrentBag<string>();

            // 1. Calculamos el total real a procesar (AHORA ES ILIMITADO, TODO EL EXCEL)
            int totalAProcesar = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (!row.IsNewRow) totalAProcesar++;
            }

            // Motor de reporte de progreso seguro para la UI
            var progress = new Progress<int>(procesadas =>
            {
                // Validación de seguridad para evitar división por cero
                int porcentaje = totalAProcesar > 0 ? (int)((double)procesadas / totalAProcesar * 100) : 0;
                if (porcentaje > 100) porcentaje = 100;

                progressBar1.Value = porcentaje;
                lblProgreso.Text = $"Procesando: {procesadas} de {totalAProcesar} ({porcentaje}%)";
            });

            // 2. Cargamos HANA a la Memoria RAM (Cruce ultra rápido)
            var ubicacionesHANA = new Dictionary<string, string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;
                string binCode = row.Cells["BinCode"].Value?.ToString();
                string absEntry = row.Cells["AbsEntry"].Value?.ToString();

                if (!string.IsNullOrEmpty(binCode))
                {
                    ubicacionesHANA[binCode] = absEntry;
                }
            }

            var tareas = new List<Task>();

            // 3. Concurrencia controlada a 15 hilos hacia Service Layer
            using (SemaphoreSlim semaphore = new SemaphoreSlim(15))
            {
                foreach (DataGridViewRow rowExcel in dataGridView2.Rows)
                {
                    if (rowExcel.IsNewRow) continue;

                    string whs = rowExcel.Cells["Almacen"].Value?.ToString() ?? "";
                    string fila = rowExcel.Cells["Sub1"].Value?.ToString() ?? "";
                    string cuerpo = rowExcel.Cells["Sub2"].Value?.ToString() ?? "";
                    string piso = rowExcel.Cells["Sub3"].Value?.ToString() ?? "";
                    string fondo = rowExcel.Cells["Sub4"].Value?.ToString() ?? "";
                    string inactive = rowExcel.Cells["Inactive"].Value?.ToString() ?? "tNO";

                    // ¡LA CLAVE! Ignoramos la columna con guiones y concatenamos directamente los niveles
                    string cleanBinCode = $"{whs}{fila}{cuerpo}{piso}{fondo}";


                    // ==========================================
                    // AGREGAR ESTA VALIDACIÓN AQUÍ:
                    // ==========================================
                    if (cleanBinCode.Contains("UBICACIÓN-DE-SISTEMA") || cleanBinCode.Contains("SYSTEM"))
                    {
                        continue; // Salta las ubicaciones de sistema para evitar el error -5002
                    }

                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            if (ubicacionesHANA.ContainsKey(cleanBinCode))
                            {
                                // ACTUALIZAR (PATCH)
                                string entry = ubicacionesHANA[cleanBinCode];
                                string jsonPatch = $@"{{ ""Inactive"": ""{inactive}"" }}";
                                var content = new StringContent(jsonPatch, System.Text.Encoding.UTF8, "application/json");

                                var response = await client.PatchAsync($"BinLocations({entry})", content);

                                // Capturamos el error real de SAP B1 si falla
                                if (!response.IsSuccessStatusCode)
                                {
                                    string sapError = await response.Content.ReadAsStringAsync();
                                    throw new Exception($"HTTP {(int)response.StatusCode} - Detalles SAP: {sapError}");
                                }

                                System.Threading.Interlocked.Increment(ref actualizadas);
                            }
                            else
                            {
                                // CREAR (POST)
                                string jsonPost = $@"{{
                            ""Warehouse"": ""{whs}"",
                            ""Sublevel1"": ""{fila}"",
                            ""Sublevel2"": ""{cuerpo}"",
                            ""Sublevel3"": ""{piso}"",
                            ""Sublevel4"": ""{fondo}"",
                            ""Inactive"": ""{inactive}""
                        }}";

                                var content = new StringContent(jsonPost, System.Text.Encoding.UTF8, "application/json");
                                var response = await client.PostAsync("BinLocations", content);

                                // Capturamos el error real de SAP B1 si falla
                                if (!response.IsSuccessStatusCode)
                                {
                                    string sapError = await response.Content.ReadAsStringAsync();
                                    throw new Exception($"HTTP {(int)response.StatusCode} - Detalles SAP: {sapError}");
                                }

                                System.Threading.Interlocked.Increment(ref creadas);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Threading.Interlocked.Increment(ref errores);

                            string tipoOperacion = ubicacionesHANA.ContainsKey(cleanBinCode) ? "ACTUALIZAR" : "CREAR";
                            logErrores.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL {tipoOperacion} | Ubicación: {cleanBinCode} | Motivo: {ex.Message}");
                        }
                        finally
                        {
                            int procesadasLocales = System.Threading.Interlocked.Increment(ref filasProcesadas);
                            ((IProgress<int>)progress).Report(procesadasLocales);

                            semaphore.Release();
                        }
                    }));
                }

                // Esperamos a que los miles de hilos terminen su trabajo
                await Task.WhenAll(tareas);
            }

            // ====================================================================
            // INICIO ESCENARIO 3: BARRIDO DE INACTIVACIÓN (Limpieza de fantasmas)
            // ====================================================================
            lblProgreso.Text = "Fase 2: Buscando y bloqueando ubicaciones sobrantes en HANA...";

            // Guardamos todos los códigos del Excel en un HashSet súper rápido
            HashSet<string> codigosExcel = new HashSet<string>();
            foreach (DataGridViewRow rowExcel in dataGridView2.Rows)
            {
                if (rowExcel.IsNewRow) continue;
                string whs = rowExcel.Cells["Almacen"].Value?.ToString() ?? "";
                string fila = rowExcel.Cells["Sub1"].Value?.ToString() ?? "";
                string cuerpo = rowExcel.Cells["Sub2"].Value?.ToString() ?? "";
                string piso = rowExcel.Cells["Sub3"].Value?.ToString() ?? "";
                string fondo = rowExcel.Cells["Sub4"].Value?.ToString() ?? "";

                codigosExcel.Add($"{whs}{fila}{cuerpo}{piso}{fondo}");
            }

            var tareasLimpieza = new List<Task>();

            using (SemaphoreSlim semaforoLimpieza = new SemaphoreSlim(15))
            {
                foreach (DataGridViewRow rowHana in dataGridView1.Rows)
                {
                    if (rowHana.IsNewRow) continue;

                    string binCodeHana = rowHana.Cells["BinCode"].Value?.ToString() ?? "";
                    string absEntryHana = rowHana.Cells["AbsEntry"].Value?.ToString() ?? "";
                    string statusActual = rowHana.Cells["Inactive"].Value?.ToString() ?? "tNO";

                    // ==========================================
                    // AGREGAR ESTA VALIDACIÓN AQUÍ:
                    // ==========================================
                    if (binCodeHana.Contains("UBICACIÓN-DE-SISTEMA") || binCodeHana.Contains("SYS"))
                    {
                        continue; // Evita que el sistema intente inactivar el pulmón o ubicaciones protegidas
                    }



                    // REGLA: Si está en HANA, pero NO está en Excel, y está ACTIVA -> ¡Inactivar!
                    if (!string.IsNullOrEmpty(binCodeHana) && !codigosExcel.Contains(binCodeHana) && statusActual == "tNO")
                    {
                        await semaforoLimpieza.WaitAsync();

                        tareasLimpieza.Add(Task.Run(async () =>
                        {
                            try
                            {
                                string jsonPatch = $@"{{ ""Inactive"": ""tYES"" }}";
                                var content = new StringContent(jsonPatch, System.Text.Encoding.UTF8, "application/json");

                                var response = await client.PatchAsync($"BinLocations({absEntryHana})", content);

                                if (response.IsSuccessStatusCode)
                                {
                                    System.Threading.Interlocked.Increment(ref inactivadasFantasmas);
                                }
                                else
                                {
                                    string sapError = await response.Content.ReadAsStringAsync();
                                    logErrores.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL INACTIVAR SOBRANTE | Ubicación: {binCodeHana} | Motivo: {sapError}");
                                    System.Threading.Interlocked.Increment(ref errores);
                                }
                            }
                            catch (Exception ex)
                            {
                                logErrores.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL INACTIVAR SOBRANTE | Ubicación: {binCodeHana} | Motivo: {ex.Message}");
                                System.Threading.Interlocked.Increment(ref errores);
                            }
                            finally
                            {
                                semaforoLimpieza.Release();
                            }
                        }));
                    }
                }
                await Task.WhenAll(tareasLimpieza);
            }
            // ====================================================================
            // FIN ESCENARIO 3
            // ====================================================================

            // 4. GENERACIÓN DEL ARCHIVO FÍSICO DE LOG (Si hubo errores)
            if (errores > 0)
            {
                try
                {
                    string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string nombreArchivo = $"Log_WMS_Makita_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                    string rutaCompleta = System.IO.Path.Combine(rutaEscritorio, nombreArchivo);

                    System.IO.File.WriteAllLines(rutaCompleta, logErrores);

                    MessageBox.Show($"Migración masiva finalizada con observaciones.\n\nCreadas: {creadas}\nActualizadas: {actualizadas}\nInactivadas (Sobrantes): {inactivadasFantasmas}\nErrores: {errores}\n\nRevisa el archivo de Log generado.",
                                    "WMS Makita - Log Generado", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    // Abre el archivo .txt de inmediato en el Bloc de Notas
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = rutaCompleta,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("No se pudo guardar/abrir el archivo Log: " + ex.Message, "Error de I/O", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show($"¡Lote MASIVO procesado limpiamente!\n\nCreadas: {creadas}\nActualizadas: {actualizadas}\nInactivadas (Sobrantes): {inactivadasFantasmas}\nErrores: 0",
                                "WMS Makita - Éxito Total", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // Reactivamos la UI
            dataGridView1.Enabled = true;
            dataGridView2.Enabled = true;
            button4.Enabled = true;
            button4.Text = "Actualizar/Crear en BD";
            lblProgreso.Text = "Proceso Completado.";

        }


        private async Task<int> EnviarSubnivelesASap(HashSet<string> lista, int nivelB1, string prefijo)
        {
            int creados = 0;

            foreach (string codigo in lista)
            {
                string jsonPost = $@"{{
            ""WarehouseSublevel"": {nivelB1},
            ""Code"": ""{codigo}"",
            ""Description"": ""{prefijo} {codigo}""
        }}";

                var content = new StringContent(jsonPost, System.Text.Encoding.UTF8, "application/json");
                var response = await client.PostAsync("WarehouseSublevelCodes", content);

                // Si SAP lo crea exitosamente (201 Created), sumamos al contador.
                // Si SAP da error (Ej: 400 porque ya existe), lo ignora y sigue.
                if (response.IsSuccessStatusCode)
                {
                    creados++;
                }
            }

            return creados;
        }

        private async void button7_Click(object sender, EventArgs e)
        {

            button7.Enabled = false;
            button7.Text = "Consultando HANA (Calculando Series)...";

            try
            {
                string connectionString = GetHanaConnectionString();

                string queryHana = @"
    SELECT 
        T0.""BinCode"" AS ""Ubicacion_Origen"",
        T0.""AbsEntry"" AS ""Id_Ubicacion_Origen"",
        T1.""ItemCode"" AS ""Codigo_Articulo"",
        T3.""ItemName"" AS ""Descripcion"",
        CASE WHEN T3.""ManSerNum"" = 'Y' THEN 1 ELSE T1.""OnHandQty"" END AS ""Cantidad_A_Mover"",
        T3.""ManSerNum"" AS ""Maneja_Series"",
        T4.""SysNumber"" AS ""Id_Serie_Interno"",
        T4.""DistNumber"" AS ""Numero_Serie""
    FROM OBIN T0
    INNER JOIN OIBQ T1 ON T0.""AbsEntry"" = T1.""BinAbs""
    INNER JOIN OITM T3 ON T1.""ItemCode"" = T3.""ItemCode""
    LEFT JOIN OSBQ T2 ON T1.""ItemCode"" = T2.""ItemCode"" AND T1.""BinAbs"" = T2.""BinAbs"" AND T2.""OnHandQty"" > 0
    LEFT JOIN OSRN T4 ON T2.""SnBMDAbs"" = T4.""AbsEntry""
    WHERE T0.""WhsCode"" = 'ALM01' 
      AND T1.""OnHandQty"" > 0 
      AND T0.""BinCode"" <> 'ALM01UBICACIÓN-DE-SISTEMA'
    ORDER BY T0.""BinCode"", T1.""ItemCode"";";

                DataTable dtStock = await Task.Run(() =>
                {
                    DataTable dt = new DataTable();
                    using (System.Data.Odbc.OdbcConnection conn = new System.Data.Odbc.OdbcConnection(connectionString))
                    {
                        conn.Open();
                        using (System.Data.Odbc.OdbcCommand cmd = new System.Data.Odbc.OdbcCommand(queryHana, conn))
                        {
                            cmd.CommandTimeout = 120;
                            using (System.Data.Odbc.OdbcDataAdapter da = new System.Data.Odbc.OdbcDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                    return dt;
                });

                dgv3.DataSource = dtStock;

                MessageBox.Show($"Radiografía del almacén completada con éxito.\n\nSe extrajeron {dtStock.Rows.Count} registros listos para reubicar.",
                                "WMS Makita - Consulta Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error al conectar con HANA por ODBC:\n\n{ex.Message}\n\nPor favor, verifica la configuración global o si tienes instalado el driver HDBODBC.",
                                "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button7.Enabled = true;
                button7.Text = "Consultar UbicacionStock";
            }


        }


        private async void button8_Click(object sender, EventArgs e)
        {
            // Reseteamos los contadores a nivel de clase
            itemsProcesados = 0;
            transferenciasExitosas = 0;
            errores = 0;
            totalLineasExitosas = 0;
            totalLineasFallidas = 0;

            int idUbicacionPulmon = 3; // ID confirmado de tu ubicación pulmón

            DataTable dtStock = dgv3.DataSource as DataTable;
            if (dtStock == null || dtStock.Rows.Count == 0)
            {
                MessageBox.Show("Primero debes consultar el stock en la grilla (Botón 7).", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            button8.Enabled = false;
            button8.Text = "Trasladando Stock...";
            button7.Enabled = false;

            System.Collections.Concurrent.ConcurrentBag<string> logReubicacion = new System.Collections.Concurrent.ConcurrentBag<string>();

            var agrupadoPorUbicacion = dtStock.AsEnumerable()
                .GroupBy(row => row["Id_Ubicacion_Origen"].ToString())
                .ToList();

            progressBar3.Minimum = 0;
            progressBar3.Maximum = agrupadoPorUbicacion.Count;
            progressBar3.Value = 0;
            progressBar3.Visible = true;

            // Procesamiento con SemaphoreSlim(1) para orden estricto en SAP
            using (SemaphoreSlim semaphore = new SemaphoreSlim(1))
            {
                var tareas = new List<Task>();

                foreach (var grupoUbicacion in agrupadoPorUbicacion)
                {
                    int idOrigen = Convert.ToInt32(grupoUbicacion.Key);
                    if (idOrigen == idUbicacionPulmon) continue;

                    string binCodeOrigen = grupoUbicacion.First()["Ubicacion_Origen"].ToString();
                    var localGrupoUbicacion = grupoUbicacion;

                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            List<object> lineasTransferencia = new List<object>();
                            var agrupadoPorArticulo = localGrupoUbicacion.GroupBy(r => r["Codigo_Articulo"].ToString());

                            foreach (var grupoArticulo in agrupadoPorArticulo)
                            {
                                string itemCode = grupoArticulo.Key;
                                string manejaSeries = grupoArticulo.First()["Maneja_Series"].ToString();

                                double cantidadTotalItem = 0;
                                var listaSeries = new List<object>();
                                var allocs = new List<object>();

                                if (manejaSeries == "Y")
                                {
                                    int serialIndex = 0;
                                    foreach (var fila in grupoArticulo)
                                    {
                                        double qty = Convert.ToDouble(fila["Cantidad_A_Mover"]);
                                        cantidadTotalItem += qty;
                                        listaSeries.Add(new { SystemSerialNumber = Convert.ToInt32(fila["Id_Serie_Interno"]), Quantity = qty });
                                        allocs.Add(new { BinAbsEntry = idOrigen, Quantity = qty, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = serialIndex, BinActionType = "batFromWarehouse" });
                                        allocs.Add(new { BinAbsEntry = idUbicacionPulmon, Quantity = qty, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = serialIndex, BinActionType = "batToWarehouse" });
                                        serialIndex++;
                                    }
                                    lineasTransferencia.Add(new { ItemCode = itemCode, Quantity = cantidadTotalItem, WarehouseCode = "ALM01", FromWarehouseCode = "ALM01", SerialNumbers = listaSeries, StockTransferLinesBinAllocations = allocs });
                                }
                                else
                                {
                                    foreach (var fila in grupoArticulo) { cantidadTotalItem += Convert.ToDouble(fila["Cantidad_A_Mover"]); }
                                    allocs.Add(new { BinAbsEntry = idOrigen, Quantity = cantidadTotalItem, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = 0, BinActionType = "batFromWarehouse" });
                                    allocs.Add(new { BinAbsEntry = idUbicacionPulmon, Quantity = cantidadTotalItem, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = 0, BinActionType = "batToWarehouse" });
                                    lineasTransferencia.Add(new { ItemCode = itemCode, Quantity = cantidadTotalItem, WarehouseCode = "ALM01", FromWarehouseCode = "ALM01", StockTransferLinesBinAllocations = allocs });
                                }
                            }

                            // =========================================================================
                            // CLAVE: PARTIR EN BLOQUES DE 20 LÍNEAS MÁXIMO POR CADA DOCUMENTO DE SAP
                            // =========================================================================
                            if (lineasTransferencia.Count > 0)
                            {
                                const int TAMANO_LOTE = 20; // Ajustado a 20 líneas para máxima fluidez
                                for (int i = 0; i < lineasTransferencia.Count; i += TAMANO_LOTE)
                                {
                                    var chunkLineas = lineasTransferencia.Skip(i).Take(TAMANO_LOTE).ToList();
                                    int cantLineas = chunkLineas.Count;

                                    bool exito = await EnviarTransferencia(chunkLineas, binCodeOrigen, logReubicacion);

                                    if (exito)
                                    {
                                        System.Threading.Interlocked.Add(ref totalLineasExitosas, cantLineas);
                                        System.Threading.Interlocked.Increment(ref transferenciasExitosas);
                                    }
                                    else
                                    {
                                        System.Threading.Interlocked.Add(ref totalLineasFallidas, cantLineas);
                                        System.Threading.Interlocked.Increment(ref errores);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Threading.Interlocked.Increment(ref errores);
                            logReubicacion.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN UBICACIÓN {binCodeOrigen} | {ex.Message}");
                        }
                        finally
                        {
                            System.Threading.Interlocked.Increment(ref itemsProcesados);
                            Invoke(new Action(() =>
                            {
                                if (progressBar3.Value < progressBar3.Maximum)
                                    progressBar3.Value++;
                            }));
                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tareas);
            }

            progressBar3.Value = progressBar3.Maximum;
            button8.Enabled = true;
            button7.Enabled = true;

            if (errores > 0)
            {
                string ruta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_WMS_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(ruta, logReubicacion);
                MessageBox.Show($"Migración finalizada con observaciones.\n\nDocumentos Exitosos: {transferenciasExitosas} | Líneas Trasladadas: {totalLineasExitosas}\nDocumentos con Error: {errores} | Líneas Fallidas: {totalLineasFallidas}\n\nRevisa el archivo de log en el escritorio.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show($"¡ÉXITO TOTAL!\n\nSe trasladaron todas las líneas de stock exitosamente al pulmón.\nTotal de líneas procesadas: {totalLineasExitosas}", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        // Función auxiliar para mantener el código limpio

        private async Task<bool> EnviarTransferencia(List<object> lineas, string ubicacion, System.Collections.Concurrent.ConcurrentBag<string> log)
        {
            try
            {
                var payload = new { U_EXX_TIPOOPER = "11", U_EXX_MOTIVTRA = "13", StockTransferLines = lineas };
                string json = JsonSerializer.Serialize(payload);
                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                var response = await client.PostAsync("StockTransfers", content);

                if (response.IsSuccessStatusCode)
                {
                    return true; // Éxito
                }
                else
                {
                    string err = await response.Content.ReadAsStringAsync();
                    log.Add($"[{DateTime.Now:HH:mm:ss}] ERROR EN {ubicacion} | {err}");
                    return false; // Error en SAP
                }
            }
            catch (Exception ex)
            {
                log.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN {ubicacion} | {ex.Message}");
                return false; // Error de red
            }
        }


        private void button9_Click(object sender, EventArgs e)
        {
            // 1. Reiniciar la barra de progreso y etiquetas
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 0;
            progressBar1.Visible = false;
            lblProgreso.Text = "Sistema reiniciado. Listo para operar.";

            // 2. Limpiar las grillas de datos
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            if (dgv3 != null)
            {
                dgv3.DataSource = null;
                dgv3.Rows.Clear();
                dgv3.Columns.Clear();
            }

            // 3. Reactivar controles principales
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;

            MessageBox.Show("Se ha limpiado la interfaz y reiniciado los contadores correctamente.",
                            "WMS Makita - Reset", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private async void button10_Click(object sender, EventArgs e)
        {
            // 1. Validamos que la grilla tenga datos cargados
            DataTable dtStock = dgv3.DataSource as DataTable;
            if (dtStock == null || dtStock.Rows.Count == 0)
            {
                MessageBox.Show("Primero debes consultar el stock en la grilla (Botón 7).", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Extraemos los códigos de artículo únicos de la grilla
            var codigosArticulos = dtStock.AsEnumerable()
                .Select(r => r["Codigo_Articulo"].ToString())
                .Distinct()
                .ToList();

            button10.Enabled = false;
            button10.Text = "Activando Masivamente...";

            // Configuración de la Barra de Progreso 2
            progressBar2.Minimum = 0;
            progressBar2.Maximum = codigosArticulos.Count;
            progressBar2.Value = 0;
            progressBar2.Visible = true;

            int actualizados = 0;
            int errores = 0;
            int procesados = 0;
            var logActivacion = new System.Collections.Concurrent.ConcurrentBag<string>();

            var tareas = new List<Task>();

            // 3. Concurrencia controlada a 15 hilos
            using (SemaphoreSlim semaphore = new SemaphoreSlim(15))
            {
                foreach (var itemCode in codigosArticulos)
                {
                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            string endpoint = $"Items('{itemCode}')";

                            // CORRECCIÓN CRÍTICA: Eliminamos ValidTo para evitar validaciones de SAP
                            // con ubicaciones por defecto inactivas. Solo enviamos el estado 'Activo'.
                            string jsonPatch = @"{
                        ""Valid"": ""tYES""
                    }";

                            var content = new StringContent(jsonPatch, Encoding.UTF8, "application/json");
                            var request = new HttpRequestMessage(new HttpMethod("PATCH"), endpoint) { Content = content };

                            var response = await client.SendAsync(request);

                            if (response.IsSuccessStatusCode)
                            {
                                System.Threading.Interlocked.Increment(ref actualizados);
                            }
                            else
                            {
                                string err = await response.Content.ReadAsStringAsync();
                                logActivacion.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL ACTIVAR {itemCode} | {err}");
                                System.Threading.Interlocked.Increment(ref errores);
                            }
                        }
                        catch (Exception ex)
                        {
                            logActivacion.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN {itemCode} | {ex.Message}");
                            System.Threading.Interlocked.Increment(ref errores);
                        }
                        finally
                        {
                            int current = System.Threading.Interlocked.Increment(ref procesados);

                            // Actualización segura de la barra de progreso
                            Invoke(new Action(() =>
                            {
                                progressBar2.Value = current > progressBar2.Maximum ? progressBar2.Maximum : current;
                            }));

                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tareas);
            }

            progressBar2.Value = progressBar2.Maximum;

            // Generar log si hubo errores
            if (errores > 0)
            {
                string ruta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_Activacion_WMS_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(ruta, logActivacion);
            }

            button10.Enabled = true;
            button10.Text = "Activar Artículos Masivamente";

            if (errores > 0)
            {
                MessageBox.Show($"¡Proceso finalizado con observaciones!\n\nArtículos activados: {actualizados}\nErrores: {errores}\n\nRevisa el archivo de log generado en el Escritorio.",
                    "WMS Makita - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show($"¡ÉXITO TOTAL!\n\nSe activaron los {actualizados} artículos limpiamente.",
                    "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                button1.Enabled = false;
                button1.Text = "Conectando...";

                // 1. Declaramos las credenciales en variables para usarlas en el JSON y en los Labels
                string miBaseDatos = "SBO_MAKITA_20260717"; ///
                string miUsuario = "manager";
                string miClave = "m1r1";

                // (Opcional) Si en las versiones modernas de .NET te da error de SSL, 
                // recuerda que esto se maneja mejor desde la declaración global del HttpClient, 
                // pero puedes mantener esta línea si te está funcionando bien en tu entorno local:
                System.Net.ServicePointManager.ServerCertificateValidationCallback += (senderCert, cert, chain, sslPolicyErrors) => true;

                // Configuramos la URL base sacada de tu SLD
                if (client.BaseAddress == null)
                {
                    client.BaseAddress = new Uri("https://192.168.1.17:50000/b1s/v1/");
                }

                // 2. Armamos el JSON inyectando las variables directamente
                string loginJson = $@"{{
            ""CompanyDB"": ""{miBaseDatos}"",
            ""UserName"": ""{miUsuario}"",
            ""Password"": ""{miClave}""
        }}";

                var content = new StringContent(loginJson, System.Text.Encoding.UTF8, "application/json");

                // Ejecutamos el POST al endpoint de Login
                HttpResponseMessage response = await client.PostAsync("Login", content);

                if (response.IsSuccessStatusCode)
                {
                    // Extraer las cookies (donde viene el B1SESSION)
                    var cookies = response.Headers.GetValues("Set-Cookie");
                    foreach (var cookie in cookies)
                    {
                        if (cookie.StartsWith("B1SESSION"))
                        {
                            sessionToken = cookie.Split(';')[0];
                            client.DefaultRequestHeaders.Add("Cookie", sessionToken);
                            break;
                        }
                    }

                    textBox1.Text = "Conexión exitosa - " + sessionToken;

                    // 3. ¡AQUÍ ACTUALIZAMOS TUS NUEVOS LABELS!
                    lblbdcon.Text = $"Base Datos: {miBaseDatos}";
                    lblUsuario.Text = $"Usuario     : {miUsuario}";

                    MessageBox.Show("Token generado correctamente. Ya podemos interactuar con la BD.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    MessageBox.Show("Error de credenciales o SL: " + error, "Error de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    lblbdcon.Text = "BD: Desconectado";
                    lblUsuario.Text = "Usuario: -";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de red/timeout: " + ex.Message, "Fallo Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                button1.Text = "Generar Token";
            }
        }

        private async void button7_Click_1(object sender, EventArgs e)
        {

            try
            {
                button1.Enabled = false;
                button1.Text = "Conectando...";

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (senderCert, cert, chain, sslPolicyErrors) => true;

                if (client.BaseAddress == null)
                {
                    client.BaseAddress = new Uri(slUrlBaseGlobal);
                }

                string loginJson = $@"{{
    ""CompanyDB"": ""{slBaseDatosGlobal}"",
    ""UserName"": ""{slUsuarioGlobal}"",
    ""Password"": ""{slClaveGlobal}""
}}";

                var content = new StringContent(loginJson, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync("Login", content);

                if (response.IsSuccessStatusCode)
                {
                    var cookies = response.Headers.GetValues("Set-Cookie");
                    foreach (var cookie in cookies)
                    {
                        if (cookie.StartsWith("B1SESSION"))
                        {
                            sessionToken = cookie.Split(';')[0];
                            client.DefaultRequestHeaders.Add("Cookie", sessionToken);
                            break;
                        }
                    }

                    textBox1.Text = "Conexión exitosa - " + sessionToken;

                    lblbdcon.Text = $"Base Datos: {slBaseDatosGlobal}";
                    lblUsuario.Text = $"Usuario     : {slUsuarioGlobal}";

                    MessageBox.Show("Token generado correctamente. Ya podemos interactuar con la BD.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    MessageBox.Show("Error de credenciales o SL: " + error, "Error de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    lblbdcon.Text = "BD: Desconectado";
                    lblUsuario.Text = "Usuario: -";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de red/timeout: " + ex.Message, "Fallo Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                button1.Text = "Generar Token";
            }

        }

        private async void button10_Click_1(object sender, EventArgs e)
        {
            // 1. Validamos que la grilla tenga datos cargados
            DataTable dtStock = dgv3.DataSource as DataTable;
            if (dtStock == null || dtStock.Rows.Count == 0)
            {
                MessageBox.Show("Primero debes consultar el stock en la grilla (Botón 7).", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Extraemos los códigos de artículo únicos de la grilla
            var codigosArticulos = dtStock.AsEnumerable()
                .Select(r => r["Codigo_Articulo"].ToString())
                .Distinct()
                .ToList();

            button10.Enabled = false;
            button10.Text = "Activando Masivamente...";

            // Configuración de la Barra de Progreso 2
            progressBar2.Minimum = 0;
            progressBar2.Maximum = codigosArticulos.Count;
            progressBar2.Value = 0;
            progressBar2.Visible = true;

            int actualizados = 0;
            int errores = 0;
            int procesados = 0;
            var logActivacion = new System.Collections.Concurrent.ConcurrentBag<string>();

            var tareas = new List<Task>();

            // 3. Concurrencia controlada a 15 hilos
            using (SemaphoreSlim semaphore = new SemaphoreSlim(15))
            {
                foreach (var itemCode in codigosArticulos)
                {
                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            string endpoint = $"Items('{itemCode}')";

                            // CORRECCIÓN CRÍTICA: Eliminamos ValidTo para evitar validaciones de SAP
                            // con ubicaciones por defecto inactivas. Solo enviamos el estado 'Activo'.
                            string jsonPatch = @"{
                        ""Valid"": ""tYES""
                    }";

                            var content = new StringContent(jsonPatch, Encoding.UTF8, "application/json");
                            var request = new HttpRequestMessage(new HttpMethod("PATCH"), endpoint) { Content = content };

                            var response = await client.SendAsync(request);

                            if (response.IsSuccessStatusCode)
                            {
                                System.Threading.Interlocked.Increment(ref actualizados);
                            }
                            else
                            {
                                string err = await response.Content.ReadAsStringAsync();
                                logActivacion.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL ACTIVAR {itemCode} | {err}");
                                System.Threading.Interlocked.Increment(ref errores);
                            }
                        }
                        catch (Exception ex)
                        {
                            logActivacion.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN {itemCode} | {ex.Message}");
                            System.Threading.Interlocked.Increment(ref errores);
                        }
                        finally
                        {
                            int current = System.Threading.Interlocked.Increment(ref procesados);

                            // Actualización segura de la barra de progreso
                            Invoke(new Action(() =>
                            {
                                progressBar2.Value = current > progressBar2.Maximum ? progressBar2.Maximum : current;
                            }));

                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tareas);
            }

            progressBar2.Value = progressBar2.Maximum;

            // Generar log si hubo errores
            if (errores > 0)
            {
                string ruta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_Activacion_WMS_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(ruta, logActivacion);
            }

            button10.Enabled = true;
            button10.Text = "Activar Artículos Masivamente";

            if (errores > 0)
            {
                MessageBox.Show($"¡Proceso finalizado con observaciones!\n\nArtículos activados: {actualizados}\nErrores: {errores}\n\nRevisa el archivo de log generado en el Escritorio.",
                    "WMS Makita - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show($"¡ÉXITO TOTAL!\n\nSe activaron los {actualizados} artículos limpiamente.",
                    "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            // 1. Reiniciar la barra de progreso y etiquetas
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 0;
            progressBar1.Visible = false;
            lblProgreso.Text = "Sistema reiniciado. Listo para operar.";

            // 2. Limpiar las grillas de datos
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            if (dgv3 != null)
            {
                dgv3.DataSource = null;
                dgv3.Rows.Clear();
                dgv3.Columns.Clear();
            }

            // 3. Reactivar controles principales
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;

            MessageBox.Show("Se ha limpiado la interfaz y reiniciado los contadores correctamente.",
                            "WMS Makita - Reset", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private async void button8_Click_1(object sender, EventArgs e)
        {
            // Reseteamos los contadores a nivel de clase
            itemsProcesados = 0;
            transferenciasExitosas = 0;
            errores = 0;
            totalLineasExitosas = 0;
            totalLineasFallidas = 0;

            int idUbicacionPulmon = 3; // ID confirmado de tu ubicación pulmón

            DataTable dtStock = dgv3.DataSource as DataTable;
            if (dtStock == null || dtStock.Rows.Count == 0)
            {
                MessageBox.Show("Primero debes consultar el stock en la grilla (Botón 7).", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            button8.Enabled = false;
            button8.Text = "Trasladando Stock...";
            button7.Enabled = false;

            System.Collections.Concurrent.ConcurrentBag<string> logReubicacion = new System.Collections.Concurrent.ConcurrentBag<string>();

            var agrupadoPorUbicacion = dtStock.AsEnumerable()
                .GroupBy(row => row["Id_Ubicacion_Origen"].ToString())
                .ToList();

            progressBar3.Minimum = 0;
            progressBar3.Maximum = agrupadoPorUbicacion.Count;
            progressBar3.Value = 0;
            progressBar3.Visible = true;

            // Procesamiento con SemaphoreSlim(1) para orden estricto en SAP
            using (SemaphoreSlim semaphore = new SemaphoreSlim(1))
            {
                var tareas = new List<Task>();

                foreach (var grupoUbicacion in agrupadoPorUbicacion)
                {
                    int idOrigen = Convert.ToInt32(grupoUbicacion.Key);
                    if (idOrigen == idUbicacionPulmon) continue;

                    string binCodeOrigen = grupoUbicacion.First()["Ubicacion_Origen"].ToString();
                    var localGrupoUbicacion = grupoUbicacion;

                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            List<object> lineasTransferencia = new List<object>();
                            var agrupadoPorArticulo = localGrupoUbicacion.GroupBy(r => r["Codigo_Articulo"].ToString());

                            foreach (var grupoArticulo in agrupadoPorArticulo)
                            {
                                string itemCode = grupoArticulo.Key;
                                string manejaSeries = grupoArticulo.First()["Maneja_Series"].ToString();

                                double cantidadTotalItem = 0;
                                var listaSeries = new List<object>();
                                var allocs = new List<object>();

                                if (manejaSeries == "Y")
                                {
                                    int serialIndex = 0;
                                    foreach (var fila in grupoArticulo)
                                    {
                                        double qty = Convert.ToDouble(fila["Cantidad_A_Mover"]);
                                        cantidadTotalItem += qty;
                                        listaSeries.Add(new { SystemSerialNumber = Convert.ToInt32(fila["Id_Serie_Interno"]), Quantity = qty });
                                        allocs.Add(new { BinAbsEntry = idOrigen, Quantity = qty, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = serialIndex, BinActionType = "batFromWarehouse" });
                                        allocs.Add(new { BinAbsEntry = idUbicacionPulmon, Quantity = qty, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = serialIndex, BinActionType = "batToWarehouse" });
                                        serialIndex++;
                                    }
                                    lineasTransferencia.Add(new { ItemCode = itemCode, Quantity = cantidadTotalItem, WarehouseCode = "ALM01", FromWarehouseCode = "ALM01", SerialNumbers = listaSeries, StockTransferLinesBinAllocations = allocs });
                                }
                                else
                                {
                                    foreach (var fila in grupoArticulo) { cantidadTotalItem += Convert.ToDouble(fila["Cantidad_A_Mover"]); }
                                    allocs.Add(new { BinAbsEntry = idOrigen, Quantity = cantidadTotalItem, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = 0, BinActionType = "batFromWarehouse" });
                                    allocs.Add(new { BinAbsEntry = idUbicacionPulmon, Quantity = cantidadTotalItem, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = 0, BinActionType = "batToWarehouse" });
                                    lineasTransferencia.Add(new { ItemCode = itemCode, Quantity = cantidadTotalItem, WarehouseCode = "ALM01", FromWarehouseCode = "ALM01", StockTransferLinesBinAllocations = allocs });
                                }
                            }

                            // =========================================================================
                            // CLAVE: PARTIR EN BLOQUES DE 20 LÍNEAS MÁXIMO POR CADA DOCUMENTO DE SAP
                            // =========================================================================
                            if (lineasTransferencia.Count > 0)
                            {
                                const int TAMANO_LOTE = 20; // Ajustado a 20 líneas para máxima fluidez
                                for (int i = 0; i < lineasTransferencia.Count; i += TAMANO_LOTE)
                                {
                                    var chunkLineas = lineasTransferencia.Skip(i).Take(TAMANO_LOTE).ToList();
                                    int cantLineas = chunkLineas.Count;

                                    bool exito = await EnviarTransferencia(chunkLineas, binCodeOrigen, logReubicacion);

                                    if (exito)
                                    {
                                        System.Threading.Interlocked.Add(ref totalLineasExitosas, cantLineas);
                                        System.Threading.Interlocked.Increment(ref transferenciasExitosas);
                                    }
                                    else
                                    {
                                        System.Threading.Interlocked.Add(ref totalLineasFallidas, cantLineas);
                                        System.Threading.Interlocked.Increment(ref errores);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Threading.Interlocked.Increment(ref errores);
                            logReubicacion.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN UBICACIÓN {binCodeOrigen} | {ex.Message}");
                        }
                        finally
                        {
                            System.Threading.Interlocked.Increment(ref itemsProcesados);
                            Invoke(new Action(() =>
                            {
                                if (progressBar3.Value < progressBar3.Maximum)
                                    progressBar3.Value++;
                            }));
                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tareas);
            }

            progressBar3.Value = progressBar3.Maximum;
            button8.Enabled = true;
            button7.Enabled = true;

            if (errores > 0)
            {
                string ruta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_WMS_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(ruta, logReubicacion);
                MessageBox.Show($"Migración finalizada con observaciones.\n\nDocumentos Exitosos: {transferenciasExitosas} | Líneas Trasladadas: {totalLineasExitosas}\nDocumentos con Error: {errores} | Líneas Fallidas: {totalLineasFallidas}\n\nRevisa el archivo de log en el escritorio.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show($"¡ÉXITO TOTAL!\n\nSe trasladaron todas las líneas de stock exitosamente al pulmón.\nTotal de líneas procesadas: {totalLineasExitosas}", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private async void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                button2.Enabled = false;
                button2.Text = "Consultando HANA, espere...";

                label1.Text = "Se encontró : Buscando...";

                // 1. Configuramos las columnas del DataGridView 1 
                dataGridView1.Columns.Clear();
                dataGridView1.Columns.Add("Nro", "#");

                // ¡AGREGAMOS EL ABSENTRY (Fundamental para el Botón 4)!
                dataGridView1.Columns.Add("AbsEntry", "ID Interno");

                dataGridView1.Columns.Add("BinCode", "Código Ubicación");
                dataGridView1.Columns.Add("WhsCode", "Almacén");
                dataGridView1.Columns.Add("Sub1", "Fila");
                dataGridView1.Columns.Add("Sub2", "Cuerpo");
                dataGridView1.Columns.Add("Sub3", "Piso");
                dataGridView1.Columns.Add("Sub4", "Fondo");
                dataGridView1.Columns.Add("Inactive", "Inactivo");

                // Ajustamos diseño visual
                dataGridView1.Columns["Nro"].Width = 40;

                // ¡LA MAGIA AQUÍ! Ocultamos la columna para no ensuciar tu diseño visual en pantalla
                dataGridView1.Columns["AbsEntry"].Visible = false;

                // 2. Endpoint actualizado: Pedimos 'AbsEntry' en el $select a la API
                string nextLink = "BinLocations?$select=AbsEntry,BinCode,Warehouse,Sublevel1,Sublevel2,Sublevel3,Sublevel4,Inactive";
                int totalRecords = 0;

                // 3. Ciclo de paginación OData (recorrerá todas las ubicaciones)
                while (!string.IsNullOrEmpty(nextLink))
                {
                    HttpResponseMessage response = await client.GetAsync(nextLink);
                    response.EnsureSuccessStatusCode();

                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                    {
                        JsonElement root = doc.RootElement;
                        JsonElement values = root.GetProperty("value");

                        foreach (JsonElement item in values.EnumerateArray())
                        {
                            // Capturamos el AbsEntry y todos los campos nativos
                            string absEntry = item.GetProperty("AbsEntry").ToString();
                            string binCode = item.GetProperty("BinCode").ToString();
                            string whsCode = item.GetProperty("Warehouse").ToString();
                            string fila = item.GetProperty("Sublevel1").ToString();
                            string cuerpo = item.GetProperty("Sublevel2").ToString();
                            string piso = item.GetProperty("Sublevel3").ToString();
                            string fondo = item.GetProperty("Sublevel4").ToString();
                            string inactive = item.GetProperty("Inactive").ToString();

                            totalRecords++;

                            // Llenamos la fila incluyendo el absEntry (que estará oculto)
                            dataGridView1.Rows.Add(totalRecords, absEntry, binCode, whsCode, fila, cuerpo, piso, fondo, inactive);
                        }

                        // 4. Verificamos si HANA nos manda una siguiente "página" de resultados
                        if (root.TryGetProperty("odata.nextLink", out JsonElement nextLinkElement))
                        {
                            string link = nextLinkElement.GetString();
                            nextLink = link.Contains("/b1s/v1/") ? link.Substring(link.IndexOf("/b1s/v1/") + 8) : link;
                        }
                        else
                        {
                            nextLink = null; // Terminamos de leer toda la base
                        }
                    }
                }

                // 5. Actualizamos el Label con el total exacto de registros encontrados
                label1.Text = $"Se encontró : {totalRecords}";

                MessageBox.Show($"¡Éxito! Se cargaron {totalRecords} ubicaciones actuales desde la base de datos.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al consultar ubicaciones: " + ex.Message, "Error Service Layer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                label1.Text = "Se encontró : Error";
            }
            finally
            {
                button2.Enabled = true;
                button2.Text = "Consulta Ubicaciones en BD";
            }


        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Seleccione el archivo DATAMASTER";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    button3.Enabled = false;
                    button3.Text = "Cargando Excel...";
                    label2.Text = "Total Excel : Cargando...";

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            DataTable dtExcel = result.Tables["DATAMASTER"];

                            if (dtExcel != null)
                            {
                                int totalOriginal = dtExcel.Rows.Count;

                                // Validamos y filtramos registros únicos basados en las columnas B, C, D, E y F (índices 1 al 5)
                                var filasUnicas = dtExcel.AsEnumerable()
                                    .GroupBy(row => new
                                    {
                                        Almacen = row[1]?.ToString().Trim(),
                                        Fila = row[2]?.ToString().Trim(),
                                        Cuerpo = row[3]?.ToString().Trim(),
                                        Piso = row[4]?.ToString().Trim(),
                                        Fondo = row[5]?.ToString().Trim()
                                    })
                                    .Select(grupo => grupo.First())
                                    .ToList();

                                int totalUnicos = filasUnicas.Count;
                                int totalDuplicados = totalOriginal - totalUnicos;

                                // 1. DGV2 estructurado idénticamente a DGV1 (sin campos extra)
                                dataGridView2.Columns.Clear();
                                dataGridView2.Columns.Add("Nro", "#");
                                dataGridView2.Columns.Add("BinCodeRef", "Ubicación Referencial");
                                dataGridView2.Columns.Add("Almacen", "Almacén");
                                dataGridView2.Columns.Add("Sub1", "Fila");
                                dataGridView2.Columns.Add("Sub2", "Cuerpo");
                                dataGridView2.Columns.Add("Sub3", "Piso");
                                dataGridView2.Columns.Add("Sub4", "Fondo");
                                dataGridView2.Columns.Add("Inactive", "Inactivo");

                                dataGridView2.Columns["Nro"].Width = 40;

                                // 2. Volcamos EXCLUSIVAMENTE los registros únicos a la grilla
                                int rowCount = 0;
                                foreach (var row in filasUnicas)
                                {
                                    rowCount++;

                                    string binRef = row[0].ToString();
                                    string whs = row[1].ToString();
                                    string fila = row[2].ToString();
                                    string cuerpo = row[3].ToString();
                                    string piso = row[4].ToString();
                                    string fondo = row[5].ToString();

                                    dataGridView2.Rows.Add(rowCount, binRef, whs, fila, cuerpo, piso, fondo, "tNO");
                                }

                                label2.Text = $"Total Excel : {totalUnicos}";

                                // 3. Notificación con el desglose exacto de duplicados omitidos y registros únicos
                                MessageBox.Show(
                                    $"¡Importación procesada con éxito!\n\n" +
                                    $"• Total de filas leídas en Excel: {totalOriginal}\n" +
                                    $"• Registros duplicados omitidos: {totalDuplicados}\n" +
                                    $"• Registros únicos cargados: {totalUnicos}",
                                    "WMS Makita - Control de Duplicados",
                                    MessageBoxButtons.OK,
                                    totalDuplicados > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information
                                );
                            }
                            else
                            {
                                MessageBox.Show("No se encontró una hoja llamada 'DATAMASTER' en el Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                label2.Text = "Total Excel : Error";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Fallo de Importación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    label2.Text = "Total Excel : Error";
                }
                finally
                {
                    button3.Enabled = true;
                    button3.Text = "Importar Ubicaciones Finales/Reales";
                }
            }


        }

        private async void button6_Click_1(object sender, EventArgs e)
        {
            button6.Enabled = false;
            button6.Text = "Sincronizando Subniveles...";

            // 1. Extraemos solo valores ÚNICOS
            HashSet<string> filasUnicas = new HashSet<string>();
            HashSet<string> cuerposUnicos = new HashSet<string>();
            HashSet<string> pisosUnicos = new HashSet<string>();
            HashSet<string> fondosUnicos = new HashSet<string>();

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                string fila = row.Cells["Sub1"].Value?.ToString();
                string cuerpo = row.Cells["Sub2"].Value?.ToString();
                string piso = row.Cells["Sub3"].Value?.ToString();
                string fondo = row.Cells["Sub4"].Value?.ToString();

                if (!string.IsNullOrEmpty(fila)) filasUnicas.Add(fila);
                if (!string.IsNullOrEmpty(cuerpo)) cuerposUnicos.Add(cuerpo);
                if (!string.IsNullOrEmpty(piso)) pisosUnicos.Add(piso);
                if (!string.IsNullOrEmpty(fondo)) fondosUnicos.Add(fondo);
            }

            int creadosTotales = 0;
            int omitidosTotales = 0;

            // 2. Ejecutamos la sincronización llamando al método externo

            // Filas (Nivel 1)
            int creadosFila = await EnviarSubnivelesASap(filasUnicas, 1, "Fila");
            creadosTotales += creadosFila;
            omitidosTotales += (filasUnicas.Count - creadosFila);

            // Cuerpos (Nivel 2)
            int creadosCuerpo = await EnviarSubnivelesASap(cuerposUnicos, 2, "Cuerpo");
            creadosTotales += creadosCuerpo;
            omitidosTotales += (cuerposUnicos.Count - creadosCuerpo);

            // Pisos (Nivel 3)
            int creadosPiso = await EnviarSubnivelesASap(pisosUnicos, 3, "Piso");
            creadosTotales += creadosPiso;
            omitidosTotales += (pisosUnicos.Count - creadosPiso);

            // Fondos (Nivel 4)
            int creadosFondo = await EnviarSubnivelesASap(fondosUnicos, 4, "Fondo");
            creadosTotales += creadosFondo;
            omitidosTotales += (fondosUnicos.Count - creadosFondo);

            // 3. Reactivamos Interfaz
            button6.Enabled = true;
            button6.Text = "Sincronizar Subniveles Faltantes";

            MessageBox.Show($"Sincronización finalizada.\n\nNuevos subniveles creados: {creadosTotales}\nYa existentes (Omitidos): {omitidosTotales}\n\n¡Ya puedes ejecutar la creación masiva de ubicaciones!",
                            "WMS Makita - Preparación de BD", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private async void button4_Click_1(object sender, EventArgs e)
        {


            button4.Enabled = false;
            button4.Text = "Procesando HANA...";
            dataGridView1.Enabled = false;
            dataGridView2.Enabled = false;

            // Contadores Thread-Safe
            int creadas = 0;
            int actualizadas = 0;
            int errores = 0;
            int filasProcesadas = 0;
            int inactivadasFantasmas = 0;

            System.Collections.Concurrent.ConcurrentBag<string> logErrores = new System.Collections.Concurrent.ConcurrentBag<string>();

            int totalAProcesar = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (!row.IsNewRow) totalAProcesar++;
            }

            var progress = new Progress<int>(procesadas =>
            {
                int porcentaje = totalAProcesar > 0 ? (int)((double)procesadas / totalAProcesar * 100) : 0;
                if (porcentaje > 100) porcentaje = 100;

                progressBar1.Value = porcentaje;
                lblProgreso.Text = $"Procesando: {procesadas} de {totalAProcesar} ({porcentaje}%)";
            });

            // 2. Cargamos HANA a la Memoria RAM (Cruce ultra rápido)
            var ubicacionesHANA = new Dictionary<string, string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;
                string binCode = row.Cells["BinCode"].Value?.ToString();
                string absEntry = row.Cells["AbsEntry"].Value?.ToString();

                if (!string.IsNullOrEmpty(binCode))
                {
                    // SOLUCIÓN AL ERROR -2035: Limpiamos los separadores y espacios
                    string binCodeLimpio = binCode.Replace("-", "").Trim();
                    ubicacionesHANA[binCodeLimpio] = absEntry;
                }
            }

            var tareas = new List<Task>();

            // 3. Concurrencia controlada a 15 hilos hacia Service Layer
            using (SemaphoreSlim semaphore = new SemaphoreSlim(15))
            {
                foreach (DataGridViewRow rowExcel in dataGridView2.Rows)
                {
                    if (rowExcel.IsNewRow) continue;

                    // Limpieza extrema de espacios vacíos (trailing spaces) de cada celda del Excel
                    string whs = rowExcel.Cells["Almacen"].Value?.ToString().Trim() ?? "";
                    string fila = rowExcel.Cells["Sub1"].Value?.ToString().Trim() ?? "";
                    string cuerpo = rowExcel.Cells["Sub2"].Value?.ToString().Trim() ?? "";
                    string piso = rowExcel.Cells["Sub3"].Value?.ToString().Trim() ?? "";
                    string fondo = rowExcel.Cells["Sub4"].Value?.ToString().Trim() ?? "";
                    string inactive = rowExcel.Cells["Inactive"].Value?.ToString().Trim() ?? "tNO";

                    // Armamos la llave limpia
                    string cleanBinCode = $"{whs}{fila}{cuerpo}{piso}{fondo}";

                    // Protección de las ubicaciones de sistema
                    if (cleanBinCode.Contains("UBICACIÓN-DE-SISTEMA") || cleanBinCode.Contains("SYSTEM"))
                    {
                        continue;
                    }

                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            if (ubicacionesHANA.ContainsKey(cleanBinCode))
                            {
                                // ==========================================
                                // ACTUALIZAR (PATCH) - La ubicación YA existe
                                // ==========================================
                                string entry = ubicacionesHANA[cleanBinCode];

                                // Inyectamos el valor "2" (Picking) 
                                string jsonPatch = $@"{{ 
                            ""Inactive"": ""{inactive}"",
                            ""U_WAR_DW_TYPE_LOCATION"": ""2""
                        }}";

                                var content = new StringContent(jsonPatch, System.Text.Encoding.UTF8, "application/json");
                                var response = await client.PatchAsync($"BinLocations({entry})", content);

                                if (!response.IsSuccessStatusCode)
                                {
                                    string sapError = await response.Content.ReadAsStringAsync();
                                    throw new Exception($"HTTP {(int)response.StatusCode} - Detalles SAP: {sapError}");
                                }

                                System.Threading.Interlocked.Increment(ref actualizadas);
                            }
                            else
                            {
                                // ==========================================
                                // CREAR (POST) - La ubicación es NUEVA
                                // ==========================================
                                string jsonPost = $@"{{
                            ""Warehouse"": ""{whs}"",
                            ""Sublevel1"": ""{fila}"",
                            ""Sublevel2"": ""{cuerpo}"",
                            ""Sublevel3"": ""{piso}"",
                            ""Sublevel4"": ""{fondo}"",
                            ""Inactive"": ""{inactive}"",
                            ""U_WAR_DW_TYPE_LOCATION"": ""2""
                        }}";

                                var content = new StringContent(jsonPost, System.Text.Encoding.UTF8, "application/json");
                                var response = await client.PostAsync("BinLocations", content);

                                if (!response.IsSuccessStatusCode)
                                {
                                    string sapError = await response.Content.ReadAsStringAsync();
                                    throw new Exception($"HTTP {(int)response.StatusCode} - Detalles SAP: {sapError}");
                                }

                                System.Threading.Interlocked.Increment(ref creadas);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Threading.Interlocked.Increment(ref errores);
                            string tipoOperacion = ubicacionesHANA.ContainsKey(cleanBinCode) ? "ACTUALIZAR" : "CREAR";
                            logErrores.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL {tipoOperacion} | Ubicación: {cleanBinCode} | Motivo: {ex.Message}");
                        }
                        finally
                        {
                            int procesadasLocales = System.Threading.Interlocked.Increment(ref filasProcesadas);
                            ((IProgress<int>)progress).Report(procesadasLocales);
                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tareas);
            }

            // ====================================================================
            // INICIO ESCENARIO 3: BARRIDO DE INACTIVACIÓN (Limpieza de fantasmas)
            // ====================================================================
            lblProgreso.Text = "Fase 2: Buscando y bloqueando ubicaciones sobrantes en HANA...";

            HashSet<string> codigosExcel = new HashSet<string>();
            foreach (DataGridViewRow rowExcel in dataGridView2.Rows)
            {
                if (rowExcel.IsNewRow) continue;
                string whs = rowExcel.Cells["Almacen"].Value?.ToString().Trim() ?? "";
                string fila = rowExcel.Cells["Sub1"].Value?.ToString().Trim() ?? "";
                string cuerpo = rowExcel.Cells["Sub2"].Value?.ToString().Trim() ?? "";
                string piso = rowExcel.Cells["Sub3"].Value?.ToString().Trim() ?? "";
                string fondo = rowExcel.Cells["Sub4"].Value?.ToString().Trim() ?? "";

                codigosExcel.Add($"{whs}{fila}{cuerpo}{piso}{fondo}");
            }

            var tareasLimpieza = new List<Task>();

            using (SemaphoreSlim semaforoLimpieza = new SemaphoreSlim(15))
            {
                foreach (DataGridViewRow rowHana in dataGridView1.Rows)
                {
                    if (rowHana.IsNewRow) continue;

                    string binCodeHana = rowHana.Cells["BinCode"].Value?.ToString().Trim() ?? "";
                    string absEntryHana = rowHana.Cells["AbsEntry"].Value?.ToString() ?? "";
                    string statusActual = rowHana.Cells["Inactive"].Value?.ToString() ?? "tNO";

                    // Limpiamos los guiones del código HANA para poder compararlo con el HashSet
                    string binCodeHanaLimpio = binCodeHana.Replace("-", "");

                    if (binCodeHanaLimpio.Contains("UBICACIÓN-DE-SISTEMA") || binCodeHanaLimpio.Contains("SYS"))
                    {
                        continue;
                    }

                    // REGLA: Si está en HANA, pero NO está en Excel, y está ACTIVA -> ¡Inactivar!
                    if (!string.IsNullOrEmpty(binCodeHanaLimpio) && !codigosExcel.Contains(binCodeHanaLimpio) && statusActual == "tNO")
                    {
                        await semaforoLimpieza.WaitAsync();

                        tareasLimpieza.Add(Task.Run(async () =>
                        {
                            try
                            {
                                string jsonPatch = $@"{{ ""Inactive"": ""tYES"" }}";
                                var content = new StringContent(jsonPatch, System.Text.Encoding.UTF8, "application/json");

                                var response = await client.PatchAsync($"BinLocations({absEntryHana})", content);

                                if (response.IsSuccessStatusCode)
                                {
                                    System.Threading.Interlocked.Increment(ref inactivadasFantasmas);
                                }
                                else
                                {
                                    string sapError = await response.Content.ReadAsStringAsync();
                                    logErrores.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL INACTIVAR SOBRANTE | Ubicación: {binCodeHana} | Motivo: {sapError}");
                                    System.Threading.Interlocked.Increment(ref errores);
                                }
                            }
                            catch (Exception ex)
                            {
                                logErrores.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL INACTIVAR SOBRANTE | Ubicación: {binCodeHana} | Motivo: {ex.Message}");
                                System.Threading.Interlocked.Increment(ref errores);
                            }
                            finally
                            {
                                semaforoLimpieza.Release();
                            }
                        }));
                    }
                }
                await Task.WhenAll(tareasLimpieza);
            }
            // ====================================================================
            // FIN ESCENARIO 3
            // ====================================================================

            // 4. GENERACIÓN DEL ARCHIVO FÍSICO DE LOG (Si hubo errores)
            if (errores > 0)
            {
                try
                {
                    string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string nombreArchivo = $"Log_WMS_Makita_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                    string rutaCompleta = System.IO.Path.Combine(rutaEscritorio, nombreArchivo);

                    System.IO.File.WriteAllLines(rutaCompleta, logErrores);

                    MessageBox.Show($"Migración masiva finalizada con observaciones.\n\nCreadas: {creadas}\nActualizadas: {actualizadas}\nInactivadas (Sobrantes): {inactivadasFantasmas}\nErrores: {errores}\n\nRevisa el archivo de Log generado.",
                                    "WMS Makita - Log Generado", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = rutaCompleta,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("No se pudo guardar/abrir el archivo Log: " + ex.Message, "Error de I/O", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show($"¡Lote MASIVO procesado limpiamente!\n\nCreadas: {creadas}\nActualizadas: {actualizadas}\nInactivadas (Sobrantes): {inactivadasFantasmas}\nErrores: 0",
                                "WMS Makita - Éxito Total", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            dataGridView1.Enabled = true;
            dataGridView2.Enabled = true;
            button4.Enabled = true;
            button4.Text = "Actualizar/Crear en BD";
            lblProgreso.Text = "Proceso Completado.";





        }

        private async void button1_Click_2(object sender, EventArgs e)
        {
            try
            {
                button1.Enabled = false;
                button1.Text = "Conectando...";

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (senderCert, cert, chain, sslPolicyErrors) => true;

                if (client.BaseAddress == null)
                {
                    client.BaseAddress = new Uri(slUrlBaseGlobal);
                }

                string loginJson = $@"{{    ""CompanyDB"": ""{slBaseDatosGlobal}"",    ""UserName"": ""{slUsuarioGlobal}"",    ""Password"": ""{slClaveGlobal}""}}";

                var content = new StringContent(loginJson, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync("Login", content);

                if (response.IsSuccessStatusCode)
                {
                    var cookies = response.Headers.GetValues("Set-Cookie");
                    foreach (var cookie in cookies)
                    {
                        if (cookie.StartsWith("B1SESSION"))
                        {
                            sessionToken = cookie.Split(';')[0];
                            client.DefaultRequestHeaders.Add("Cookie", sessionToken);
                            break;
                        }
                    }

                    textBox1.Text = "Conexión exitosa - " + sessionToken;

                    // Asignación directa a los labels usando las variables globales
                    label6.Text = $"Base de Datos: {slBaseDatosGlobal}";
                    label7.Text = $"Usuario: {slUsuarioGlobal}";

                    MessageBox.Show("Token generado correctamente. Ya podemos interactuar con la BD.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    MessageBox.Show("Error de credenciales o SL: " + error, "Error de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    label6.Text = "Base de Datos: Desconectado";
                    label7.Text = "Usuario: -";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de red/timeout: " + ex.Message, "Fallo Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                button1.Text = "Generar Token";
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Seleccione el Layout de Picking (Fase 2)";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 1. Preparamos la interfaz visual para la Fase 2
                    dgvFase2.Columns.Clear();
                    dgvFase2.Columns.Add("Nro", "#");
                    dgvFase2.Columns.Add("UbicacionDestino", "Ubicación Destino (SAP)");
                    dgvFase2.Columns.Add("CodigoArticulo", "Código Artículo");
                    dgvFase2.Columns.Add("EstadoValidacion", "Estado Validación");

                    // Columna oculta que llenaremos en el Paso 2 (Validación)
                    dgvFase2.Columns.Add("AbsEntryDestino", "ID Destino");
                    dgvFase2.Columns["AbsEntryDestino"].Visible = false;

                    dgvFase2.Columns["Nro"].Width = 40;
                    dgvFase2.Columns["UbicacionDestino"].Width = 180;
                    dgvFase2.Columns["CodigoArticulo"].Width = 150;
                    dgvFase2.Columns["EstadoValidacion"].Width = 150;

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });

                            // Asumimos que la data está en la primera hoja del Excel
                            DataTable dtExcel = result.Tables[0];

                            if (dtExcel != null)
                            {
                                int filaNro = 0;
                                int omitidosBlancos = 0;

                                foreach (DataRow row in dtExcel.Rows)
                                {
                                    // Columna A: Ubicación | Columna B: Código
                                    string ubicacion = row[0]?.ToString().Trim();
                                    string codigo = row[1]?.ToString().Trim();

                                    // REGLA DE NEGOCIO: Ignorar "Huecos" (filas sin código de artículo)
                                    if (string.IsNullOrEmpty(codigo))
                                    {
                                        omitidosBlancos++;
                                        continue;
                                    }

                                    if (!string.IsNullOrEmpty(ubicacion))
                                    {
                                        filaNro++;
                                        // Insertamos en la grilla. Estado inicial: "Pendiente"
                                        dgvFase2.Rows.Add(filaNro, ubicacion, codigo, "Pendiente", "");
                                    }
                                }

                                // Actualizamos los controles visuales (Ajusta los nombres de tus label)
                                label3.Text = $"Cargados: {filaNro} | Huecos: {omitidosBlancos}";

                                MessageBox.Show($"¡Layout de Picking importado con éxito!\n\n" +
                                                $"• Artículos a reubicar: {filaNro}\n" +
                                                $"• Espacios vacíos reservados (omitidos): {omitidosBlancos}",
                                                "WMS Makita - Fase 2", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Error de Importación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }




            }
        }

        private async void button12_Click(object sender, EventArgs e)
        {

            button12.Enabled = false;
            button12.Text = "Consultando HANA (Validando)...";
            int validos = 0;
            int erroresDestino = 0;
            int erroresStock = 0;

            try
            {
                var dictUbicaciones = new Dictionary<string, string>(); // Guardará: BinCode -> AbsEntry
                var stockPulmon = new HashSet<string>();                // Guardará: Códigos de artículos con stock en Pulmón

                // 2. Ejecutamos las consultas a HANA en segundo plano (para no congelar la pantalla)
                await Task.Run(() =>
                {
                    using (var conn = new System.Data.Odbc.OdbcConnection(GetHanaConnectionString()))
                    {
                        conn.Open();

                        // A) Traemos todas las ubicaciones de ALM01 y sus AbsEntry
                        string queryUbic = "SELECT \"BinCode\", \"AbsEntry\" FROM OBIN WHERE \"WhsCode\" = 'ALM01'";
                        using (var cmd = new System.Data.Odbc.OdbcCommand(queryUbic, conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                dictUbicaciones[reader.GetString(0)] = reader.GetString(1);
                            }
                        }

                        // B) Traemos ÚNICAMENTE los artículos que actualmente tienen stock en el pulmón (BinAbs = 3)
                        string queryStock = "SELECT DISTINCT \"ItemCode\" FROM OIBQ WHERE \"BinAbs\" = 3 AND \"OnHandQty\" > 0";
                        using (var cmd = new System.Data.Odbc.OdbcCommand(queryStock, conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                stockPulmon.Add(reader.GetString(0));
                            }
                        }
                    }
                });

                // 3. Cruzamos la información de HANA contra las filas del Excel importadas en la grilla
                foreach (DataGridViewRow row in dgvFase2.Rows)
                {
                    if (row.IsNewRow) continue;

                    string ubicacionDestino = row.Cells["UbicacionDestino"].Value?.ToString().Trim();
                    string codigoArticulo = row.Cells["CodigoArticulo"].Value?.ToString().Trim();

                    // ==========================================
                    // SOLUCIÓN: Limpiar los guiones para que coincida exactamente con la base de datos de SAP
                    // ==========================================
                    string cleanDestino = ubicacionDestino.Replace("-", "");

                    // Validaciones lógicas usando cleanDestino
                    bool existeUbicacion = dictUbicaciones.ContainsKey(cleanDestino);
                    bool hayStock = stockPulmon.Contains(codigoArticulo);

                    // 4. Pintamos la grilla según el resultado
                    if (!existeUbicacion)
                    {
                        row.Cells["EstadoValidacion"].Value = "ERROR: Destino No Existe en SAP";
                        row.DefaultCellStyle.BackColor = Color.LightCoral; // Rojo suave
                        erroresDestino++;
                    }
                    else if (!hayStock)
                    {
                        row.Cells["EstadoValidacion"].Value = "ERROR: Artículo Sin Stock en Pulmón";
                        row.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow; // Amarillo suave
                        erroresStock++;
                    }
                    else
                    {
                        row.Cells["EstadoValidacion"].Value = "OK - Listo";
                        row.Cells["AbsEntryDestino"].Value = dictUbicaciones[cleanDestino]; // Asignación con cleanDestino
                        row.DefaultCellStyle.BackColor = Color.LightGreen; // Verde suave
                        validos++;
                    }
                }

                // 5. Actualizamos visuales y damos el veredicto
                label3.Text = $"Validados: {validos} | Err. Destino: {erroresDestino} | Err. Stock: {erroresStock}";

                if (erroresDestino == 0 && erroresStock == 0 && validos > 0)
                {
                    MessageBox.Show("¡Validación Perfecta!\n\nTodos los destinos existen estructuralmente y hay stock disponible en el pulmón.\nYa puedes ejecutar la reubicación masiva.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Validación finalizada con observaciones.\n\nListos para reubicar: {validos}\nUbicaciones destino inexistentes: {erroresDestino}\nArtículos sin stock en Pulmón: {erroresStock}\n\nPor favor, revisa las filas marcadas en color antes de continuar.", "WMS Makita - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de conexión durante la validación ODBC:\n\n" + ex.Message, "Fallo Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button12.Enabled = true;
                button12.Text = "2. Pre-Validar Distribución";
            }


        }

        private async void button13_Click(object sender, EventArgs e)
        {
            button13.Enabled = false;
            button13.Text = "Ejecutando Putaway...";
            button12.Enabled = false;

            int documentosExitosos = 0;
            int lineasTrasladadas = 0;
            int erroresPutaway = 0;
            int idUbicacionPulmon = 3;

            System.Collections.Concurrent.ConcurrentBag<string> logPutaway = new System.Collections.Concurrent.ConcurrentBag<string>();

            //string dbServer = "192.168.1.17:30015";
            //string dbUser = "SAPINST";
            //string dbPass = "Ki$ta66mpe";
            //string dbSchema = "SBO_MAKITA_20260717";
            //string connStr = $"Driver={{HDBODBC}};ServerNode={dbServer};UID={dbUser};PWD={dbPass};CS={dbSchema};";



            string connStr = GetHanaConnectionString();

            try
            {
                var filasValidadas = dgvFase2.Rows.Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow && r.Cells["EstadoValidacion"].Value?.ToString() == "OK - Listo")
                    .ToList();

                if (filasValidadas.Count == 0)
                {
                    // Calculamos el total de filas reales del Excel importado
                    int totalFilas = dgvFase2.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);

                    MessageBox.Show($"Las {totalFilas} ubicaciones definidas en el excel, fueron reubicadas con éxito a las nuevas ubicaciones finales.",
                                    "WMS Makita - Operación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    button13.Enabled = true;
                    button13.Text = "3. Ejecutar Putaway (Reubicar)";
                    button12.Enabled = true;
                    return;
                }

                // === 1. LOG Y UI: Trazabilidad de Inicio ===
                System.Diagnostics.Debug.WriteLine($"\n=== INICIANDO PUTAWAY: {filasValidadas.Count} filas validadas ===");
                label3.Text = "Paso 1: Obteniendo stock detallado de HANA...";

                DataTable dtStockPulmon = new DataTable();
                string queryPulmon = @"SELECT T1.""ItemCode"" AS ""Codigo_Articulo"", CASE WHEN T3.""ManSerNum"" = 'Y' THEN 1 ELSE T1.""OnHandQty"" END AS ""Cantidad_A_Mover"", T3.""ManSerNum"" AS ""Maneja_Series"", T4.""SysNumber"" AS ""Id_Serie_Interno"" FROM OIBQ T1 INNER JOIN OITM T3 ON T1.""ItemCode"" = T3.""ItemCode"" LEFT JOIN OSBQ T2 ON T1.""ItemCode"" = T2.""ItemCode"" AND T1.""BinAbs"" = T2.""BinAbs"" AND T2.""OnHandQty"" > 0 LEFT JOIN OSRN T4 ON T2.""SnBMDAbs"" = T4.""AbsEntry"" WHERE T1.""BinAbs"" = 3 AND T1.""OnHandQty"" > 0;";

                await Task.Run(() =>
                {
                    System.Diagnostics.Debug.WriteLine("-> Ejecutando Query ODBC en Pulmón...");
                    using (var conn = new System.Data.Odbc.OdbcConnection(connStr))
                    {
                        conn.Open();
                        using (var cmd = new System.Data.Odbc.OdbcCommand(queryPulmon, conn))
                        using (var da = new System.Data.Odbc.OdbcDataAdapter(cmd))
                        {
                            da.Fill(dtStockPulmon);
                        }
                    }
                    System.Diagnostics.Debug.WriteLine($"-> Query OK. Filas de stock recuperadas: {dtStockPulmon.Rows.Count}");
                });

                // === 2. LOG Y UI: Trazabilidad de Procesamiento ===
                int procesadosActuales = 0;
                label3.Text = $"Paso 2: Iniciando transferencias (0 / {filasValidadas.Count})...";

                using (SemaphoreSlim semaphore = new SemaphoreSlim(1))
                {
                    var tareas = new List<Task>();

                    foreach (var filaExcel in filasValidadas)
                    {
                        string itemCode = filaExcel.Cells["CodigoArticulo"].Value.ToString().Trim();
                        int idDestinoFinal = Convert.ToInt32(filaExcel.Cells["AbsEntryDestino"].Value);
                        string nombreDestino = filaExcel.Cells["UbicacionDestino"].Value.ToString();

                        var stockArticulo = dtStockPulmon.AsEnumerable()
                                            .Where(r => r["Codigo_Articulo"].ToString() == itemCode)
                                            .ToList();

                        if (stockArticulo.Count == 0) continue;

                        await semaphore.WaitAsync();

                        tareas.Add(Task.Run(async () =>
                        {
                            try
                            {
                                System.Diagnostics.Debug.WriteLine($"--- Armando JSON para {itemCode} hacia {nombreDestino} ---");
                                List<object> lineasTransferencia = new List<object>();
                                string manejaSeries = stockArticulo.First()["Maneja_Series"].ToString();

                                double cantidadTotalItem = 0;
                                var listaSeries = new List<object>();
                                var allocs = new List<object>();

                                if (manejaSeries == "Y")
                                {
                                    int serialIndex = 0;
                                    foreach (var stockRow in stockArticulo)
                                    {
                                        double qty = Convert.ToDouble(stockRow["Cantidad_A_Mover"]);
                                        cantidadTotalItem += qty;
                                        listaSeries.Add(new { SystemSerialNumber = Convert.ToInt32(stockRow["Id_Serie_Interno"]), Quantity = qty });

                                        allocs.Add(new { BinAbsEntry = idUbicacionPulmon, Quantity = qty, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = serialIndex, BinActionType = "batFromWarehouse" });
                                        allocs.Add(new { BinAbsEntry = idDestinoFinal, Quantity = qty, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = serialIndex, BinActionType = "batToWarehouse" });
                                        serialIndex++;
                                    }
                                    lineasTransferencia.Add(new { ItemCode = itemCode, Quantity = cantidadTotalItem, WarehouseCode = "ALM01", FromWarehouseCode = "ALM01", SerialNumbers = listaSeries, StockTransferLinesBinAllocations = allocs });
                                }
                                else
                                {
                                    foreach (var stockRow in stockArticulo) { cantidadTotalItem += Convert.ToDouble(stockRow["Cantidad_A_Mover"]); }

                                    allocs.Add(new { BinAbsEntry = idUbicacionPulmon, Quantity = cantidadTotalItem, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = 0, BinActionType = "batFromWarehouse" });
                                    allocs.Add(new { BinAbsEntry = idDestinoFinal, Quantity = cantidadTotalItem, AllowNegativeQuantity = "tNO", SerialAndBatchNumbersBaseLine = 0, BinActionType = "batToWarehouse" });

                                    lineasTransferencia.Add(new { ItemCode = itemCode, Quantity = cantidadTotalItem, WarehouseCode = "ALM01", FromWarehouseCode = "ALM01", StockTransferLinesBinAllocations = allocs });
                                }

                                if (lineasTransferencia.Count > 0)
                                {
                                    const int TAMANO_LOTE = 20;
                                    for (int i = 0; i < lineasTransferencia.Count; i += TAMANO_LOTE)
                                    {
                                        var chunkLineas = lineasTransferencia.Skip(i).Take(TAMANO_LOTE).ToList();

                                        System.Diagnostics.Debug.WriteLine($"> Enviando POST a Service Layer ({itemCode})...");

                                        bool exito = await EnviarTransferencia(chunkLineas, $"Traslado a {nombreDestino}", logPutaway);

                                        System.Diagnostics.Debug.WriteLine($"> Respuesta SAP ({itemCode}): {(exito ? "EXITO" : "ERROR")}");

                                        if (exito)
                                        {
                                            System.Threading.Interlocked.Add(ref lineasTrasladadas, chunkLineas.Count);
                                            System.Threading.Interlocked.Increment(ref documentosExitosos);
                                        }
                                        else
                                        {
                                            System.Threading.Interlocked.Increment(ref erroresPutaway);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[!] EXCEPCIÓN en {itemCode}: {ex.Message}");
                                System.Threading.Interlocked.Increment(ref erroresPutaway);
                                logPutaway.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN AL MOVER {itemCode} | {ex.Message}");
                            }
                            finally
                            {
                                int procesadosLocal = System.Threading.Interlocked.Increment(ref procesadosActuales);
                                Invoke(new Action(() =>
                                {
                                    label3.Text = $"Paso 2: Transfiriendo... ({procesadosLocal} / {filasValidadas.Count}) completados";
                                }));
                                semaphore.Release();
                            }
                        }));
                    }

                    await Task.WhenAll(tareas);
                }

                System.Diagnostics.Debug.WriteLine("=== PUTAWAY FINALIZADO ===");

                if (erroresPutaway > 0)
                {
                    string ruta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_Putaway_WMS_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                    File.WriteAllLines(ruta, logPutaway);
                    MessageBox.Show($"Distribución finalizada con observaciones.\n\nDocumentos Exitosos: {documentosExitosos} | Artículos Trasladados: {lineasTrasladadas}\nErrores: {erroresPutaway}\n\nSe generó un archivo de log en el escritorio.", "WMS Makita - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show($"¡PUTAWAY EXITOSO!\n\nSe distribuyeron masivamente todos los artículos validados hacia el área de Picking.\nTotal traslados generados: {documentosExitosos}", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error crítico durante la reubicación masiva: {ex.Message}", "Fallo de Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button13.Enabled = true;
                button13.Text = "3. Ejecutar Putaway (Reubicar)";
                button12.Enabled = true;
                label3.Text = "Proceso terminado.";
            }
        }

        private async void consultarUbiBD_F3_Click(object sender, EventArgs e)
        {

            try
            {
                consultarUbiBD_F3.Enabled = false;
                consultarUbiBD_F3.Text = "Consultando HANA, espere...";

                dgv_x.DataSource = null;
                dgv_x.Rows.Clear();
                dgv_x.Columns.Clear();

                // Agregamos ROW_NUMBER() para generar la columna "#" dinámicamente desde HANA
                string queryHana = @"
            SELECT 
                ROW_NUMBER() OVER(ORDER BY T0.""BinCode"") AS ""#"",
                T0.""AbsEntry"" AS ""ID Interno"", 
                T0.""BinCode"" AS ""Código Ubicación"", 
                T0.""WhsCode"" AS ""Almacén"", 
                T0.""SL1Code"" AS ""Fila"", 
                T0.""SL2Code"" AS ""Cuerpo"", 
                T0.""SL3Code"" AS ""Piso"", 
                T0.""SL4Code"" AS ""Fondo"", 
                CASE WHEN T0.""Disabled"" = 'Y' THEN 'tYES' ELSE 'tNO' END AS ""Inactivo"",
                T0.""U_WAR_DW_TYPE_LOCATION"" AS ""Tipo (UDF)""
            FROM OBIN T0
            WHERE T0.""WhsCode"" = 'ALM01'
            ORDER BY T0.""BinCode"";";

                DataTable dtUbicaciones = await Task.Run(() =>
                {
                    DataTable dt = new DataTable();
                    using (var conn = new System.Data.Odbc.OdbcConnection(GetHanaConnectionString()))
                    {
                        conn.Open();
                        using (var cmd = new System.Data.Odbc.OdbcCommand(queryHana, conn))
                        using (var da = new System.Data.Odbc.OdbcDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                    return dt;
                });

                dgv_x.DataSource = dtUbicaciones;

                // Ajustes visuales de las columnas
                if (dgv_x.Columns.Contains("#")) dgv_x.Columns["#"].Width = 40;
                if (dgv_x.Columns.Contains("ID Interno")) dgv_x.Columns["ID Interno"].Visible = false;
                if (dgv_x.Columns.Contains("Código Ubicación")) dgv_x.Columns["Código Ubicación"].Width = 150;

                MessageBox.Show($"¡Éxito! Se cargaron {dtUbicaciones.Rows.Count} ubicaciones de la base de datos.", "WMS Makita - Fase 3", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error al consultar HANA por ODBC:\n\n{ex.Message}", "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                consultarUbiBD_F3.Enabled = true;
                consultarUbiBD_F3.Text = "Consultar Ubicación BD";
            }


        }

        private void importExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Seleccione el Excel de Zonas de Reserva";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    importExcel.Enabled = false;
                    importExcel.Text = "Cargando DATAMASTER...";

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });

                            // Buscamos explícitamente la hoja llamada DATAMASTER
                            DataTable dtExcel = result.Tables["DATAMASTER"];




                            if (dtExcel != null)
                            {
                                int totalOriginal = dtExcel.Rows.Count;

                                var filasUnicas = dtExcel.AsEnumerable()
                                    .GroupBy(row => new
                                    {
                                        Almacen = row[0]?.ToString().Trim(),
                                        Fila = row[1]?.ToString().Trim(),
                                        Cuerpo = row[2]?.ToString().Trim(),
                                        Piso = row[3]?.ToString().Trim(),
                                        Fondo = row[4]?.ToString().Trim()
                                    }).Select(grupo => grupo.First()).ToList();

                                int totalUnicos = filasUnicas.Count;
                                int duplicados = totalOriginal - totalUnicos;

                                // Agregamos la columna "#" al principio
                                dgv_y.Columns.Clear();
                                dgv_y.Columns.Add("Nro", "#");
                                dgv_y.Columns.Add("Almacen", "Almacén");
                                dgv_y.Columns.Add("Sub1", "Fila");
                                dgv_y.Columns.Add("Sub2", "Cuerpo");
                                dgv_y.Columns.Add("Sub3", "Piso");
                                dgv_y.Columns.Add("Sub4", "Fondo");

                                dgv_y.Columns["Nro"].Width = 40;

                                int countValidos = 0;
                                foreach (var row in filasUnicas)
                                {
                                    string whs = row[0]?.ToString().Trim();
                                    string fila = row[1]?.ToString().Trim();
                                    string cuerpo = row[2]?.ToString().Trim();
                                    string piso = row[3]?.ToString().Trim();
                                    string fondo = row[4]?.ToString().Trim();

                                    if (!string.IsNullOrEmpty(fila))
                                    {
                                        countValidos++;
                                        // Inyectamos el contador en la primera posición
                                        dgv_y.Rows.Add(countValidos, whs, fila, cuerpo, piso, fondo);
                                    }
                                }




                                MessageBox.Show($"¡Layout de Reserva importado con éxito!\n\n" +
                                                $"• Filas leídas: {totalOriginal}\n" +
                                                $"• Duplicados omitidos: {duplicados}\n" +
                                                $"• Ubicaciones listas para cruzar: {countValidos}",
                                                "WMS Makita - Fase 3", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("No se encontró una hoja llamada 'DATAMASTER' en el Excel.",
                                                "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Fallo de Importación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // Restauramos el botón
                    importExcel.Enabled = true;
                    importExcel.Text = "Importar/Leer Excel Ubicacion por Crear";
                }
            }




        }

        private async void vali_crear_subniveles_Click(object sender, EventArgs e)
        {
            if (dgv_y.Rows.Count == 0)
            {
                MessageBox.Show("Por favor, importe primero el Excel de Zonas de Reserva.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            vali_crear_subniveles.Enabled = false;
            vali_crear_subniveles.Text = "Sincronizando Estructura...";
            label4.Text = "Analizando subniveles únicos...";

            // 1. Extraemos solo valores ÚNICOS del Excel importado (dgv_y)
            HashSet<string> filasUnicas = new HashSet<string>();
            HashSet<string> cuerposUnicos = new HashSet<string>();
            HashSet<string> pisosUnicos = new HashSet<string>();
            HashSet<string> fondosUnicos = new HashSet<string>();

            foreach (DataGridViewRow row in dgv_y.Rows)
            {
                if (row.IsNewRow) continue;

                string fila = row.Cells["Sub1"].Value?.ToString().Trim();
                string cuerpo = row.Cells["Sub2"].Value?.ToString().Trim();
                string piso = row.Cells["Sub3"].Value?.ToString().Trim();
                string fondo = row.Cells["Sub4"].Value?.ToString().Trim();

                if (!string.IsNullOrEmpty(fila)) filasUnicas.Add(fila);
                if (!string.IsNullOrEmpty(cuerpo)) cuerposUnicos.Add(cuerpo);
                if (!string.IsNullOrEmpty(piso)) pisosUnicos.Add(piso);
                if (!string.IsNullOrEmpty(fondo)) fondosUnicos.Add(fondo);
            }

            int creadosTotales = 0;
            int omitidosTotales = 0;

            label4.Text = "Inyectando estructura en SAP...";

            // 2. Ejecutamos la sincronización llamando al método externo que ya tienes creado
            int creadosFila = await EnviarSubnivelesASap(filasUnicas, 1, "Fila");
            creadosTotales += creadosFila;
            omitidosTotales += (filasUnicas.Count - creadosFila);

            int creadosCuerpo = await EnviarSubnivelesASap(cuerposUnicos, 2, "Cuerpo");
            creadosTotales += creadosCuerpo;
            omitidosTotales += (cuerposUnicos.Count - creadosCuerpo);

            int creadosPiso = await EnviarSubnivelesASap(pisosUnicos, 3, "Piso");
            creadosTotales += creadosPiso;
            omitidosTotales += (pisosUnicos.Count - creadosPiso);

            int creadosFondo = await EnviarSubnivelesASap(fondosUnicos, 4, "Fondo");
            creadosTotales += creadosFondo;
            omitidosTotales += (fondosUnicos.Count - creadosFondo);

            // 3. Reactivamos Interfaz
            vali_crear_subniveles.Enabled = true;
            vali_crear_subniveles.Text = "Validar crear Subniveles";
            label4.Text = "Estructura sincronizada. Listo para crear.";

            MessageBox.Show($"Validación de Subniveles Finalizada.\n\nNuevos creados: {creadosTotales}\nYa existentes (Omitidos): {omitidosTotales}\n\n¡Ya puedes ejecutar la creación de ubicaciones!",
                            "WMS Makita - Preparación", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        private async void crear_ubicacion_Click(object sender, EventArgs e)
        {
            // 1. Validaciones de seguridad
            if (dgv_y.Rows.Count == 0 || (dgv_y.Rows.Count == 1 && dgv_y.Rows[0].IsNewRow))
            {
                MessageBox.Show("Debe importar el layout de Reserva desde el Excel antes de ejecutar.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int filasHana = dgv_x.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
            if (filasHana == 0)
            {
                DialogResult respuesta = MessageBox.Show(
                    "La consulta de Base de Datos está vacía.\n¿Deseas continuar asumiendo que el almacén es 100% nuevo?",
                    "WMS Makita - Control", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (respuesta == DialogResult.No) return;
            }

            crear_ubicacion.Enabled = false;
            crear_ubicacion.Text = "Construyendo/Actualizando Layout...";
            vali_crear_subniveles.Enabled = false;

            int creadas = 0;
            int actualizadas = 0;
            int errores = 0;
            int procesadasActuales = 0;
            System.Collections.Concurrent.ConcurrentBag<string> logFase3 = new System.Collections.Concurrent.ConcurrentBag<string>();

            // 2. Cargamos HANA en un Diccionario para tener el AbsEntry (necesario para el PATCH)
            label4.Text = "Cruzando datos en memoria RAM...";
            var existentesHana = new Dictionary<string, string>();
            foreach (DataGridViewRow row in dgv_x.Rows)
            {
                if (row.IsNewRow) continue;
                string binCode = row.Cells["Código Ubicación"].Value?.ToString();
                string absEntry = row.Cells["ID Interno"].Value?.ToString();

                if (!string.IsNullOrEmpty(binCode))
                {
                    existentesHana[binCode.Replace("-", "").Trim()] = absEntry;
                }
            }

            // Calculamos el total exacto del Excel
            int totalAProcesar = dgv_y.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
            progressBarF3.Minimum = 0;
            progressBarF3.Maximum = totalAProcesar;
            progressBarF3.Value = 0;
            progressBarF3.Visible = true;

            // 3. Procesamiento dual (UPDATE / CREATE) en bloques
            using (SemaphoreSlim semaphore = new SemaphoreSlim(15))
            {
                var tareas = new List<Task>();

                foreach (DataGridViewRow row in dgv_y.Rows)
                {
                    if (row.IsNewRow) continue;

                    string whs = row.Cells["Almacen"].Value?.ToString().Trim() ?? "";
                    string fila = row.Cells["Sub1"].Value?.ToString().Trim() ?? "";
                    string cuerpo = row.Cells["Sub2"].Value?.ToString().Trim() ?? "";
                    string piso = row.Cells["Sub3"].Value?.ToString().Trim() ?? "";
                    string fondo = row.Cells["Sub4"].Value?.ToString().Trim() ?? "";

                    string cleanBinCode = $"{whs}{fila}{cuerpo}{piso}{fondo}";

                    // Protegemos el pulmón
                    if (cleanBinCode.Contains("UBICACIÓN-DE-SISTEMA") || cleanBinCode.Contains("SYSTEM")) continue;

                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            if (existentesHana.ContainsKey(cleanBinCode))
                            {
                                // === ACTUALIZAR (PATCH) ===
                                string entry = existentesHana[cleanBinCode];
                                string jsonPatch = $@"{{ ""Inactive"": ""tNO"", ""U_WAR_DW_TYPE_LOCATION"": ""3"" }}";

                                var content = new StringContent(jsonPatch, System.Text.Encoding.UTF8, "application/json");
                                var response = await client.PatchAsync($"BinLocations({entry})", content);

                                if (response.IsSuccessStatusCode)
                                {
                                    System.Threading.Interlocked.Increment(ref actualizadas);
                                }
                                else
                                {
                                    string err = await response.Content.ReadAsStringAsync();
                                    logFase3.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL ACTUALIZAR {cleanBinCode} | {err}");
                                    System.Threading.Interlocked.Increment(ref errores);
                                }
                            }
                            else
                            {
                                // === CREAR NUEVA (POST) ===
                                string jsonPost = $@"{{
                            ""Warehouse"": ""{whs}"",
                            ""Sublevel1"": ""{fila}"",
                            ""Sublevel2"": ""{cuerpo}"",
                            ""Sublevel3"": ""{piso}"",
                            ""Sublevel4"": ""{fondo}"",
                            ""Inactive"": ""tNO"",
                            ""U_WAR_DW_TYPE_LOCATION"": ""3""
                        }}";

                                var content = new StringContent(jsonPost, System.Text.Encoding.UTF8, "application/json");
                                var response = await client.PostAsync("BinLocations", content);

                                if (response.IsSuccessStatusCode)
                                {
                                    System.Threading.Interlocked.Increment(ref creadas);
                                }
                                else
                                {
                                    string err = await response.Content.ReadAsStringAsync();
                                    logFase3.Add($"[{DateTime.Now:HH:mm:ss}] ERROR AL CREAR {cleanBinCode} | {err}");
                                    System.Threading.Interlocked.Increment(ref errores);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logFase3.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN EN {cleanBinCode} | {ex.Message}");
                            System.Threading.Interlocked.Increment(ref errores);
                        }
                        finally
                        {
                            int current = System.Threading.Interlocked.Increment(ref procesadasActuales);
                            int porcentaje = (int)((double)current / totalAProcesar * 100);

                            Invoke(new Action(() =>
                            {
                                progressBarF3.Value = current > progressBarF3.Maximum ? progressBarF3.Maximum : current;
                                label4.Text = $"Procesando zonas de reserva... {porcentaje}% ({current} de {totalAProcesar})";
                            }));
                            semaphore.Release();
                        }
                    }));
                }

                await Task.WhenAll(tareas);
            }

            progressBarF3.Value = progressBarF3.Maximum;
            label4.Text = "¡Proceso finalizado!";
            RestaurarBotonesFase3();

            if (errores > 0)
            {
                string rutaLog = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_Fase3_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(rutaLog, logFase3);
                MessageBox.Show($"Proceso masivo finalizado con observaciones.\n\nNuevas Reservas Creadas: {creadas}\nReservas Actualizadas: {actualizadas}\nErrores: {errores}\n\nRevisa el log en tu escritorio.", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show($"¡Fase 3 Completada Exitosamente!\n\nSe procesaron las {totalAProcesar} operaciones:\n• Nuevas Creadas: {creadas}\n• Existentes Actualizadas: {actualizadas}", "WMS Makita", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        // Método auxiliar para reactivar la UI
        private void RestaurarBotonesFase3()
        {
            crear_ubicacion.Enabled = true;
            crear_ubicacion.Text = "Crear Ubicaciones";
            vali_crear_subniveles.Enabled = true;
        }

        private async void f4_importarExcel_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    f4_importarExcel.Enabled = false;
                    label5.Text = "Leyendo Excel masivo...";
                    f4_progressBar.Style = ProgressBarStyle.Marquee;
                    f4_progressBar.Visible = true;

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    await Task.Run(() =>
                    {
                        using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                                });

                                DataTable rawData = result.Tables[0];
                                dtExcelFase4Completo = new DataTable();
                                dtExcelFase4Completo.Columns.Add("ItemCode", typeof(string));
                                dtExcelFase4Completo.Columns.Add("SerieSAP", typeof(string));
                                dtExcelFase4Completo.Columns.Add("SerieFisico", typeof(string));
                                dtExcelFase4Completo.Columns.Add("ComentarioExcel", typeof(string));
                                dtExcelFase4Completo.Columns.Add("EstadoValidacion", typeof(string));
                                dtExcelFase4Completo.Columns.Add("SysNumber", typeof(int));

                                foreach (DataRow row in rawData.Rows)
                                {
                                    dtExcelFase4Completo.Rows.Add(
                                row[1]?.ToString().Trim(), // Ahora el Artículo está en la Columna B (índice 1)
                                row[2]?.ToString().Trim(), // Ahora la SerieSAP está en la Columna C (índice 2)
                                row[3]?.ToString().Trim(), // Ahora la SerieFisico está en la Columna D (índice 3)
                                row[4]?.ToString().Trim(), // Ahora el Comentario está en la Columna E (índice 4)
                                "Pendiente",
                                0
                                );
                                }
                            }
                        }
                    });

                    // DataBinding masivo ultra rápido
                    f4_dgv.DataSource = dtExcelFase4Completo;
                    f4_dgv.Columns["SysNumber"].Visible = false;
                    f4_dgv.Columns["EstadoValidacion"].Width = 250;

                    f4_progressBar.Style = ProgressBarStyle.Blocks;
                    f4_progressBar.Value = 100;
                    label5.Text = $"Registros importados: {dtExcelFase4Completo.Rows.Count:N0}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Fallo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    f4_importarExcel.Enabled = true;
                    f4_progressBar.Visible = false;
                }
            }




        }



        private async void f4_prevalidacion_Click(object sender, EventArgs e)
        {
            if (dtExcelFase4Completo == null || dtExcelFase4Completo.Rows.Count == 0) return;

            f4_prevalidacion.Enabled = false;
            label5.Text = "Descargando maestro de series y auditando...";
            f4_progressBar.Style = ProgressBarStyle.Marquee;
            f4_progressBar.Visible = true;

            var dicSeriesSAP = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            //string querySeries = @"SELECT ""ItemCode"", ""DistNumber"", ""SysNumber"" FROM OSRN";


            // AHORA (Traemos el AbsEntry):
            string querySeries = @"SELECT ""ItemCode"", ""DistNumber"", ""AbsEntry"" FROM OSRN";


            int listosParaActualizar = 0;
            int erroresDetectados = 0;
            int falsosOkDetectados = 0;
            int seriesOkReales = 0; // <-- VARIABLE AGREGADA PARA SOLUCIONAR EL ERROR

            try
            {
                await Task.Run(() =>
                {
                    // 1. Cargar HANA
                    using (var conn = new System.Data.Odbc.OdbcConnection(GetHanaConnectionString()))
                    {
                        conn.Open();
                        using (var cmd = new System.Data.Odbc.OdbcCommand(querySeries, conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string llaveUnica = $"{reader.GetString(0)}|{reader.GetString(1)}";
                                dicSeriesSAP[llaveUnica] = reader.GetInt32(2);
                            }
                        }
                    }

                    // 2. Procesar DataTable directamente en RAM
                    foreach (DataRow row in dtExcelFase4Completo.Rows)
                    {
                        string itemCode = row["ItemCode"].ToString();
                        string serieSap = row["SerieSAP"].ToString();
                        string serieFisico = row["SerieFisico"].ToString();
                        string comentario = row["ComentarioExcel"].ToString().ToUpper();

                        if (string.IsNullOrEmpty(itemCode)) continue;

                        if (comentario == "CAMBIAR SERIE")
                        {
                            if (!dicSeriesSAP.ContainsKey($"{itemCode}|{serieSap}"))
                            {
                                row["EstadoValidacion"] = "ERROR: Origen no existe";
                                erroresDetectados++;
                            }
                            else if (dicSeriesSAP.ContainsKey($"{itemCode}|{serieFisico}"))
                            {
                                row["EstadoValidacion"] = "ERROR: Destino ya existe";
                                erroresDetectados++;
                            }
                            else
                            {
                                row["EstadoValidacion"] = "OK - Listo para Actualizar";
                                row["SysNumber"] = dicSeriesSAP[$"{itemCode}|{serieSap}"];
                                listosParaActualizar++;
                            }
                        }
                        else if (comentario == "SERIE OK")
                        {
                            if (dicSeriesSAP.ContainsKey($"{itemCode}|{serieFisico}"))
                            {
                                row["EstadoValidacion"] = "OK - Verificado en SAP";
                                seriesOkReales++; // EL CONTADOR AHORA FUNCIONA
                            }
                            else
                            {
                                row["EstadoValidacion"] = "ALERTA: Falso OK";
                                falsosOkDetectados++;
                            }
                        }
                    }
                });

                // 3. Forzar refresco visual masivo (DataBinding)
                f4_dgv.Refresh();

                // Opcional: Colorear la grilla post-proceso
                foreach (DataGridViewRow row in f4_dgv.Rows)
                {
                    string estado = row.Cells["EstadoValidacion"].Value?.ToString() ?? "";
                    if (estado.StartsWith("OK - L")) row.DefaultCellStyle.BackColor = Color.LightGreen;
                    else if (estado.StartsWith("ALERTA")) row.DefaultCellStyle.BackColor = Color.Orange;
                    else if (estado.StartsWith("ERROR")) row.DefaultCellStyle.BackColor = Color.LightCoral;
                }

                f4_progressBar.Style = ProgressBarStyle.Blocks;
                f4_progressBar.Value = 100;
                label5.Text = $"Listos: {listosParaActualizar} | Falsos OK: {falsosOkDetectados} | Errores: {erroresDetectados}";

                // 4. RESUMEN EMERGENTE RESTAURADO
                MessageBox.Show($"¡Auditoría Completa de Inventario Finalizada!\n\n" +
                                $"• Series listas para corregir (PATCH): {listosParaActualizar}\n" +
                                $"• Series 'OK' reales confirmadas en SAP: {seriesOkReales}\n" +
                                $"• Falsos 'SERIE OK' (Inventario fantasma): {falsosOkDetectados}\n" +
                                $"• Errores lógicos detectados: {erroresDetectados}",
                                "WMS Makita - Reporte de Auditoría", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Fallo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                f4_prevalidacion.Enabled = true;
                f4_progressBar.Visible = false;
            }
        }





        private async void f4_exportarlog_Click(object sender, EventArgs e)
        {


            if (dtExcelFase4Completo == null || dtExcelFase4Completo.Rows.Count == 0) return;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text File|*.txt";
            sfd.FileName = $"Log_Fase4_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    f4_exportarlog.Text = "Exportando...";
                    f4_exportarlog.Enabled = false;

                    await Task.Run(() =>
                    {
                        using (StreamWriter sw = new StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8))
                        {
                            sw.WriteLine("ItemCode\tSerieSAP\tSerieFisico\tEstadoValidacion");
                            foreach (DataRow row in dtExcelFase4Completo.Rows)
                            {
                                sw.WriteLine($"{row["ItemCode"]}\t{row["SerieSAP"]}\t{row["SerieFisico"]}\t{row["EstadoValidacion"]}");
                            }
                        }
                    });

                    MessageBox.Show("Log exportado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    f4_exportarlog.Text = "Exportar_log";
                    f4_exportarlog.Enabled = true;
                }
            }







        }

        private async void f4_actualizarmasiva_Click(object sender, EventArgs e)
        {


            // 1. Extraer datos limpios antes de abrir hilos
            var listaAProcesar = dtExcelFase4Completo.AsEnumerable()
                .Where(r => r.Field<string>("EstadoValidacion") == "OK - Listo para Actualizar")
                .Select(r => new {
                    ItemCode = r.Field<string>("ItemCode"),
                    SerieFisico = r.Field<string>("SerieFisico"),
                    SysNumber = r.Field<int>("SysNumber"),
                    Fila = r
                }).ToList();

            if (listaAProcesar.Count == 0) return;

            f4_actualizarmasiva.Enabled = false;
            f4_progressBar.Visible = true;
            f4_progressBar.Maximum = listaAProcesar.Count;
            f4_progressBar.Value = 0;

            int actualizadas = 0, errores = 0, procesadas = 0;
            var logPatch = new System.Collections.Concurrent.ConcurrentBag<string>();

            using (SemaphoreSlim semaphore = new SemaphoreSlim(15))
            {
                var tareas = new List<Task>();

                foreach (var item in listaAProcesar)
                {
                    await semaphore.WaitAsync();

                    tareas.Add(Task.Run(async () =>
                    {
                        try
                        {
                            //string endpoint = $"SerialNumbers(ItemCode='{item.ItemCode}',SystemSerialNumber={item.SysNumber})";

                            //string endpoint = $"SerialNumbers(ItemCode='{item.ItemCode}',SystemSerialNumber={item.SysNumber})";
                            //string endpoint = $"SerialNumberDetails(ItemCode='{Uri.EscapeDataString(item.ItemCode)}',SystemSerialNumber={item.SysNumber})";
                            string endpoint = $"SerialNumberDetails({item.SysNumber})";




                            //string jsonPatch = $@"{{ ""InternalSerialNumber"": ""{item.SerieFisico}"" }}";
                            //string jsonPatch = $@"{{ ""MfrSerialNo"": ""{item.SerieFisico}"" }}";
                            string jsonPatch = $@"{{ ""SerialNumber"": ""{item.SerieFisico}"" }}";




                            var content = new StringContent(jsonPatch, System.Text.Encoding.UTF8, "application/json");

                            var response = await client.PatchAsync(endpoint, content);

                            if (response.IsSuccessStatusCode)
                            {
                                System.Threading.Interlocked.Increment(ref actualizadas);
                                item.Fila["EstadoValidacion"] = "ACTUALIZADO EN SAP"; // Modificación segura en RAM
                            }
                            else
                            {
                                string err = await response.Content.ReadAsStringAsync();
                                logPatch.Add($"[{DateTime.Now:HH:mm:ss}] ERROR {item.ItemCode}: {err}");
                                System.Threading.Interlocked.Increment(ref errores);
                            }
                        }
                        catch (Exception ex)
                        {
                            logPatch.Add($"[{DateTime.Now:HH:mm:ss}] EXCEPCIÓN {item.ItemCode}: {ex.Message}");
                            System.Threading.Interlocked.Increment(ref errores);
                        }
                        finally
                        {
                            int current = System.Threading.Interlocked.Increment(ref procesadas);
                            if (current % 10 == 0 || current == f4_progressBar.Maximum) // Evita saturar la UI
                            {
                                Invoke(new Action(() => {
                                    f4_progressBar.Value = current;
                                    label5.Text = $"Actualizando: {current} de {listaAProcesar.Count}";
                                }));
                            }
                            semaphore.Release();
                        }
                    }));
                }
                await Task.WhenAll(tareas);
            }

            f4_dgv.Refresh(); // Refleja el texto "ACTUALIZADO EN SAP" visualmente al terminar
            f4_progressBar.Visible = false;
            f4_actualizarmasiva.Enabled = true;

            // --- CÓDIGO RESTAURADO PARA GUARDAR EL LOG EN EL ESCRITORIO ---
            if (errores > 0)
            {
                string rutaLog = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Log_ErrorPATCH_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(rutaLog, logPatch);

                MessageBox.Show($"Actualización rechazada por Service Layer.\n\nÉxitos: {actualizadas}\nFallos: {errores}\n\nSe ha creado un reporte técnico en tu Escritorio.\nRevisa el archivo para ver la respuesta exacta de SAP.",
                                "WMS Makita - Alerta de SAP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show($"¡Sinceramiento de Series Finalizado!\n\nSe corrigieron exitosamente {actualizadas} series en SAP.",
                                "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {

        }
    }
}
