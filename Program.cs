using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniDbServer
{
    // ----------------------------------------------------------------------
    // Modello del record salvato nel database Access
    // ----------------------------------------------------------------------
    public class Record
    {
        public int Id { get; set; }
        public string Nome { get; set; } = "";
        public string Descrizione { get; set; } = "";
    }

    // Corpo della richiesta per l'endpoint di query SQL libera
    public class QueryRequest
    {
        public string Sql { get; set; } = "";
    }

    // Corpo della richiesta per l'endpoint di transazione: una sequenza di operazioni
    // SQL qualsiasi (SELECT, INSERT, UPDATE, DELETE, DDL...) da eseguire tutte insieme
    // o annullare tutte insieme.
    public class TransactionRequest
    {
        public List<string> Operazioni { get; set; } = new();
    }

    // ----------------------------------------------------------------------
    // Accesso al database Access (.accdb) tramite OleDb
    // ----------------------------------------------------------------------
    public class AccessDb
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private const string TableName = "Records";

        // Timeout minimo per l'esecuzione di ogni comando SQL: 3 minuti (180 secondi).
        // Il default di OleDbCommand sarebbe 30 secondi, troppo poco per query lunghe
        // su tabelle grandi: viene impostato esplicitamente su ogni comando creato.
        private const int TimeoutSecondi = 180;

        // Percorso del file .accdb/.mdb a cui questa istanza è collegata.
        public string DbPath => _dbPath;

        public AccessDb(string dbPath)
        {
            _dbPath = dbPath;
            _connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_dbPath};";
        }

        private OleDbConnection GetConnection() => new OleDbConnection(_connectionString);

        // Crea un OleDbCommand già impostato con il timeout esteso di default.
        private static OleDbCommand CreaComando(string sql, OleDbConnection conn, OleDbTransaction? tx = null)
        {
            var cmd = new OleDbCommand(sql, conn, tx);
            cmd.CommandTimeout = TimeoutSecondi;
            return cmd;
        }

        // Crea il file .accdb e la tabella se non esistono già.
        // 'log' è opzionale: se fornito, riceve i messaggi di avanzamento
        // (usato per mostrarli nella finestra dell'applicazione).
        public void EnsureDatabase(Action<string>? log = null)
        {
            if (!File.Exists(_dbPath))
            {
                log?.Invoke($"Il file '{_dbPath}' non esiste: provo a crearlo...");
                try
                {
                    // Crea il file .accdb vuoto usando ADOX (richiede Access Database Engine installato)
                    Type? catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
                        ?? throw new InvalidOperationException("ADOX non disponibile su questo sistema.");
                    dynamic catalog = Activator.CreateInstance(catalogType)!;
                    catalog.Create(_connectionString);
                    log?.Invoke("File .accdb creato correttamente.");
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        "Impossibile creare automaticamente il database Access. " +
                        "Crea manualmente un file .accdb vuoto con questo nome/percorso, oppure installa " +
                        "il 'Microsoft Access Database Engine 2016 Redistributable'. Dettaglio errore: " + ex.Message);
                }
            }

            // Crea la tabella se manca
            using var conn = GetConnection();
            conn.Open();

            bool tableExists = false;
            var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object?[] { null, null, TableName, "TABLE" });
            if (schema != null && schema.Rows.Count > 0) tableExists = true;

            if (!tableExists)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandTimeout = TimeoutSecondi;
                cmd.CommandText = $@"CREATE TABLE {TableName} (
                    ID AUTOINCREMENT PRIMARY KEY,
                    Nome TEXT(255),
                    Descrizione TEXT(255)
                )";
                cmd.ExecuteNonQuery();
                log?.Invoke($"Tabella '{TableName}' creata.");
            }
        }

        public List<Record> GetAll()
        {
            var list = new List<Record>();
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"SELECT ID, Nome, Descrizione FROM {TableName} ORDER BY ID", conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new Record
                {
                    Id = reader.GetInt32(0),
                    Nome = reader.IsDBNull(1) ? "" : reader.GetString(1),
                    Descrizione = reader.IsDBNull(2) ? "" : reader.GetString(2)
                });
            }
            return list;
        }

        public Record? GetById(int id)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"SELECT ID, Nome, Descrizione FROM {TableName} WHERE ID = ?", conn);
            cmd.Parameters.AddWithValue("@id", id);
            using var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                return new Record
                {
                    Id = reader.GetInt32(0),
                    Nome = reader.IsDBNull(1) ? "" : reader.GetString(1),
                    Descrizione = reader.IsDBNull(2) ? "" : reader.GetString(2)
                };
            }
            return null;
        }

        public int Insert(Record r)
        {
            using var conn = GetConnection();
            conn.Open();
            using (var cmd = CreaComando($"INSERT INTO {TableName} (Nome, Descrizione) VALUES (?, ?)", conn))
            {
                cmd.Parameters.AddWithValue("@nome", r.Nome ?? "");
                cmd.Parameters.AddWithValue("@descrizione", r.Descrizione ?? "");
                cmd.ExecuteNonQuery();
            }
            using (var idCmd = CreaComando("SELECT @@IDENTITY", conn))
            {
                var result = idCmd.ExecuteScalar();
                return Convert.ToInt32(result);
            }
        }

        public bool Update(int id, Record r)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"UPDATE {TableName} SET Nome = ?, Descrizione = ? WHERE ID = ?", conn);
            cmd.Parameters.AddWithValue("@nome", r.Nome ?? "");
            cmd.Parameters.AddWithValue("@descrizione", r.Descrizione ?? "");
            cmd.Parameters.AddWithValue("@id", id);
            return cmd.ExecuteNonQuery() > 0;
        }

        public bool Delete(int id)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando($"DELETE FROM {TableName} WHERE ID = ?", conn);
            cmd.Parameters.AddWithValue("@id", id);
            return cmd.ExecuteNonQuery() > 0;
        }

        // Esegue una qualsiasi query SQL scritta dall'utente:
        // - se è una SELECT, restituisce le righe trovate
        // - altrimenti (INSERT/UPDATE/DELETE/CREATE TABLE/ALTER/...) la esegue come comando
        //   e restituisce il numero di righe interessate.
        // Pensato per un uso locale/offline: non applica filtri di sicurezza sul testo SQL.
        public object ExecuteRawQuery(string sql)
        {
            using var conn = GetConnection();
            conn.Open();
            using var cmd = CreaComando(sql, conn);

            bool isSelect = sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase);

            if (isSelect)
            {
                using var reader = cmd.ExecuteReader();
                var righe = new List<Dictionary<string, object?>>();
                while (reader.Read())
                {
                    var riga = new Dictionary<string, object?>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        riga[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    }
                    righe.Add(riga);
                }
                return righe;
            }
            else
            {
        // Esegue una SEQUENZA di operazioni SQL (di qualsiasi tipo: SELECT, INSERT, UPDATE,
        // DELETE, CREATE TABLE, ecc.) dentro un'unica transazione:
        // - se TUTTE le operazioni vanno a buon fine, viene fatto un solo commit finale;
        // - se anche UNA sola operazione fallisce, tutte le altre già eseguite in questa
        //   chiamata vengono annullate (rollback) e il database torna com'era prima.
        // Questo evita che una sequenza (es. "svuota tabella" + tanti "inserisci riga")
        // possa interrompersi a metà lasciando i dati in uno stato incoerente.
        public List<object> EseguiTransazione(List<string> operazioni)
        {
            using var conn = GetConnection();
            conn.Open();
            using var tx = conn.BeginTransaction();

            var risultati = new List<object>();
            try
            {
                foreach (var sql in operazioni)
                {
                    if (string.IsNullOrWhiteSpace(sql)) continue;

                    using var cmd = CreaComando(sql, conn, tx);
                    bool isSelect = sql.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase);

                    if (isSelect)
                    {
                        using var reader = cmd.ExecuteReader();
                        var righe = new List<Dictionary<string, object?>>();
                        while (reader.Read())
                        {
                            var riga = new Dictionary<string, object?>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                riga[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                            }
                            righe.Add(riga);
                        }
                        risultati.Add(righe);
                    }
                    else
                    {
                        int righeInteressate = cmd.ExecuteNonQuery();
                        risultati.Add(new { righeInteressate });
                    }
                }

                tx.Commit();
                return risultati;
            }
            catch
            {
                // Anche solo un errore su una qualsiasi operazione annulla TUTTO
                // quello già eseguito in questa transazione.
                try { tx.Rollback(); } catch { /* connessione già chiusa/non valida */ }
                throw;
            }
        }
    }

    // Corpo della richiesta per selezionare quale database .accdb/.mdb usare
    public class DatabaseSelectRequest
    {
        public string Percorso { get; set; } = "";
        public bool Crea { get; set; } = false; // se true, crea il file (con tabella Records) se non esiste
    }

    // ----------------------------------------------------------------------
    // Piccola finestra dell'applicazione: mostra il log del server e un pulsante
    // per aprire la pagina web nel browser predefinito.
    // ----------------------------------------------------------------------
    public class MainForm : Form
    {
        private readonly TextBox _logTextBox;
        private readonly Button _btnApriBrowser;

        public MainForm()
        {
            Text = "MiniDbServer";
            Width = 560;
            Height = 400;
            MinimumSize = new Size(400, 250);
            StartPosition = FormStartPosition.CenterScreen;

            _logTextBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.Black,
                ForeColor = Color.FromArgb(140, 255, 140),
                BorderStyle = BorderStyle.FixedSingle
            };

            _btnApriBrowser = new Button
            {
                Text = "Apri nel browser",
                Dock = DockStyle.Bottom,
                Height = 42,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            _btnApriBrowser.Click += (s, e) => ApriBrowser();

            Controls.Add(_logTextBox);
            Controls.Add(_btnApriBrowser);
        }

        private void ApriBrowser()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "http://localhost:5000/home.htm",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AggiungiLog("Impossibile aprire il browser: " + ex.Message);
            }
        }

        // Aggiunge una riga al log. Sicuro da chiamare anche da thread diversi
        // da quello dell'interfaccia grafica (es. dal thread del server web).
        public void AggiungiLog(string messaggio)
        {
            string riga = $"[{DateTime.Now:HH:mm:ss}] {messaggio}{Environment.NewLine}";

            if (_logTextBox.IsHandleCreated && _logTextBox.InvokeRequired)
            {
                try { _logTextBox.Invoke(new Action(() => ScriviRiga(riga))); }
                catch (ObjectDisposedException) { /* finestra già chiusa */ }
                catch (InvalidOperationException) { /* handle non ancora pronto/finestra in chiusura */ }
            }
            else
            {
                ScriviRiga(riga);
            }
        }

        private void ScriviRiga(string riga)
        {
            _logTextBox.AppendText(riga);
        }
    }

    // ----------------------------------------------------------------------
    // Server HTTP incorporato
    // ----------------------------------------------------------------------
    public class Program
    {
        private static AccessDb _db = null!;
        private static string _exeDir = "";
        private static readonly object _dbLock = new();
        private static MainForm? _mainForm;

        [STAThread]
        public static void Main(string[] args)
        {
            ApplicationConfiguration.Initialize();

            _mainForm = new MainForm();

            // Il server web gira su un thread separato, così non blocca la finestra grafica.
            _ = Task.Run(() => AvviaServerAsync(args));

            Application.Run(_mainForm);
        }

        // Scrive un messaggio sia in console (utile se lanciato da terminale/debug)
        // sia nella casella di log della finestra grafica.
        private static void Log(string messaggio)
        {
            Console.WriteLine(messaggio);
            _mainForm?.AggiungiLog(messaggio);
        }

        private static async Task AvviaServerAsync(string[] args)
        {
            // Percorso del database: passalo come argomento, oppure viene usato "dati.accdb"
            // nella stessa cartella dell'eseguibile.
            _exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = args.Length > 0 ? args[0] : Path.Combine(_exeDir, "dati.accdb");

            Log("=== MiniDbServer ===");
            Log($"Database: {dbPath}");

            _db = new AccessDb(dbPath);

            try
            {
                _db.EnsureDatabase(Log);
            }
            catch (Exception ex)
            {
                Log("ERRORE inizializzazione database:");
                Log(ex.Message);
                return;
            }

            string prefix = "http://localhost:5000/";
            var listener = new HttpListener();
            listener.Prefixes.Add(prefix);

            try
            {
                listener.Start();
            }
            catch (HttpListenerException ex)
            {
                Log("Impossibile avviare il server: " + ex.Message);
                Log("Prova ad avviare l'exe come amministratore, oppure verifica che la porta 5000 sia libera.");
                return;
            }

            Log($"Server avviato su {prefix}");
            Log("Usa il pulsante 'Apri nel browser' per vedere la pagina, oppure vai su http://localhost:5000/home.htm");

            while (true)
            {
                var context = await listener.GetContextAsync();
                _ = HandleRequestAsync(context); // gestisce ogni richiesta in background
            }
        }

        private static async Task HandleRequestAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            try
            {
                string path = req.Url?.AbsolutePath ?? "/";
                string method = req.HttpMethod;

                Log($"{method} {path}");

                res.Headers.Add("Access-Control-Allow-Origin", "*");

                if (method == "GET" && (path == "/" || path == "/index.html" || path == "/home.htm"))
                {
                    await WriteHtmlAsync(res, HtmlPage);
                    return;
                }

                if (path.StartsWith("/api/records"))
                {
                    await HandleApiAsync(context, path, method);
                    return;
                }

                if (path == "/api/query" && method == "POST")
                {
                    await HandleQueryAsync(context);
                    return;
                }

                if (path == "/api/transazione" && method == "POST")
                {
                    await HandleTransactionAsync(context);
                    return;
                }

                if (path == "/api/database/elenco" && method == "GET")
                {
                    await HandleListDatabasesAsync(context);
                    return;
                }

                if (path == "/api/database/attuale" && method == "GET")
                {
                    await WriteJsonAsync(res, new { percorso = _db.DbPath });
                    return;
                }

                if (path == "/api/database/seleziona" && method == "POST")
                {
                    await HandleSelectDatabaseAsync(context);
                    return;
                }

                res.StatusCode = 404;
                await WriteJsonAsync(res, new { errore = "Non trovato" });
            }
            catch (Exception ex)
            {
                Log("ERRORE: " + ex.Message);
                try
                {
                    res.StatusCode = 500;
                    await WriteJsonAsync(res, new { errore = ex.Message });
                }
                catch { /* connessione già chiusa */ }
            }
            finally
            {
                res.OutputStream.Close();
            }
        }

        private static async Task HandleApiAsync(HttpListenerContext context, string path, string method)
        {
            var req = context.Request;
            var res = context.Response;

            // /api/records         -> lista (GET) / creazione (POST)
            // /api/records/{id}    -> singolo (GET) / modifica (PUT) / cancellazione (DELETE)
            string[] parts = path.Trim('/').Split('/');
            int? id = null;
            if (parts.Length == 3 && int.TryParse(parts[2], out int parsedId)) id = parsedId;

            switch (method)
            {
                case "GET" when id == null:
                    await WriteJsonAsync(res, _db.GetAll());
                    break;

                case "GET" when id != null:
                    var rec = _db.GetById(id.Value);
                    if (rec == null) { res.StatusCode = 404; await WriteJsonAsync(res, new { errore = "Record non trovato" }); }
                    else await WriteJsonAsync(res, rec);
                    break;

                case "POST":
                    var newRec = await ReadJsonAsync<Record>(req);
                    if (newRec == null) { res.StatusCode = 400; await WriteJsonAsync(res, new { errore = "Corpo non valido" }); break; }
                    int newId = _db.Insert(newRec);
                    newRec.Id = newId;
                    res.StatusCode = 201;
                    await WriteJsonAsync(res, newRec);
                    break;

                case "PUT" when id != null:
                    var updated = await ReadJsonAsync<Record>(req);
                    if (updated == null) { res.StatusCode = 400; await WriteJsonAsync(res, new { errore = "Corpo non valido" }); break; }
                    bool ok = _db.Update(id.Value, updated);
                    if (!ok) { res.StatusCode = 404; await WriteJsonAsync(res, new { errore = "Record non trovato" }); }
                    else { updated.Id = id.Value; await WriteJsonAsync(res, updated); }
                    break;

                case "DELETE" when id != null:
                    bool deleted = _db.Delete(id.Value);
                    if (!deleted) { res.StatusCode = 404; await WriteJsonAsync(res, new { errore = "Record non trovato" }); }
                    else await WriteJsonAsync(res, new { successo = true });
                    break;

                default:
                    res.StatusCode = 405;
                    await WriteJsonAsync(res, new { errore = "Metodo non supportato" });
                    break;
            }
        }

        private static async Task HandleQueryAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            var body = await ReadJsonAsync<QueryRequest>(req);
            if (body == null || string.IsNullOrWhiteSpace(body.Sql))
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Devi specificare il campo 'sql'." });
                return;
            }

            try
            {
                var risultato = _db.ExecuteRawQuery(body.Sql);
                await WriteJsonAsync(res, new { successo = true, risultato });
            }
            catch (Exception ex)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        private static async Task HandleTransactionAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            var body = await ReadJsonAsync<TransactionRequest>(req);
            if (body == null || body.Operazioni == null || body.Operazioni.Count == 0)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Devi specificare almeno un'operazione nel campo 'operazioni'." });
                return;
            }

            try
            {
                var risultati = _db.EseguiTransazione(body.Operazioni);
                await WriteJsonAsync(res, new { successo = true, operazioniEseguite = risultati.Count, risultati });
            }
            catch (Exception ex)
            {
                // Nessuna delle operazioni è stata mantenuta: rollback già avvenuto in EseguiTransazione.
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message, nota = "Nessuna modifica è stata applicata: la transazione è stata annullata interamente." });
            }
        }

        // Cerca nella cartella dell'exe tutti i file .accdb e .mdb, e indica quale
        // dei due è attualmente collegato al server.
        private static async Task HandleListDatabasesAsync(HttpListenerContext context)
        {
            var res = context.Response;

            var trovati = new List<object>();
            try
            {
                var file = Directory.GetFiles(_exeDir, "*.accdb")
                    .Concat(Directory.GetFiles(_exeDir, "*.mdb"))
                    .OrderBy(f => f);

                foreach (var percorso in file)
                {
                    var info = new FileInfo(percorso);
                    trovati.Add(new
                    {
                        nomeFile = info.Name,
                        percorso = info.FullName,
                        dimensioneByte = info.Length,
                        ultimaModifica = info.LastWriteTime,
                        selezionato = string.Equals(info.FullName, _db.DbPath, StringComparison.OrdinalIgnoreCase)
                    });
                }

                await WriteJsonAsync(res, new { successo = true, cartella = _exeDir, database = trovati });
            }
            catch (Exception ex)
            {
                res.StatusCode = 500;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        // Cambia il database attivo. Il percorso può essere assoluto oppure solo il nome
        // del file (in tal caso viene cercato nella cartella dell'exe).
        private static async Task HandleSelectDatabaseAsync(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            var body = await ReadJsonAsync<DatabaseSelectRequest>(req);
            if (body == null || string.IsNullOrWhiteSpace(body.Percorso))
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = "Devi specificare il campo 'percorso'." });
                return;
            }

            string percorsoCompleto = Path.IsPathRooted(body.Percorso)
                ? body.Percorso
                : Path.Combine(_exeDir, body.Percorso);

            try
            {
                lock (_dbLock)
                {
                    var nuovoDb = new AccessDb(percorsoCompleto);

                    if (!File.Exists(percorsoCompleto))
                    {
                        if (!body.Crea)
                            throw new FileNotFoundException($"Il file '{percorsoCompleto}' non esiste. Manda 'crea: true' se vuoi crearlo.");

                        nuovoDb.EnsureDatabase(Log); // crea file .accdb + tabella Records di base
                    }

                    _db = nuovoDb;
                }

                await WriteJsonAsync(res, new { successo = true, percorso = percorsoCompleto });
            }
            catch (Exception ex)
            {
                res.StatusCode = 400;
                await WriteJsonAsync(res, new { successo = false, errore = ex.Message });
            }
        }

        private static async Task<T?> ReadJsonAsync<T>(HttpListenerRequest req)
        {
            using var reader = new StreamReader(req.InputStream, req.ContentEncoding);
            string body = await reader.ReadToEndAsync();
            if (string.IsNullOrWhiteSpace(body)) return default;
            return JsonSerializer.Deserialize<T>(body, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }

        private static async Task WriteJsonAsync(HttpListenerResponse res, object data)
        {
            res.ContentType = "application/json; charset=utf-8";
            string json = JsonSerializer.Serialize(data);
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes);
        }

        private static async Task WriteHtmlAsync(HttpListenerResponse res, string html)
        {
            res.ContentType = "text/html; charset=utf-8";
            byte[] bytes = Encoding.UTF8.GetBytes(html);
            res.ContentLength64 = bytes.Length;
            await res.OutputStream.WriteAsync(bytes);
        }

        // ------------------------------------------------------------------
        // Pagina web incorporata: nessun file esterno necessario
        // ------------------------------------------------------------------
        private const string HtmlPage = @"<!DOCTYPE html>
<html lang=""it"">
<head>
<meta charset=""UTF-8"">
<title>Gestione Database</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; max-width: 800px; margin: 40px auto; background: #f4f5f7; color: #222; }
  h1 { color: #2c3e50; }
  form { background: white; padding: 16px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 24px; }
  input { padding: 8px; margin-right: 8px; border: 1px solid #ccc; border-radius: 4px; width: 200px; }
  button { padding: 8px 16px; border: none; border-radius: 4px; background: #2c7be5; color: white; cursor: pointer; }
  button:hover { background: #1a5fc0; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.1); }
  th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; }
  th { background: #2c3e50; color: white; }
  .azioni button { margin-right: 6px; font-size: 12px; padding: 4px 10px; }
  .btn-elimina { background: #e63757; }
  .btn-elimina:hover { background: #c72a45; }
  .btn-modifica { background: #f5a623; }
  .btn-modifica:hover { background: #d9910f; }
</style>
</head>
<body>
  <h1>Gestione Record (Database Access)</h1>

  <div style=""background:white;padding:16px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);margin-bottom:24px;"">
    <strong>Database attivo:</strong> <span id=""db-attuale"">(caricamento...)</span><br><br>
    <select id=""db-elenco"" style=""padding:8px;border:1px solid #ccc;border-radius:4px;min-width:280px;""></select>
    <button type=""button"" id=""btn-db-aggiorna"">Aggiorna elenco</button>
    <button type=""button"" id=""btn-db-seleziona"">Usa questo database</button>
  </div>

  <form id=""form-record"">
    <input type=""hidden"" id=""record-id"">
    <input type=""text"" id=""nome"" placeholder=""Nome"" required>
    <input type=""text"" id=""descrizione"" placeholder=""Descrizione"">
    <button type=""submit"" id=""btn-salva"">Aggiungi</button>
    <button type=""button"" id=""btn-annulla"" style=""display:none;background:#888;"">Annulla</button>
  </form>

  <table>
    <thead><tr><th>ID</th><th>Nome</th><th>Descrizione</th><th>Azioni</th></tr></thead>
    <tbody id=""tabella-corpo""></tbody>
  </table>

  <h2>Query SQL libera</h2>
  <form id=""form-query"">
    <textarea id=""sql-input"" rows=""4"" style=""width:100%;box-sizing:border-box;padding:8px;border:1px solid #ccc;border-radius:4px;font-family:Consolas,monospace;""
      placeholder=""Es: SELECT * FROM Records WHERE Nome LIKE '%mario%'  oppure  UPDATE Records SET Nome='X' WHERE ID=1""></textarea>
    <br><br>
    <button type=""submit"">Esegui query</button>
  </form>
  <div id=""risultato-query""></div>

<script>
const apiUrl = '/api/records';
const dbAttualeSpan = document.getElementById('db-attuale');
const dbElencoSelect = document.getElementById('db-elenco');
const btnDbAggiorna = document.getElementById('btn-db-aggiorna');
const btnDbSeleziona = document.getElementById('btn-db-seleziona');

async function aggiornaElencoDatabase() {
  const risposta = await fetch('/api/database/elenco');
  const dati = await risposta.json();
  dbElencoSelect.innerHTML = '';
  if (!dati.successo || dati.database.length === 0) {
    const opzione = document.createElement('option');
    opzione.textContent = 'Nessun file .accdb/.mdb trovato nella cartella';
    dbElencoSelect.appendChild(opzione);
    return;
  }
  dati.database.forEach(db => {
    const opzione = document.createElement('option');
    opzione.value = db.percorso;
    opzione.textContent = db.nomeFile + (db.selezionato ? '  (attivo)' : '');
    if (db.selezionato) opzione.selected = true;
    dbElencoSelect.appendChild(opzione);
  });
}

async function aggiornaDbAttuale() {
  const risposta = await fetch('/api/database/attuale');
  const dati = await risposta.json();
  dbAttualeSpan.textContent = dati.percorso;
}

btnDbAggiorna.addEventListener('click', async () => {
  await aggiornaElencoDatabase();
});

btnDbSeleziona.addEventListener('click', async () => {
  const percorso = dbElencoSelect.value;
  if (!percorso) return;
  const risposta = await fetch('/api/database/seleziona', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ percorso, crea: false })
  });
  const dati = await risposta.json();
  if (!dati.successo) {
    alert('Errore nel selezionare il database: ' + dati.errore);
    return;
  }
  await aggiornaDbAttuale();
  await aggiornaElencoDatabase();
  caricaRecord();
});

const form = document.getElementById('form-record');
const inputId = document.getElementById('record-id');
const inputNome = document.getElementById('nome');
const inputDescrizione = document.getElementById('descrizione');
const btnSalva = document.getElementById('btn-salva');
const btnAnnulla = document.getElementById('btn-annulla');
const corpo = document.getElementById('tabella-corpo');

async function caricaRecord() {
  const risposta = await fetch(apiUrl);
  const dati = await risposta.json();
  corpo.innerHTML = '';
  dati.forEach(r => {
    const riga = document.createElement('tr');
    riga.innerHTML = `
      <td>${r.id}</td>
      <td>${escapeHtml(r.nome)}</td>
      <td>${escapeHtml(r.descrizione)}</td>
      <td class=""azioni"">
        <button class=""btn-modifica"" onclick=""modificaRecord(${r.id}, '${escapeAttr(r.nome)}', '${escapeAttr(r.descrizione)}')"">Modifica</button>
        <button class=""btn-elimina"" onclick=""eliminaRecord(${r.id})"">Elimina</button>
      </td>`;
    corpo.appendChild(riga);
  });
}

function escapeHtml(s) {
  return (s ?? '').replace(/[&<>""']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','""':'&quot;',""'"":'&#39;'}[c]));
}
function escapeAttr(s) { return escapeHtml(s).replace(/'/g, ""\\'""); }

function modificaRecord(id, nome, descrizione) {
  inputId.value = id;
  inputNome.value = nome;
  inputDescrizione.value = descrizione;
  btnSalva.textContent = 'Salva modifiche';
  btnAnnulla.style.display = 'inline-block';
}

function annullaModifica() {
  inputId.value = '';
  form.reset();
  btnSalva.textContent = 'Aggiungi';
  btnAnnulla.style.display = 'none';
}
btnAnnulla.addEventListener('click', annullaModifica);

async function eliminaRecord(id) {
  if (!confirm('Confermi la cancellazione del record ' + id + '?')) return;
  await fetch(apiUrl + '/' + id, { method: 'DELETE' });
  caricaRecord();
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const dati = { nome: inputNome.value, descrizione: inputDescrizione.value };
  const id = inputId.value;

  if (id) {
    await fetch(apiUrl + '/' + id, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(dati)
    });
  } else {
    await fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(dati)
    });
  }
  annullaModifica();
  caricaRecord();
});

const formQuery = document.getElementById('form-query');
const sqlInput = document.getElementById('sql-input');
const risultatoDiv = document.getElementById('risultato-query');

formQuery.addEventListener('submit', async (e) => {
  e.preventDefault();
  const sql = sqlInput.value.trim();
  if (!sql) return;

  const risposta = await fetch('/api/query', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sql })
  });
  const dati = await risposta.json();
  mostraRisultatoQuery(dati);
  caricaRecord(); // se la query ha toccato la tabella Records, aggiorna la vista sopra
});

function mostraRisultatoQuery(dati) {
  if (!dati.successo) {
    risultatoDiv.innerHTML = '<p style=""color:#e63757;"">Errore: ' + escapeHtml(dati.errore) + '</p>';
    return;
  }
  const r = dati.risultato;
  if (Array.isArray(r)) {
    if (r.length === 0) { risultatoDiv.innerHTML = '<p>Query eseguita: nessuna riga restituita.</p>'; return; }
    const colonne = Object.keys(r[0]);
    let html = '<table><thead><tr>' + colonne.map(c => '<th>' + escapeHtml(c) + '</th>').join('') + '</tr></thead><tbody>';
    r.forEach(riga => {
      html += '<tr>' + colonne.map(c => '<td>' + escapeHtml(String(riga[c] ?? '')) + '</td>').join('') + '</tr>';
    });
    html += '</tbody></table>';
    risultatoDiv.innerHTML = html;
  } else {
    risultatoDiv.innerHTML = '<p>Righe interessate: ' + r.righeInteressate + '</p>';
  }
}

aggiornaDbAttuale();
aggiornaElencoDatabase();
caricaRecord();
</script>
</body>
</html>";
    }
}
