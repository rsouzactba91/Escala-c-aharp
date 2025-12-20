using System;
using System.Data;
using System.IO;
using Microsoft.Data.Sqlite;
using System.Collections.Generic;

namespace Escala
{
    public static class DatabaseService
    {
        private static string DbPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "escala.db");
        private static string ConnectionString => $"Data Source={DbPath}";

        public static void Initialize()
        {
            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();

                // 1. Tabela Principal de Dados Mensais (Estilo Planilha)
                var cmd = connection.CreateCommand();
                cmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS MonthlyData (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        RowIndex INTEGER,
                        C1 TEXT, C2 TEXT, C3 TEXT, C4 TEXT, C5 TEXT, C6 TEXT, C7 TEXT, C8 TEXT, C9 TEXT, C10 TEXT,
                        C11 TEXT, C12 TEXT, C13 TEXT, C14 TEXT, C15 TEXT, C16 TEXT, C17 TEXT, C18 TEXT, C19 TEXT, C20 TEXT,
                        C21 TEXT, C22 TEXT, C23 TEXT, C24 TEXT, C25 TEXT, C26 TEXT, C27 TEXT, C28 TEXT, C29 TEXT, C30 TEXT,
                        C31 TEXT, C32 TEXT, C33 TEXT, C34 TEXT, C35 TEXT, C36 TEXT
                    );

                    CREATE TABLE IF NOT EXISTS DailyAssignments (
                        DateKey TEXT NOT NULL,
                        StaffName TEXT NOT NULL,
                        TimeSlot TEXT NOT NULL,
                        Post TEXT,
                        PRIMARY KEY (DateKey, StaffName, TimeSlot)
                    );
                ";
                cmd.ExecuteNonQuery();

                // 2. Tabelas de Configuração (Postos, Horários e Configs Gerais)
                var cmdConfig = connection.CreateCommand();
                cmdConfig.CommandText = @"
                    CREATE TABLE IF NOT EXISTS ConfigPostos (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        Nome TEXT UNIQUE
                    );

                    CREATE TABLE IF NOT EXISTS ConfigHorarios (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        Descricao TEXT UNIQUE, 
                        Ordem INTEGER
                    );

                    -- NOVA TABELA: Substitui os arquivos .txt
                    CREATE TABLE IF NOT EXISTS AppSettings (
                        Chave TEXT PRIMARY KEY,
                        Valor TEXT
                    );
                ";
                cmdConfig.ExecuteNonQuery();

                InicializarDadosPadrao(connection);
            }
        }

        private static void InicializarDadosPadrao(SqliteConnection conn)
        {
            // Verifica e inicia Postos se estiver vazio
            var cmdCheckPostos = new SqliteCommand("SELECT COUNT(*) FROM ConfigPostos", conn);
            long countPostos = (long)cmdCheckPostos.ExecuteScalar();

            if (countPostos == 0)
            {
                var postosPadrao = new[] { "", "CAIXA", "VALET", "QRF", "CIRC.", "REP|CIRC", "CS1", "CS2", "CS3", "SUP", "APOIO", "TREIN", "CFTV", "ECHO 21" };
                foreach (var p in postosPadrao)
                {
                    var cmd = new SqliteCommand("INSERT INTO ConfigPostos (Nome) VALUES (@nome)", conn);
                    cmd.Parameters.AddWithValue("@nome", p);
                    cmd.ExecuteNonQuery();
                }
            }

            // Verifica e inicia Horários se estiver vazio
            var cmdCheckHorarios = new SqliteCommand("SELECT COUNT(*) FROM ConfigHorarios", conn);
            long countHorarios = (long)cmdCheckHorarios.ExecuteScalar();

            if (countHorarios == 0)
            {
                var horariosPadrao = new[] {
                    "08:00 x 08:40", "08:41 x 09:40", "09:41 x 10:40", "10:41 x 11:40",
                    "11:41 x 12:40", "12:41 x 13:40", "13:41 x 14:40", "14:41 x 15:40",
                    "15:41 x 16:40", "16:41 x 17:40", "17:41 x 18:40", "18:41 x 19:40",
                    "19:41 x 20:40", "20:41 x 21:40", "21:41 x 22:40", "22:41 x 23:40",
                    "23:41 x 00:40", "00:41 x 01:40"
                };

                int ordem = 1;
                foreach (var h in horariosPadrao)
                {
                    var cmd = new SqliteCommand("INSERT INTO ConfigHorarios (Descricao, Ordem) VALUES (@desc, @ordem)", conn);
                    cmd.Parameters.AddWithValue("@desc", h);
                    cmd.Parameters.AddWithValue("@ordem", ordem++);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // --- MÉTODOS GENÉRICOS PARA CONFIGURAÇÕES (SUBSTITUI OS ARQUIVOS TXT) ---

        public static void SetSetting(string chave, string valor)
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                // INSERT OR REPLACE: Se já existe a chave, atualiza. Se não, cria.
                var cmd = new SqliteCommand("INSERT OR REPLACE INTO AppSettings (Chave, Valor) VALUES (@k, @v)", conn);
                cmd.Parameters.AddWithValue("@k", chave);
                cmd.Parameters.AddWithValue("@v", valor);
                cmd.ExecuteNonQuery();
            }
        }

        public static string GetSetting(string chave, string valorPadrao)
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmd = new SqliteCommand("SELECT Valor FROM AppSettings WHERE Chave = @k", conn);
                cmd.Parameters.AddWithValue("@k", chave);
                var result = cmd.ExecuteScalar();
                return result != null ? result.ToString() : valorPadrao;
            }
        }

        // --- MÉTODOS ESPECÍFICOS (Agora usam o banco internamente) ---

        public static void SetHorarioPadraoFolguista(string horario)
        {
            SetSetting("HorarioFolguista", horario);
        }

        public static string GetHorarioPadraoFolguista()
        {
            return GetSetting("HorarioFolguista", "12:40 X 21:00");
        }

        public static void SetHorarioPadraoIntermediario(string horario)
        {
            SetSetting("HorarioIntermediario", horario);
        }

        public static string GetHorarioPadraoIntermediario()
        {
            return GetSetting("HorarioIntermediario", "12:40 X 21:00");
        }


        // --- DEMAIS MÉTODOS (MANTERAM-SE IGUAIS MAS ORGANIZADOS) ---

        public static List<string> GetPostosConfigurados()
        {
            var lista = new List<string>();
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmd = new SqliteCommand("SELECT Nome FROM ConfigPostos ORDER BY Nome", conn);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read()) lista.Add(reader.GetString(0));
                }
            }
            if (!lista.Contains("")) lista.Insert(0, "");
            return lista;
        }

        public static List<string> GetHorariosConfigurados()
        {
            var lista = new List<string>();
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmd = new SqliteCommand("SELECT Descricao FROM ConfigHorarios ORDER BY Ordem", conn);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read()) lista.Add(reader.GetString(0));
                }
            }
            return lista;
        }

        public static void AdicionarPosto(string nome)
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmd = new SqliteCommand("INSERT INTO ConfigPostos (Nome) VALUES (@nome)", conn);
                cmd.Parameters.AddWithValue("@nome", nome.ToUpper());
                cmd.ExecuteNonQuery();
            }
        }

        public static void RemoverPosto(string nome)
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmd = new SqliteCommand("DELETE FROM ConfigPostos WHERE Nome = @nome", conn);
                cmd.Parameters.AddWithValue("@nome", nome);
                cmd.ExecuteNonQuery();
            }
        }

        public static void AdicionarHorario(string descricao)
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmdOrd = new SqliteCommand("SELECT IFNULL(MAX(Ordem), 0) + 1 FROM ConfigHorarios", conn);
                int novaOrdem = Convert.ToInt32(cmdOrd.ExecuteScalar());

                var cmd = new SqliteCommand("INSERT INTO ConfigHorarios (Descricao, Ordem) VALUES (@desc, @ordem)", conn);
                cmd.Parameters.AddWithValue("@desc", descricao);
                cmd.Parameters.AddWithValue("@ordem", novaOrdem);
                cmd.ExecuteNonQuery();
            }
        }

        public static void RemoverHorario(string descricao)
        {
            using (var conn = new SqliteConnection(ConnectionString))
            {
                conn.Open();
                var cmd = new SqliteCommand("DELETE FROM ConfigHorarios WHERE Descricao = @desc", conn);
                cmd.Parameters.AddWithValue("@desc", descricao);
                cmd.ExecuteNonQuery();
            }
        }

        public static void SaveMonthlyData(DataTable dt)
        {
            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    var cmdClear = connection.CreateCommand();
                    cmdClear.Transaction = transaction;
                    cmdClear.CommandText = "DELETE FROM MonthlyData";
                    cmdClear.ExecuteNonQuery();

                    foreach (DataRow row in dt.Rows)
                    {
                        var cmdInsert = connection.CreateCommand();
                        cmdInsert.Transaction = transaction;

                        var cols = new List<string>();
                        var vals = new List<string>();

                        for (int i = 1; i <= 36; i++)
                        {
                            cols.Add($"C{i}");
                            vals.Add($"@C{i}");
                            cmdInsert.Parameters.AddWithValue($"@C{i}", row[i - 1] ?? DBNull.Value);
                        }

                        cmdInsert.CommandText = $"INSERT INTO MonthlyData ({string.Join(",", cols)}) VALUES ({string.Join(",", vals)})";
                        cmdInsert.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
            }
        }

        public static DataTable GetMonthlyData()
        {
            var dt = new DataTable();
            for (int i = 1; i <= 36; i++) dt.Columns.Add($"C{i}");

            if (!File.Exists(DbPath)) return dt;

            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();
                var cmd = connection.CreateCommand();
                cmd.CommandText = "SELECT * FROM MonthlyData ORDER BY Id";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var row = dt.NewRow();
                        for (int i = 1; i <= 36; i++)
                        {
                            row[i - 1] = reader[$"C{i}"];
                        }
                        dt.Rows.Add(row);
                    }
                }
            }
            return dt;
        }

        public static void SaveAssignment(int dia, string nome, string horario, string posto)
        {
            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();
                var cmd = connection.CreateCommand();
                cmd.CommandText = @"
                    INSERT OR REPLACE INTO DailyAssignments (DateKey, StaffName, TimeSlot, Post)
                    VALUES (@DateKey, @StaffName, @TimeSlot, @Post)
                ";

                cmd.Parameters.AddWithValue("@DateKey", $"Dia {dia}");
                cmd.Parameters.AddWithValue("@StaffName", nome);
                cmd.Parameters.AddWithValue("@TimeSlot", horario);
                cmd.Parameters.AddWithValue("@Post", posto ?? "");

                cmd.ExecuteNonQuery();
            }
        }

        public static Dictionary<string, Dictionary<string, string>> GetAssignmentsForDay(int dia)
        {
            var result = new Dictionary<string, Dictionary<string, string>>();

            if (!File.Exists(DbPath)) return result;

            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();
                var cmd = connection.CreateCommand();
                cmd.CommandText = "SELECT StaffName, TimeSlot, Post FROM DailyAssignments WHERE DateKey = @DateKey";
                cmd.Parameters.AddWithValue("@DateKey", $"Dia {dia}");

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string nome = reader.GetString(0);
                        string horario = reader.GetString(1);
                        string posto = reader.GetString(2);

                        if (!result.ContainsKey(nome))
                        {
                            result[nome] = new Dictionary<string, string>();
                        }
                        result[nome][horario] = posto;
                    }
                }
            }
            return result;
        }

        public static void ClearAllAssignments()
        {
            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();
                var cmd = connection.CreateCommand();
                cmd.CommandText = "DELETE FROM DailyAssignments";
                cmd.ExecuteNonQuery();
            }
        }

        public static void ClearAssignmentsForDay(int dia)
        {
            using (var connection = new SqliteConnection(ConnectionString))
            {
                connection.Open();
                var cmd = connection.CreateCommand();
                cmd.CommandText = "DELETE FROM DailyAssignments WHERE DateKey = @DateKey";
                cmd.Parameters.AddWithValue("@DateKey", $"Dia {dia}");
                cmd.ExecuteNonQuery();
            }
        }
    }
}