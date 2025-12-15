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

                // 1. Tabela Mensal (Armazena o Excel importado)
                // Vamos armazenar como JSON ou colunas dinâmicas? 
                // Para simplificar e manter a performance, vamos armazenar linha a linha, coluna a coluna é muito complexo dado o DataTable dinâmico.
                // Melhor: Armazenar células C1..C36
                
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
                        DateKey TEXT NOT NULL, -- Ex: 2023-10-05 (Vamos usar Dia apenas? Melhor usar chave 'Dia 5')
                        StaffName TEXT NOT NULL,
                        TimeSlot TEXT NOT NULL,
                        Post TEXT,
                        PRIMARY KEY (DateKey, StaffName, TimeSlot)
                    );
                ";
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
                    // Limpa tabela anterior
                    var cmdClear = connection.CreateCommand();
                    cmdClear.Transaction = transaction;
                    cmdClear.CommandText = "DELETE FROM MonthlyData";
                    cmdClear.ExecuteNonQuery();

                    // Insere novos dados
                    foreach (DataRow row in dt.Rows)
                    {
                        var cmdInsert = connection.CreateCommand();
                        cmdInsert.Transaction = transaction;
                        
                        // Monta a query dinamicamente é feio mas funcional para 36 colunas fixas
                        // Melhor usar parametros
                        var cols = new List<string>();
                        var vals = new List<string>();
                        
                        for (int i = 1; i <= 36; i++)
                        {
                            cols.Add($"C{i}");
                            vals.Add($"@C{i}");
                            cmdInsert.Parameters.AddWithValue($"@C{i}", row[i-1] ?? DBNull.Value);
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
                            row[i-1] = reader[$"C{i}"];
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
                
                // UPSERT (SQLite supporta INSERT OR REPLACE)
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
            // StaffName -> { TimeSlot -> Post }

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
    }
}
