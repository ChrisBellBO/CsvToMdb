using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using ADODB;
using ADOX;
using JRO;
using LumenWorks.Framework.IO.Csv;
using System.Globalization;
using System.Linq;

namespace CsvToMdb
{
  static class Program
  {
    private static string SqlEscape(string sql)
    {
      return sql?.Replace("'", "''");
    }

    readonly struct ColumnInfo
    {
      public readonly ADOX.DataTypeEnum ColumnType;
      public readonly int ColumnSize;
      public readonly int MinValue;
      public readonly int MaxValue;

      public ColumnInfo(ADOX.DataTypeEnum columnType, int columnSize)
      {
        ColumnType = columnType;
        ColumnSize = columnSize;
        MinValue = 0;
        MaxValue = 0;
      }

      public ColumnInfo(ADOX.DataTypeEnum columnType, int columnSize, int minValue, int maxValue)
      {
        ColumnType = columnType;
        ColumnSize = columnSize;
        MinValue = minValue;
        MaxValue = maxValue;
      }
    }

    private static CsvReader GetCsvReader(StreamReader reader)
    {
      return new CsvReader(reader, true, Delimiter);
    }

    private static Dictionary<string, ColumnInfo> ChooseColumnTypes(string csvFile)
    {
      var columnTypes = new Dictionary<string, ColumnInfo>();

      // first populate in order of CSV columns
      using (var reader = new StreamReader(csvFile))
      {
        // populate
        using (var csvReader = GetCsvReader(reader))
        {
          csvReader.ReadNextRecord();
          var column = 0;
          var headers = csvReader.GetFieldHeaders();
          while (column < csvReader.FieldCount)
          {
            var field = headers[column];
            if (!Ignore.Contains(field))
              columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adInteger, 0);
            column++;
          }
        }
      }

      using (var reader = new StreamReader(csvFile))
      {
        // populate
        using (var csvReader = GetCsvReader(reader))
        {
          while (csvReader.ReadNextRecord())
          {
            foreach (string field in csvReader.GetFieldHeaders())
            {
              if (!Ignore.Contains(field) && !string.IsNullOrEmpty(csvReader[field]))
              {
                if (!int.TryParse(csvReader[field], out var number))
                {
                  if (double.TryParse(csvReader[field], out _))
                  {
                    if (columnTypes[field].ColumnType == ADOX.DataTypeEnum.adInteger)
                      columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adSingle, 0);
                  }
                  else
                  {
                    if (DateTime.TryParse(csvReader[field], out _))
                    {
                      if (columnTypes[field].ColumnType == ADOX.DataTypeEnum.adInteger)
                        columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adDate, 0);
                    }
                    else
                    {
                      // not working
                      if (csvReader[field] == "Yes" || csvReader[field] == "No")
                      {
                        if (columnTypes[field].ColumnType == ADOX.DataTypeEnum.adInteger)
                          columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adBoolean, 1);
                      }
                      else
                      {
                        var maxLength = Math.Max(columnTypes[field].ColumnSize, csvReader[field].Length);
                        if (maxLength > 255)
                          columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adLongVarWChar, maxLength);
                        else
                          columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adVarWChar, maxLength);
                      }
                    }
                  }
                }
                else
                {
                  // it's a number, check min and max
                  if (columnTypes[field].ColumnType == ADOX.DataTypeEnum.adInteger)
                  {
                    columnTypes[field] = new ColumnInfo(ADOX.DataTypeEnum.adInteger, 0,
                      Math.Min(columnTypes[field].MinValue, number),
                      Math.Max(columnTypes[field].MaxValue, number));
                  }
                }
              }
            }
          }
        }
      }

      return columnTypes;
    }

    static string GetConnectionString(string filename)
    {
      return string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5", filename);
    }

    static string AccessFieldName(string fieldName)
    {
      fieldName = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(fieldName);
      return fieldName.Replace(" ", "").Replace("/", "").Replace("?", "").Replace("-", "");
    }

    private static char Delimiter;
    private static string[] Ignore;

    static void Main(string[] args)
    {
      try
      {
        if (args.Length < 1)
        {
          // show usage
          Console.WriteLine("CSV to MDB application");
          Console.WriteLine("Usage - CsvToMdb.exe csvFile [delimiter] [primary key]");
          Console.WriteLine("[delimiter] : The character that is used as the separator between columns");
          Console.WriteLine("[primary key] : The column that will be set as the primary key");
          Console.WriteLine("[ignore columns] : Column that will not be added to the MDB (comma delimited)");
          Console.WriteLine("Press any key");
          Console.ReadKey();
        }
        else
        {
          // read command line
          var csvFile = args[0];

          Delimiter = CsvReader.DefaultDelimiter;
          if (args.Length > 1)
            Delimiter = args[1][0];

          Ignore = new string[] {};
          if (args.Length > 3)
            Ignore = args[3].Split(',');
          
          var fileNameWithPath = Path.ChangeExtension(csvFile, "mdb");

          if (File.Exists(fileNameWithPath))
            File.Delete(fileNameWithPath);

          // create the database
          Console.WriteLine("Creating the database");
          var cat = new Catalog();
          var connectionString = GetConnectionString(fileNameWithPath);
          cat.Create(connectionString);

          Console.WriteLine("Choosing column types");
          var columnTypes = ChooseColumnTypes(csvFile);

          // create the table
          Console.WriteLine("Creating the table");
          var tableName = Path.GetFileNameWithoutExtension(csvFile);

          var table = new Table {Name = tableName};

          foreach (var field in columnTypes.Keys)
          {
            var colType = columnTypes[field].ColumnType;
            if (colType == ADOX.DataTypeEnum.adInteger)
            {
              var min = columnTypes[field].MinValue;
              var max = columnTypes[field].MaxValue;
              if (min >= 0 && max <= 255)
                colType = ADOX.DataTypeEnum.adUnsignedTinyInt;
              else if (min >= -32768 && max <= 32767)
                colType = ADOX.DataTypeEnum.adSmallInt;
            }

            table.Columns.Append(AccessFieldName(field), colType, columnTypes[field].ColumnSize);
          }

          // make every column nullable
          foreach (ADOX.Column column in table.Columns)
          {
            if (column.Type != ADOX.DataTypeEnum.adBoolean)
              column.Attributes = ColumnAttributesEnum.adColNullable;
          }

          // primary key
          if (args.Length > 2)
          {
            table.Keys.Append("primaryKey", KeyTypeEnum.adKeyPrimary, args[2]);
          }

          cat.Tables.Append(table);
          var con = cat.ActiveConnection as Connection;
          con?.Close();
          Marshal.ReleaseComObject(cat);

          Console.WriteLine("Populating the table");
          var cursor = Console.CursorTop;
          var processed = 1;
          using (var reader = new StreamReader(csvFile))
          {
            // populate
            using (var csvReader = GetCsvReader(reader))
            {
              var conn = new OleDbConnection(connectionString);
              conn.Open();
    
              while (csvReader.ReadNextRecord())
              {
                using (var command = conn.CreateCommand())
                {
                  var commandText = "INSERT INTO [" + tableName + "] (";

                  // add field names
                  var count = 0;
                  foreach (string field in csvReader.GetFieldHeaders())
                  {
                    if (columnTypes.ContainsKey(field))
                    {
                      if (count > 0)
                        commandText += ", ";
                      commandText += "[" + AccessFieldName(field) + "]";

                      count++;
                    }
                  }
                  commandText += ") SELECT ";

                  // add 
                  count = 0;
                  foreach (string field in csvReader.GetFieldHeaders())
                  {
                    if (columnTypes.ContainsKey(field))
                    {
                      if (count > 0)
                        commandText += ", ";

                      var fieldVal = csvReader[field];
                      if (columnTypes[field].ColumnType == ADOX.DataTypeEnum.adBoolean)
                      {
                        switch (fieldVal)
                        {
                          case "Yes":
                            commandText += "-1";
                            break;
                          case "No":
                            commandText += "0";
                            break;
                          default:
                            throw new Exception("Unexpected boolean value - " + fieldVal);
                        }
                      }
                      else if (string.IsNullOrEmpty(fieldVal))
                        commandText += "NULL";
                      else if (columnTypes[field].ColumnType == ADOX.DataTypeEnum.adDate)
                      {
                        // format so SQL likes it
                        var dateTime = DateTime.Parse(fieldVal);
                        commandText += "'" + dateTime.ToString("dd/MM/yyyy") + "'";
                      }
                      else
                        commandText += "'" + SqlEscape(fieldVal) + "'";

                      count++;
                    }
                  }
                  command.CommandText = commandText;

                  Console.WriteLine("Processing record " + processed);
                  Console.SetCursorPosition(0, cursor);
                  command.ExecuteNonQuery();
                }

                processed++;

                // compress the database since it can get large very quickly
                if (processed % 100000 == 0)
                {
                  conn.Close();
                  conn.Dispose();

                  var jro = new JetEngine();
                  var backupFileName = Path.Combine(Path.GetDirectoryName(fileNameWithPath), "backup.mdb");
                  var backupFile = GetConnectionString(backupFileName);
                  jro.CompactDatabase(connectionString, backupFile);
                  File.Delete(fileNameWithPath);
                  File.Move(backupFileName, fileNameWithPath);

                  conn = new OleDbConnection(connectionString);
                  conn.Open();
                }
              }
            }
          }
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("An error occured - " + ex.Message);
        Console.WriteLine("Press any key");
        Console.ReadKey();
      }
    }
  }
}
