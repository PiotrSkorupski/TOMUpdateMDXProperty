
using Microsoft.AnalysisServices.Tabular;
using System;
using System.Security;

namespace TOMUpdateProperty
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.Gray;

            if (args.Length == 3)
            {
                Console.WriteLine(String.Format("Using server name: {0}", args[0]));
                Console.WriteLine(String.Format("Using database name: {0}", args[1]));
                Console.WriteLine(String.Format("Disable / Enable IsAvailableInMDX: {0}", args[2]));
                // 
                // Connect to the local default instance of Analysis Services 
                // 
                string ConnectionString = "Provider=MSOLAP;Data Source=" + args[0] + ";";
                string databaseName = args[1];
                int updatedFieldsCount = 0;
                int fieldCount = 0;
                string disableEnable = args[2];


                // 
                // The using syntax ensures the correct use of the 
                // Microsoft.AnalysisServices.Tabular.Server object. 
                // 
                try
                {
                    var server = new Server();
                    server.Connect(ConnectionString);
                    if (server.Connected)
                    {
                        // Connect to the server
                        Database db = server.Databases[databaseName];

                        // Connect to the database
                        Model model = db.Model;
                        
                        Console.WriteLine("Connection established successfully.");
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Server name:\t\t{0}", server.Name);
                        Console.WriteLine("Server product name:\t{0}", server.ProductName);
                        Console.WriteLine("Server product level:\t{0}", server.ProductLevel);
                        Console.WriteLine("Server version:\t\t{0}", server.Version);
                        
                        Console.ResetColor();
                        Console.WriteLine();

                        var tables = model.Tables;

                        foreach (Table t in tables)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine(String.Format("Processing table: {0}", t.Name));
                            Console.ResetColor();

                            var columns = t.Columns;
                            foreach (Column c in columns)
                            {
                                fieldCount++;
                                Console.ForegroundColor = ConsoleColor.Gray;
                                Console.Write("Veryfying column: ");
                                Console.ForegroundColor = ConsoleColor.DarkMagenta;
                                Console.WriteLine(c.Name);
                                Console.ResetColor();
                                Console.Write(String.Format("IsAvailableInMDX value: {0}. ", c.IsAvailableInMDX.ToString()));

                                if (c.IsKey == false)
                                {
                                    if (disableEnable == "D")
                                    {
                                        if (c.IsAvailableInMDX == true)
                                        {
                                            Console.WriteLine(c.IsKey.ToString());
                                            Console.ForegroundColor = ConsoleColor.Magenta;
                                            Console.WriteLine(String.Format("IsAvailableInMDX value is TRUE. Setting to FALSE."));
                                            c.IsAvailableInMDX = false;

                                            updatedFieldsCount++;
                                        }
                                        else
                                        {
                                            Console.WriteLine(String.Format("IsAvailableInMDX value is FALSE. Skipping."));
                                        }
                                    }
                                    else if (disableEnable == "E")
                                    {
                                        if (c.IsAvailableInMDX == false)
                                        {
                                            Console.WriteLine(c.IsKey.ToString());
                                            Console.ForegroundColor = ConsoleColor.Magenta;
                                            Console.WriteLine(String.Format("IsAvailableInMDX value is FALSE. Setting to TRUE."));
                                            c.IsAvailableInMDX = true;

                                            updatedFieldsCount++;
                                        }
                                        else
                                        {
                                            Console.WriteLine(String.Format("IsAvailableInMDX value is TRUE. Skipping."));
                                        }
                                    }
                                } else
                                {
                                    Console.WriteLine(String.Format("Column {0} is system column. Skipping.", c.Name));
                                }

                            }
                        }

                        db.Update(Microsoft.AnalysisServices.UpdateOptions.ExpandFull);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(String.Format("Updated {0} columns of {1} total.", updatedFieldsCount, fieldCount));
                        Console.ResetColor();

                        Console.WriteLine("Press Enter to close this console window.");
                        Console.ReadLine();
                    }
                    else
                    { Console.WriteLine("AAS Connection failure"); }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("exception type of " + ex.GetType().FullName +  ": " + ex.Message + ex.InnerException);
                    Console.ResetColor();
                    Console.ReadLine();
                }
            } else
            {
                Console.WriteLine("Please provide server name,database name and disable / enable parameters. D for disable, E for enable");
            }
        }
    }
}