using Microsoft.AnalysisServices.Tabular;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TOM_Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateAndRefreshModel();
            //GenerateDatabaseSchema("c:\\temp\\file.json");
            Console.ReadLine();
        }

        public static void GenerateDatabaseSchema(string filename)
        {
            //Ignore some properties that are not needed for editing.
            SerializeOptions SSDTSerializeOptions = new SerializeOptions()
            {
                IgnoreInferredObjects = true,
                IgnoreInferredProperties = true,
                IgnoreTimestamps = true,
                SplitMultilineStrings = true
            };

            string schemafile = JsonSerializer.GenerateSchema(typeof (Database), SSDTSerializeOptions);
            try
            {
                using (StreamWriter file = File.CreateText(filename))
                {
                    file.Write(schemafile);
                    Console.WriteLine("File written to disk.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Could not save the file to disk. Original error: " + ex.Message);
            }
        }

        public static void CreateAndRefreshModel()
        {
            string serverConnectionString = @"Provider=MSOLAP;Data Source=.";
            var server = new Server();
            server.Connect(serverConnectionString);                           //Connect to the server.

            Database Db = new Database("TOM demo");                           //Create a new DB
            server.Databases.Add(Db);                                           //Add it to the server
            Model m = new Model();                                            //Create a new model
            Db.Model = m;                                                     //Add it to the database

            //Add a datasource connection
            ProviderDataSource foodmartDataSource = new ProviderDataSource()
            {
                Name = "Demo sales",
                ConnectionString = "Provider=SERVERNAME;Data Source=.;Initial Catalog=CATALOG;Integrated Security=SSPI;Persist Security Info=false",
                ImpersonationMode = ImpersonationMode.ImpersonateServiceAccount,
            };
            m.DataSources.Add(foodmartDataSource);

            Table dateTable = new Table();                                    //Create a new datetable 
            dateTable.Name = "DimDate";
            Partition p1 = new Partition();                                   //Create a new partition
            dateTable.Partitions.Add(p1);                                     //Add the partition to the datetable
            p1.Source = new QueryPartitionSource()                            //Add the partition source
            {
                DataSource = foodmartDataSource,
                Query = @"SELECT [Datekey]
                                ,[FullDateLabel]
                                FROM [dbo].[DimDate]",
            };
            m.Tables.Add(dateTable);                                          //Add the datatable to the model
            dateTable.Columns.Add(new DataColumn()                            //Add the columns
            {
                Name = "Datekey",
                DataType = DataType.DateTime,
                SourceColumn = "Datekey",
                FormatString = "General Date",
            });
            dateTable.Columns.Add(new DataColumn()
            {
                Name = "FullDateLabel",
                DataType = DataType.String,
                SourceColumn = "FullDateLabel",
            });

            Db.Update(Microsoft.AnalysisServices.UpdateOptions.ExpandFull);     // Update database with ExpandFull to commit Database together with Model
            Console.WriteLine("Database created");
            m.RequestRefresh(RefreshType.Full);                                  //Refresh the data
            m.SaveChanges();                                                    //Commit the changes
            Console.WriteLine("Database refreshed");
        }

    }

}
