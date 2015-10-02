using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace PIOLEDB_CSharp_01
{
    class Program
    {
        static void Main(string[] args)
        {
            // Please change the value of $PIServerName to refer your PI Data Archive host            string sPIServer = "PIServer";

            // Creating the connection string
            string sConnection = "Provider=PIOLEDB;Data Source=" + sPIServer + ";Integrated Security=SSPI;"; 
            
            // Defining the query string
            string sQuery = "SELECT tag, time, value FROM picomp2 WHERE tag = 'CDT158' AND time > 't'"; 
            
            // Create OLEDBConnection using the predefined connection string
            OleDbConnection pioledbConnection = new OleDbConnection(sConnection); 
            
            // Open the connection
            pioledbConnection.Open();

            // Create the OLEDB Command object, refer the query string and the connection
            OleDbCommand pioledbCommand = new OleDbCommand(sQuery, pioledbConnection);

            // Create reader object and execute the command
            OleDbDataReader Reader = pioledbCommand.ExecuteReader();

            // Get the amount of columns 
            int iFieldCount = Reader.FieldCount;
            
            // Repeat reading while there is content
            while (Reader.Read()) 
            { 
                // Read row content column by column
                for (int iField = 0; iField < iFieldCount; iField++) 
                { 
                    // Header
                    string sField = Reader.GetName(iField); 
                    // Value
                    string sValue = Reader.GetValue(iField).ToString(); 
                    // Output to console
                    Console.WriteLine("{0} = {1}", sField, sValue); 
                } 
                
                // Insert blank line to console
                Console.WriteLine(); 
            }

            // Close the reader and the connection
            Reader.Close();
            pioledbConnection.Close();
            
            // Wait for user input before closing console
            Console.Write("Done. Press any key to quit .. ");
            Console.ReadKey();
        }
    }
}
