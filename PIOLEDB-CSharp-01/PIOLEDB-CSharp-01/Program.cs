/// <license>
/// Copyright 2015 OSIsoft, LLC
/// Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in
/// compliance with the License. You may obtain a copy of the License at
///
///      http://www.apache.org/licenses/LICENSE-2.0
///
/// Unless required by applicable law or agreed to in writing, software distributed under the License is
/// distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
/// See the License for the specific language governing permissions and limitations under the License.
/// </license>
/// <author>
/// Gregor Beck, OSIsoft Europe GmbH
/// </author>
/// <contact>
/// pidevclub@osisoft.com
/// </contact>

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
            // Please change the value of $PIServerName to refer your PI Data Archive host
            string sPIServer = "PIServer";

            // Creating the connection string
            string sConnection = "Provider=PIOLEDB;Data Source=" + sPIServer + ";Integrated Security=SSPI;";

            // Defining the query string
            string sQuery = "SELECT tag, time, value FROM picomp2 WHERE tag = 'CDT158' AND time > 't'";

            // Create OLEDBConnection using the predefined connection string
            using (OleDbConnection pioledbConnection = new OleDbConnection(sConnection))
            {
                // Create the OLEDB Command object, refer the query string and the connection
                OleDbCommand pioledbCommand = new OleDbCommand(sQuery, pioledbConnection);

                // Open the connection
                pioledbConnection.Open();

                // Create reader object and execute the command
                OleDbDataReader Reader = pioledbCommand.ExecuteReader();

                // Get the amount of columns 
                int iFieldCount = Reader.FieldCount;

                // Repeat reading while there is content

                bool headersPrinted = false;
                while (Reader.Read())
                {

                    // column headers
                    if (!headersPrinted)
                    {
                        for (int iField = 0; iField < iFieldCount; iField++)
                        {
                            // Header
                            string sField = Reader.GetName(iField);
                            Console.Write("|{0,5}|", sField);
                        }
                        Console.WriteLine();
                        headersPrinted = true;
                    }


                    for (int iField = 0; iField < iFieldCount; iField++)
                    {
                        // Value
                        string sValue = Reader.GetValue(iField).ToString();
                        // Output to console
                        Console.Write("|{0,5}|", sValue);
                    }

                    // Insert blank line to console
                    Console.WriteLine();
                }

                // Close the reader and the connection
                Reader.Close();
            }

            // Wait for user input before closing console
            Console.Write("Done. Press any key to quit .. ");
            Console.ReadKey();
        }
        }
    }
}
