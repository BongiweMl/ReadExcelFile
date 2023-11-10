using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ClientExcel
{
    
    
    class ReadExcelFile : IDisposable // use to release excel objects properly after reading data,
                                      // ReadExcel file was giving an error within a 'using' statement.
    {
        private string path;
        private _Application excel;
        private Workbook wb;
        private Worksheet ws;

       /// <summary>
       /// reads data from excel file
       /// </summary>
       /// <param name="path"></param>
       /// <param name="sheet"></param>
        public ReadExcelFile(string path, int sheet)
        {
            this.path = path;
            excel = new _Excel.Application();
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            try
            {
                i++;
                j++;

                if (ws.Cells[i, j].Value2 != null)
                {
                    return ws.Cells[i, j].Value2.ToString();
                }

                return "";
            }
            catch (Exception ex)
            {
                // exception for debugging 
                Console.WriteLine($"Error reading cell: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// string cli:interpolates the strings together
        /// </summary>
        /// <param name="clientList"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        public string DisplayList(List<Client> clientList, int position)
        {
            Client client = clientList.ElementAt(position);
            string cli = $"{client.FirstName} {client.Surname} {client.IdNumber} {client.PolicyNumber}";
            return cli;
        }
        public List<Client> ReadDataAndCreateClientList()
        {
            List<Client> clientList = new List<Client>();
            int row = 0, column = 0;

            try
            {
                while (!string.IsNullOrEmpty(ReadCell(row, column)))
                {
                    var Name = ReadCell(row, 0);
                    var Surname = ReadCell(row, column + 1);
                    var IdNumber = ReadCell(row, column + 2);
                    var PolicyNumber = ReadCell(row, column + 3);

                    clientList.Add(new Client(Name, Surname, IdNumber, PolicyNumber));
                    row++;
                }
            }
            catch (Exception ex)
            {
                // Handles exceptions
                Console.WriteLine($"Error reading data from Excel: {ex.Message}");
            }

            return clientList; // prints client information
        }
        public void Dispose()
        {
            // Release Excel objects to avoid memory leaks
            wb.Close(false);
            excel.Quit();

          
        }
    }
}


        
