using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ClientExcel
{
   /// <summary>
   /// initializes variables and reads data from excel
   /// 
   /// </summary>
    class Program
    {

        static void Main(string[] args)
        {
            
            string path = "C:/Users/BongiweML/Desktop/tesing2.xlsx";
            int sheetIndex = 1; // Gets data from first sheet

            try
            {
                using (ReadExcelFile rd = new ReadExcelFile(path, sheetIndex))
                {
                    List<Client> clientList = rd.ReadDataAndCreateClientList(); //used to obtain the client list

                    //separator line
                    Console.WriteLine(new string('-', 64));

                    foreach (var client in clientList)
                    {
                        //prints table rows, left-aligned in 15 character width column
                        Console.WriteLine("|{0,-15} |{1,-15}|{2,-15}|{3,-15}|", client.FirstName, client.Surname, client.IdNumber, client.PolicyNumber);
                    
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            Console.ReadKey();
        }
    }
}
