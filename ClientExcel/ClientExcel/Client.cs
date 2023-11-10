using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClientExcel
{/// <summary>
/// gets and sets the strings and represents client data
/// </summary>
    class Client
    {
        private string firstName;
        private string surname;
        private string idNumber;
        private string policyNumber;

        public Client(string Cli_firstName, string Cli_Surname, string Cli_IdNumber, string Cli_PolicyNumber)
        {
            firstName = Cli_firstName;
            surname = Cli_Surname;
            idNumber = Cli_IdNumber;
            policyNumber = Cli_PolicyNumber;
        }

        public string FirstName
        {
            get { return firstName; }
            set { firstName = value; }
        }

        public string Surname
        {
            get { return surname; }
            set { surname = value; }
        }

        public string IdNumber
        {
            get { return idNumber; }
            set { idNumber = value; }
        }

        public string PolicyNumber
        {
            get { return policyNumber; }
            set { policyNumber = value; }
        }

        public override string ToString()
        {
            return $"{firstName }  {surname }  {idNumber}  {policyNumber}";
        }
    }
}

