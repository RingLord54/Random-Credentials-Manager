using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace Random_Credentials_Generator
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            WorkBook workbook = WorkBook.Load("Names.xlsx");
            WorkSheet sheet = workbook.GetWorkSheet("Sheet1");

            Console.WriteLine("Name Generator Application");
            Console.WriteLine("------------------------------------------------------------");
            Console.WriteLine("\nWelcome to the Name Generator Application!\n");
            Console.Write("Would you like to:\n- Add a name to the list of names that can be selected (Enter 0)\n- Randomly Generate a Name (Enter 1)\n");
            Console.Write("\nInput: ");
            int num = Convert.ToInt32(Console.ReadLine());

            Console.WriteLine("\nYour choice: " + num + "\n");

            Console.WriteLine("------------------------------------------------------------");
            Console.WriteLine("Thank you for using the Name Generator Application!");

            //IronXL_Utilities.AddName(workbook, sheet, "Bonny", 1);
        }
    }
}
