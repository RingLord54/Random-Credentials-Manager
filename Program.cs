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
            // Fetches the Excel workbook called "Names.xlsx" from the local files
            WorkBook workbook = WorkBook.Load("Names.xlsx");

            // Fetches the Excel spreadsheet called "Sheet1" within the workbook
            WorkSheet sheet = workbook.GetWorkSheet("Sheet1");

            // Menu Title
            String title = "Random Credentials Generator";

            // Menu Options
            String[] options =
            {
                "Add Name to the Names Excel Spreadsheet",
                "Generate Random Credentials",
                "Exit Application"
            };

            // Will store the User's selection on the menu
            int choice;

            // Create Menu Object instance
            Menu menu = new Menu(title, options);

            // Keep going until the User wishes to close the application
            do
            {
                choice = menu.GetUserChoice();
                if (choice == options.Length)
                {
                    break;
                }
                Console.WriteLine("\nSelection made: " + choice + "\n");
            } while (true);

            // Close Application
            Console.WriteLine("\nExiting Application...\n");
            Console.WriteLine("Thank you for using the Random Credentials Generator");
        }
    }
}
