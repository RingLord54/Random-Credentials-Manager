using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace Random_Credentials_Generator
{
    public class IronXL_Utilities
    {

        /// <summary>
        /// The purpose of this function is to take an integer parameter which should
        /// either be 1 or 0 and will then return the cell range address to use to search 
        /// through the excel workbook sheet with. 
        /// </summary>
        /// <param name="column">Column's index position within the Excel Spreadsheet (A = 0, B = 1)</param>
        /// <returns>The cell range address</returns>
        public static string GetAddressRange(int column)
        {
            // Will store the Column letter
            String letter;
            switch (column)
            {
                case 0: letter = "A"; break;
                case 1: letter = "B"; break; 

                // Return error string
                // This will be looked for to ensure no errors have occured during this process
                default: return "Failed Lookup";
            }

            // String Address is formatted with the letter value
            String address = letter + "2:" + letter + "1000";
            return address;
        }


        public static string GetBlankCell(WorkSheet sheet, int column)
        {
            String cellBlank = "";
            String address = GetAddressRange(column);
            if (address != "Failed Lookup")
            {
                foreach (var cell in sheet[address])
                {
                    if (cell.Text == "")
                    {
                        cellBlank = cell.AddressString;
                        break;
                    }
                }
                return cellBlank;
            }
            return "Failed Lookup";
        }


        public static void AddName(WorkBook workbook, WorkSheet sheet, string name, int column)
        {
            string address = GetAddressRange(column);

            if (address != "Failed Lookup")
            {
                foreach (var cell in sheet[address])
                {
                    if (cell.Text == name)
                    {
                        Console.WriteLine("Name unsuccessfully added to the list!");
                        return;
                    }
                }
                string cellEdit = GetBlankCell(sheet, column);

                sheet[cellEdit].Value = name;

                workbook.SaveAs("Names.xlsx");

                Console.WriteLine("Name successfully added to the list!");

                return;
            }
            Console.WriteLine("Incorrect Column value entered");
        }


    }
}
