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
        /// <param name="column">Column's Index Position</param>
        /// <returns>
        /// Successful - The cell range address<br></br>
        /// Unsuccessful - An error string value used for error checking
        /// </returns>
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



        /// <summary>
        /// The purpose of this function is to return the cell address of the first blank cell
        /// to appear within the cell address range provided
        /// </summary>
        /// <param name="sheet">The Excel Sheet within the Workbook</param>
        /// <param name="column">The Column's Index Position</param>
        /// <returns>
        /// Successful - The cell address of the first blank cell within the given cell address Range<br></br>
        /// Unsuccessful - An error string value used for error checking (comes from the GetAddressRange Function)
        /// </returns>
        public static string GetBlankCell(WorkSheet sheet, int column)
        {
            // Stores the cell address of the blank cell
            String cellBlank = "";

            // Stores the address cell range for the excel sheet
            String address = GetAddressRange(column);

            // If the address cell range lookup function didn't fail
            if (address != "Failed Lookup")
            {
                // Loop through each individual cell within the provided range
                foreach (var cell in sheet[address])
                {
                    // If the cell is blank
                    if (cell.Text == "")
                    {
                        // Store the blank cell's address into the variable and end the loop
                        cellBlank = cell.AddressString;
                        break;
                    }
                }
                return cellBlank;
            }
            // Return error string
            // This will be looked for to ensure no errors have occured during this process
            return "Failed Lookup";
        }



        /// <summary>
        /// This purpose of this function is to take the name you wish to add and add it
        /// to the Excel Spreadsheet to extend the list of names found there. This method will
        /// automatically check to make sure that the name isn't already present within the column
        /// you're trying to add it to. If the name isn't in the list already, then it will be added
        /// </summary>
        /// <param name="workbook">The Excel Workbook</param>
        /// <param name="sheet">The Excel Sheet within the Workbook</param>
        /// <param name="name">The Name to be Added</param>
        /// <param name="column">The Column's Index Position</param>
        public static void AddName(WorkBook workbook, WorkSheet sheet, string name, int column)
        {
            // Stores the address cell range for the excel sheet
            string address = GetAddressRange(column);

            // If the address cell range lookup function didn't fail
            if (address != "Failed Lookup")
            {
                // Loop through each individual cell within the provided range
                foreach (var cell in sheet[address])
                {
                    // If the name already exists
                    if (cell.Text == name)
                    {
                        // Let the user know by displaying a message
                        Console.WriteLine("Name unsuccessfully added to the list, as it already exists!");
                        return;
                    }
                }
                // If the name didn't exist in the excel spreadsheet
                // Find the first blank cell within the range provided
                string cellEdit = GetBlankCell(sheet, column);

                // Change the blank cell's value to the name value
                sheet[cellEdit].Value = name;

                // Overwrite the Excel Workbook with the new one with the changes made
                workbook.SaveAs("Names.xlsx");

                // Let the user know by displaying a message that the addition was successful
                Console.WriteLine("Name successfully added to the list!");
                return;
            }
            // If the address cell range lookup function failed. Just return an
            // error message letting the user know why the code failed the way it did
            Console.WriteLine("Incorrect Column value entered");
        }



    }
}
