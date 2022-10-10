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


        public static string GetAddressRange(int column)
        {
            String value = "";
            switch (column)
            {
                case 0: value = "A"; break;
                case 1: value = "B"; break;
                default: return "Failed Lookup";
            }
            String address = value + "2:" + value + "1000";
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
