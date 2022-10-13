using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Random_Credentials_Generator
{
    /// <summary>
    /// The purpose of this class is to create an interactive menu based on the options and
    /// title provided, by displaying the information and allowing the user to enter
    /// selections based on the options shown. Checks are also done to ensure good choices
    /// are made.
    /// </summary>
    internal class Menu
    {
        private string title;
        private string[] options;

        public Menu(string title, string[] options)
        {
            this.title = title;
            this.options = options;
        }

        private void Display()
        {
            Console.WriteLine(title);
            for (int i = 0; i < title.Length; i++)
            {
                Console.Write("+");
            }
            Console.WriteLine();
            for (int j = 0; j < options.Length; j++)
            {
                Console.WriteLine((j+1) + ") " + options[j]);
            }
            Console.WriteLine();
        }

        public int GetUserChoice()
        {
            int value = 0;
            Display();
            do {
                try {
                    Console.WriteLine("Enter Selection: ");
                    value = Convert.ToInt32(Console.ReadLine());
                    break;
                }
                catch(Exception ex) {
                    Console.WriteLine("Invalid selection entered. Please try again.");
                }
            } while (true);

            return value;
        }
    }
}
