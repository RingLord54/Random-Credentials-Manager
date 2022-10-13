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
        private string title;       // Menu Title
        private string[] options;   // Menu Options


        /// <summary>
        /// The menu class' constructor
        /// </summary>
        /// <param name="title">Menu Title</param>
        /// <param name="options">Menu Options</param>
        public Menu(string title, string[] options)
        {
            this.title = title;
            this.options = options;
        }


        /// <summary>
        /// The purpose of this function is to display the title and the options
        /// of the menu provided in an appealing and organised way, to make it seem
        /// more professional to the user.
        /// </summary>
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


        /// <summary>
        /// The purpose of this function is to display the menu's title and options and
        /// then provide the User the option to input a selection to then be processed
        /// in order to carry out an operation. A try catch is implemented to ensure integer
        /// values only are entered.
        /// </summary>
        /// <returns>The number value of the selected menu option</returns>
        public int GetUserChoice()
        {
            int value = 0;
            Display();
            do 
            {
                // Try to get a valid user selection
                try 
                {
                    Console.Write("Enter Selection: ");
                    value = Convert.ToInt32(Console.ReadLine());

                    // If the entered value is valid break from the selection loop
                    break;
                }
                // If the input isn't an integer value
                catch(Exception ex) 
                {
                    // Display a message notifying the User of what went wrong
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Invalid selection entered. Please try again.\n");
                }
            // Keep going until a valid selection is made by the user
            } while (true);

            return value;
        }
    }
}
