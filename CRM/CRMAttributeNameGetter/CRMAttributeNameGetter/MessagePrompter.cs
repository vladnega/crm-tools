using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRMAttributeNameGetter
{
    public class MessagePrompter
    {
        public MessagePrompter()
        {

        }

        public void Prompt(string format, params object[] args)
        {
            Console.WriteLine(format, args);
        }

        public string Read()
        {
            return Console.ReadLine();
        }

        public string ReadPassword()
        {
            var Password = "";

            ConsoleKeyInfo key;
            do
            {
                key = Console.ReadKey(true);
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    Password += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && Password.Length > 0)
                    {
                        Password = Password.Substring(0, (Password.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            } while (key.Key != ConsoleKey.Enter);

            return Password;
        }
    }
}
