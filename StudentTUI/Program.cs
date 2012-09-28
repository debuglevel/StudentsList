using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Debuglevel.StudierendeLib;

namespace Debuglevel.StudierendenListe
{
    /// <summary>
    /// Text User Interface
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Text User Interface entry point
        /// </summary>
        /// <param name="args">unused</param>
        static void Main(string[] args)
        {
            string domain;
            string username;
            SecureString password;
            getCredentials(out domain, out username, out password);
            tryClear();

            Console.Error.WriteLine("Processing... please wait...");

            var faculty = "IWI";
            var program = new List<string>() {"WIM", "WIB"};

            Studierende studierende = new Studierende(username, password, domain, faculty, program);
            var csv = studierende.AsCSV();
            tryClear();

            Console.WriteLine(csv);

            Console.ReadLine();
        }

        /// <summary>
        /// prompt for credentials
        /// </summary>
        /// <param name="domain">the domain of the active directory</param>
        /// <param name="username">an user within the active directory</param>
        /// <param name="password">the users password</param>
        private static void getCredentials(out string domain, out string username, out SecureString password)
        {
            Console.Error.Write("Domain: ");
            domain = Console.ReadLine();

            Console.Error.Write(@"User: ");
            username = Console.ReadLine();

            if (username.Contains('\\') == false)
            {
                username = String.Format(@"{0}\{1}", domain, username);
                Console.Error.WriteLine(@"I extended your user to " + username);
            }

            Console.Error.Write(@"Password: ");
            password = getPassword();
        }

        private static void tryClear()
        {
            try
            {
                Console.Clear();
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// prompt for password in a secrure way
        /// </summary>
        /// <returns></returns>
        private static SecureString getPassword()
        {
            SecureString pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }

    }
}
