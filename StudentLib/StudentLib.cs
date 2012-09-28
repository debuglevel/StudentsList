using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Debuglevel.LdapCLient;

namespace Debuglevel.StudierendeLib
{
    /// <summary>
    /// Interface to the LDAP server
    /// </summary>
    public class Studierende
    {
        /// <summary>
        /// the LDAP Client to establish a connection to the Active Directory and fetch the information
        /// </summary>
        private LdapClient ldapClient;

        /// <summary>
        /// the base of the LDAP tree
        /// </summary>
        private readonly string baseTree = "OU=HS Studenten,OU=HS Users,OU=HS,DC=ads,DC=hs-karlsruhe,DC=de";

        /// <summary>
        /// the faculty of the university we're interested in
        /// </summary>
        public string Faculty { get; set; }

        /// <summary>
        /// the programs within the faculty we're interested in
        /// </summary>
        public List<string> Programs { get; private set; }

        /// <summary>
        /// Initializing this interface, the LDAP client and internal stuff
        /// </summary>
        /// <param name="faculty">the university's faculty we're interested in</param>
        /// <param name="programs">the faculty's programs we're interested in</param>
        public Studierende(string username, SecureString password, string ldapServer, string faculty, List<string> programs)
            : this(username, password, ldapServer)
        {
            this.Faculty = faculty;
            this.Programs.AddRange(programs);
        }

        /// <summary>
        /// Initializing the LDAP client and internal stuff
        /// </summary>
        /// <param name="username">user within the AD</param>
        /// <param name="password">the user's password</param>
        /// <param name="ldapServer">the AD</param>
        public Studierende(string username, SecureString password, string ldapServer)
        {
            this.ldapClient = new LdapClient(ldapServer, username, password);

            this.Programs = new List<string>();
        }

        /// <summary>
        /// get the subtree OU-attribute for a faculty
        /// </summary>
        /// <param name="faculty">the university's faculty</param>
        /// <returns>subtree OU-attribute for the faculty</returns>
        private string getFacultySubtree(string faculty)
        {
            return "OU=_fk_" + faculty;
        }

        /// <summary>
        /// get the subtree OU-attribute for a program
        /// </summary>
        /// <param name="faculty">the faculty's program</param>
        /// <returns>subtree OU-attribute for the program</returns>
        private string getProgramSubtree(string program)
        {
            return "OU=_sg_" + program;
        }

        /// <summary>
        /// gets all students of all programs in the faculty
        /// </summary>
        /// <returns>all students of all programs in the faculty</returns>
        private IEnumerable<SearchResult> getStudents()
        {
            var results = new List<SearchResult>();

            foreach (var program in this.Programs)
            {
                string[] subTree = new string[] {this.getFacultySubtree(this.Faculty), this.getProgramSubtree(program)};
                results.AddRange(ldapClient.GetUsers(this.baseTree, subTree).Cast<SearchResult>());
            }

            return results;
        }

        /// <summary>
        /// gets the program of a student from the LDAP tree
        /// </summary>
        /// <param name="tree">the LDAP tree</param>
        /// <returns>the student's program</returns>
        private string getProgramFromTree(string tree)
        {
            foreach (var program in this.Programs)
            {
                if (tree.Contains(this.getProgramSubtree(program)) == true)
                {
                    return program;
                }
            }

            return "unknown";
        }

        /// <summary>
        /// returns all students as CSV
        /// </summary>
        /// <returns>students as CSV</returns>
        public string AsCSV()
        {
            var items = this.getStudents();

            var results = items.Select(i => new
            {
                Mail = i.Properties["mail"][0],
                Vorname = i.Properties["msDS-PhoneticFirstName"][0],
                Nachname = i.Properties["msDS-PhoneticLastName"][0],
                Username = i.Properties["name"][0],
                ErstellungsDatum = DateTime.Parse(i.Properties["whenCreated"][0].ToString()),
                Studiengang = getProgramFromTree(i.Properties["memberOf"][0].ToString())
            });

            var csvContext = new LINQtoCSV.CsvContext();

            var writer = new System.IO.StringWriter();
            csvContext.Write(results, writer);
            writer.Close();

            return writer.ToString();
        }

    }
}
