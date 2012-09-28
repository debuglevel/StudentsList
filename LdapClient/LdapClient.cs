using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Debuglevel.LdapCLient
{
    /// <summary>
    /// LDAP Client to get user information
    /// </summary>
    public class LdapClient
    {
        private string username;
        public SecureString password;
        public string ldapServer;

        /// <summary>
        /// Initializes the LDAP Client
        /// </summary>
        /// <param name="server">LDAP server</param>
        /// <param name="username">user</param>
        /// <param name="password">password</param>
        public LdapClient(string server, string username, SecureString password)
        {
            this.ldapServer = server;
            this.username = username;
            this.password = password;
        }

        /// <summary>
        /// gets the LDAP directory entry for a path
        /// </summary>
        /// <param name="baseTree">base of the ldap path</param>
        /// <param name="subtreeAttributes">additional path elements</param>
        /// <returns>directory entry</returns>
        private DirectoryEntry getDirectoryEntry(string baseTree, params string[] subtreeAttributes)
        {
            string subtree = String.Join(",", subtreeAttributes.Reverse());

            if (subtreeAttributes.Any())
            {
                subtree += ",";
            }

            string ldapPath = String.Format("{0}{1}", subtree, baseTree);

            string connectionPath = String.Format("LDAP://{0}/{1}", this.ldapServer, ldapPath);

            DirectoryEntry entry = new DirectoryEntry(connectionPath, this.username, convertSecureStringToString(this.password));
            return entry;
        }

        /// <summary>
        /// converts a SecureString back to a String
        /// </summary>
        /// <param name="secureString">the SecureString</param>
        /// <returns>content of the SecureString</returns>
        private string convertSecureStringToString(SecureString secureString)
        {
            IntPtr ptr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(this.password);
            return System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr);
        }

        /// <summary>
        /// get all users
        /// </summary>
        /// <param name="baseTree">the base LDAP tree</param>
        /// <param name="subtreeAttributes">subtree to append to the baseTree. array [0],[1],[2] becomes [2],[1],[0],tree </param>
        /// <returns>all users</returns>
        public SearchResultCollection GetUsers(string baseTree, params string[] subtreeAttributes)
        {
            DirectoryEntry entry = getDirectoryEntry(baseTree, subtreeAttributes);

            DirectorySearcher mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = "(&(objectCategory=user))";
            SearchResultCollection results = mySearcher.FindAll();

            return results;
        }

    }
}

