using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace AudienceTargetting
{
    class Program
    {
        static void Main(string[] args)
        {
            string webSPOUrl = "https://psrsolutions1.sharepoint.com/sites/Dev/";
            string userName = "ramakrishna@psrsolutions1.onmicrosoft.com";
            Console.WriteLine("Password:");
            SecureString password = FetchPasswordFromConsole();
            try
            {
                using (var context = new ClientContext(webSPOUrl))
                {
                    context.Credentials = new SharePointOnlineCredentials(userName, password);
                    Web web = context.Web;
                    var list = context.Web.Lists.GetByTitle("Site Pages");
                    context.Load(list);
                    var fieldSchemaXml = @"<Field ID='{7f759147-c861-4cd6-a11f-5aa3037d9634}' Type='UserMulti' List='UserInfo' Name='_ModernAudienceTargetUserField' StaticName='_ModernAudienceTargetUserField' DisplayName='Audience' Required='FALSE' SourceID='{e533a454-bc7a-4e33-954c-f7c9c17244ef}' ColName='int2' RowOrdinal='0' ShowField='ImnName' ShowInDisplayForm='TRUE' ShowInListSettings='FALSE' UserSelectionMode = 'GroupsOnly' UserSelectionScope = '0' Mult = 'TRUE' Sortable = 'FALSE' Version = '1' /> ";

                    var field = list.Fields.AddFieldAsXml(fieldSchemaXml, true, AddFieldOptions.AddFieldInternalNameHint);
                    context.ExecuteQuery();

                    // Get the content type from list
                    var ContentTypeColl = list.ContentTypes;
                    context.Load(ContentTypeColl);
                    context.ExecuteQuery();

                    foreach (var contentType in ContentTypeColl)
                    {
                        Console.WriteLine(contentType.Name);
                        var ColumnColl = list.Fields;
                        context.Load(ColumnColl);
                        context.ExecuteQuery();
                        var ColumnName = "Audience";
                        var Column = ColumnColl.Where(c => c.Title == ColumnName).FirstOrDefault();

                        if (Column != null)
                        {
                            //Check if column already added to the content type
                            var FieldCollection = contentType.Fields;
                            context.Load(FieldCollection);
                            context.ExecuteQuery();
                           var Field = FieldCollection.Where(f => f.Title == ColumnName).FirstOrDefault();
                            if (Field == null && !(contentType.Sealed))
                            {
                                //Add Field to content type
                                var FieldLink = new FieldLinkCreationInformation();
                                FieldLink.Field = Column;
                                contentType.FieldLinks.Add(FieldLink);
                                contentType.Update(false);
                                context.ExecuteQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error is : " + ex.Message);
            }
        }

        private static SecureString FetchPasswordFromConsole()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        password = password.Substring(0, password.Length - 1);
                        int pos = Console.CursorLeft;
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            Console.WriteLine();
            var securePassword = new SecureString();
            //Convert string to secure string  
            foreach (char c in password)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}
