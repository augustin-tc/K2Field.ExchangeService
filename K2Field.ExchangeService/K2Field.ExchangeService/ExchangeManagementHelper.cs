using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.ExchangeService
{
    public class ExchangeManagementHelper
    {
        String UserName { get; set; }
        String Password { get; set; }

        public ExchangeManagementHelper(string userName, string password)
        {
            UserName = userName;
            Password = password;
        }

        public DataTable GetMailBoxPermissions(string mailBox, bool includeAll, DataTable dtResults)
        {

           
            SecureString securestring = new SecureString();

            foreach (char pass in Password)
            {
                securestring.AppendChar(pass);
            }

            PSCredential pscred = new PSCredential(UserName, securestring);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("https://ps.outlook.com/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pscred);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.MaximumConnectionRedirectionCount = 2;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                runspace.Open();
                using (PowerShell powershell = PowerShell.Create())
                {
                    powershell.Runspace = runspace;
                  
                    powershell.AddCommand("Get-MailboxPermission");
                    powershell.AddParameter("Identity", mailBox);
                                  
                    var results = powershell.Invoke();

                    foreach (PSObject result in results)
                    {
                        string user = "";
                        string accessRights = "";
                        bool deny = false;
                        bool isInherited = false;

                        if (result.Properties["User"].Value != null)
                            user = result.Properties["User"].Value.ToString();
                        if (result.Properties["AccessRights"].Value != null)
                            accessRights = result.Properties["AccessRights"].Value.ToString();
                        if (result.Properties["Deny"].Value != null)
                            deny = bool.Parse(result.Properties["Deny"].Value.ToString());
                        if (result.Properties["IsInherited"].Value != null)
                            isInherited = bool.Parse(result.Properties["IsInherited"].Value.ToString());

                        if (includeAll || (!includeAll && (!isInherited && !deny && user!= "NT AUTHORITY\\SELF")))
                        {
                            DataRow row = dtResults.NewRow();
                            row["User"] = user;
                            row["AccessRights"] = accessRights;
                            row["Deny"] = deny;
                            row["IsInherited"] = isInherited;
                            dtResults.Rows.Add(row);
                        }

                    }
                }
                runspace.Close();
                return dtResults;
            }

            #region old
            /*
            using (Runspace runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace())
            {

                using (PowerShell PowerShellInstance = PowerShell.Create())
                {
                    
                    runspace.Open();
                    PowerShellInstance.Runspace = runspace;

                    PowerShellInstance.AddScript(string.Format("$secpasswd = ConvertTo-SecureString '{0}' -AsPlainText -Force", Password));
                    PowerShellInstance.AddScript(string.Format("$LiveCred = New-Object System.Management.Automation.PSCredential '{0}', $secpasswd", UserName));
                    PowerShellInstance.AddScript("$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $LiveCred -Authentication Basic -AllowRedirection");
                    //PowerShellInstance.AddScript("Import-PSSession $Session");
                    PowerShellInstance.AddScript("Get-MailBox");
                    PowerShellInstance.AddScript(string.Format("Get-MailboxPermission -Identity '{0}' | Where {{ ($_.IsInherited -eq $False) -and -not($_.User -like 'NT AUTHORITY\\SELF') }} | FT User,AccessRights", mailBox));
                    PowerShellInstance.AddScript("Remove-PSSession $Session");
                    
                    Collection<PSObject> PSOutput = PowerShellInstance.Invoke();

                    //debug
                    string commands = "";
                    foreach (var cmd in PowerShellInstance.Commands.Commands)
                    {
                        commands += cmd.ToString() + "\n";
                    }

                    foreach (PSObject outputItem in PSOutput)
                    {

                        if (outputItem != null)
                        {

                        }
                    }
                    if (PowerShellInstance.Streams.Error.Count > 0)
                    {
                        // error records were written to the error stream.
                        // do something with the items found.
                        StringBuilder strErrors = new StringBuilder();
                        foreach (var error in PowerShellInstance.Streams.Error)
                        {
                            strErrors.AppendLine(error.ToString());
                        }
                        string errors = strErrors.ToString();

                    }
                }
            }*/
            #endregion
        }
    }

}
