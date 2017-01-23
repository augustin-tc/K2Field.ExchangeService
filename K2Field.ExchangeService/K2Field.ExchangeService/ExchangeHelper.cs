using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.ExchangeService
{
    class ExchangeHelper
    {
        String UserName { get; set; }
        String Password { get; set; }

        public ExchangeHelper(string userName, string password)
        {
            UserName = userName;
            Password = password;
        }

        public DataTable getDistributionListMembers(string name, DataTable dtResults)
        {


           Microsoft.Exchange.WebServices.Data.ExchangeService service =
                new Microsoft.Exchange.WebServices.Data.ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(UserName, Password);
            service.AutodiscoverUrl(UserName, RedirectionUrlValidationCallback);

            // Return the expanded group.
            ExpandGroupResults groupMembers = service.ExpandGroup(name);

            foreach(EmailAddress address in groupMembers)
            {
                DataRow dr = dtResults.NewRow();
                dr["Email"] = address.Address;
                dr["FullName"] = address.Name;
                dtResults.Rows.Add(dr);
            }
            return dtResults;
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }



    }
}
