using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.ExchangeService
{
    class ServiceBroker : ServiceAssemblyBase
    {
        public override string GetConfigSection()
        {
            return base.GetConfigSection();
        }
        public override string DescribeSchema()
        {
            if (Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode != AuthenticationMode.Static)
            {
                throw new Exception("Error : the authentication mode must be set to Static.");
            }
            #region Membership Service
            ServiceObject svo = new ServiceObject("Exchange Membership");
            svo.MetaData.DisplayName = "Exchange Membership";
            svo.Active = true;

            svo.Properties.Add(CreateProperty("Email", SoType.Text));
            svo.Properties.Add(CreateProperty("FullName", SoType.Text));
            MethodParameter nameParameter = CreateParameter("Name", SoType.Text, true);
            Method GetDistributionListMembers = CreateMethod("GetDistributionListMembers", MethodType.List);

            GetDistributionListMembers.MethodParameters.Add(nameParameter);
            GetDistributionListMembers.ReturnProperties.Add("Email");
            GetDistributionListMembers.ReturnProperties.Add("FullName");


            svo.Methods.Create(GetDistributionListMembers);

            Service.ServiceObjects.Add(svo);
            #endregion

            #region exchange management
            ServiceObject svoMgt = new ServiceObject("Exchange Management");
            svoMgt.MetaData.DisplayName = "Exchange Management";
            svoMgt.Active = true;
            /*
            row[""] = user;
            row[""] = accessRights;
            row["Deny"] = deny;
            row["IsInherited"] = isInherited;*/


            svoMgt.Properties.Add(CreateProperty("User", SoType.Text));
            svoMgt.Properties.Add(CreateProperty("AccessRights", SoType.Text));
            svoMgt.Properties.Add(CreateProperty("Deny", SoType.YesNo));
            svoMgt.Properties.Add(CreateProperty("IsInherited", SoType.YesNo));
                

            MethodParameter MailBoxNameParameter = CreateParameter("MailBoxName", SoType.Text, true);
            MethodParameter includeAllParam = CreateParameter("IncludeAll", SoType.YesNo, true);
            Method GetMailBoxPermissions = CreateMethod("GetMailBoxPermissions", MethodType.List);

            GetMailBoxPermissions.MethodParameters.Add(MailBoxNameParameter);
            GetMailBoxPermissions.MethodParameters.Add(includeAllParam);
            GetMailBoxPermissions.ReturnProperties.Add("User");
            GetMailBoxPermissions.ReturnProperties.Add("AccessRights");
            GetMailBoxPermissions.ReturnProperties.Add("Deny");
            GetMailBoxPermissions.ReturnProperties.Add("IsInherited");

            svoMgt.Methods.Create(GetMailBoxPermissions);

            Service.ServiceObjects.Add(svoMgt);
            #endregion exchange management


            return base.DescribeSchema();
        }

        public override void Execute()
        {

            //get currently called service
            ServiceObject calledSvo = this.Service.ServiceObjects[0];
            //get currently called method 
            Method method = calledSvo.Methods[0];
            string userName = Service.ServiceConfiguration.ServiceAuthentication.UserName;
            string password = Service.ServiceConfiguration.ServiceAuthentication.Password;
            //generic results dataTable
            DataTable dtResults = new DataTable();
            for (int i = 0; i < method.ReturnProperties.Count; i++)
            {
                dtResults.Columns.Add(method.ReturnProperties[i]);
            }
            if (calledSvo.Name == "Exchange Membership")
            {
                
                ExchangeHelper helper = new ExchangeHelper(userName, password);
                if (method.Name == "GetDistributionListMembers")
                {
                    string listName = method.MethodParameters["Name"].Value.ToString();
                    dtResults = helper.getDistributionListMembers(listName, dtResults);
                }
            }
            else if (calledSvo.Name == "Exchange Management")
            {
                if (method.Name == "GetMailBoxPermissions")
                {
                    ExchangeManagementHelper helper = new ExchangeManagementHelper(userName, password);
                    string mailBoxName = method.MethodParameters["MailBoxName"].Value.ToString();
                    bool includeAll = bool.Parse(method.MethodParameters["IncludeAll"].Value.ToString());

                     dtResults = helper.GetMailBoxPermissions(mailBoxName, includeAll, dtResults);
                }
            }
               
            calledSvo.Properties.InitResultTable();
            foreach (DataRow dr in dtResults.Rows)
            {
                foreach (DataColumn column in dtResults.Columns)
                {
                    if (dr != null)
                        calledSvo.Properties[column.ColumnName].Value = dr[column].ToString();
                }
                calledSvo.Properties.BindPropertiesToResultTable();
            }
            
        }


        public override void Extend()
        {
            throw new NotImplementedException();
        }

        Property CreateProperty(string Name, SoType soType)
        {
            Property prop = new Property(Name);
            prop.MetaData.DisplayName = Name;
            prop.SoType = soType;
            return prop;
        }

        MethodParameter CreateParameter(string Name, SoType soType, bool isRequired)
        {
            MethodParameter parameter = new MethodParameter();
            parameter.Name = Name;
            parameter.MetaData.DisplayName = Name;
            parameter.IsRequired = isRequired;
            return parameter;
        }

        Method CreateMethod(string Name, MethodType type)
        {
            Method method = new Method(Name);
            method.MetaData.DisplayName = Name;
            method.Type = type;
            return method;
        }
    }
}
