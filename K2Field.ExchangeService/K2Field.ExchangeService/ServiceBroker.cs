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

            return base.DescribeSchema();
        }

        public override void Execute()
        {
            //récupération service appellé 
            ServiceObject calledSvo = this.Service.ServiceObjects[0];
            //récupération méthode 
            Method method = calledSvo.Methods[0];



            //création générique table de retour
            DataTable dtResults = new DataTable();
            for (int i = 0; i < method.ReturnProperties.Count; i++)
            {
                dtResults.Columns.Add(method.ReturnProperties[i]);
            }
            if (calledSvo.Name == "Exchange Membership")
            {
                string userName = Service.ServiceConfiguration.ServiceAuthentication.UserName;
                string password = Service.ServiceConfiguration.ServiceAuthentication.Password;
                ExchangeHelper helper = new ExchangeHelper(userName, password);
                if (method.Name == "GetDistributionListMembers")
                {
                    string listName = method.MethodParameters["Name"].Value.ToString();
                    dtResults = helper.getDistributionListMembers(listName, dtResults);
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
