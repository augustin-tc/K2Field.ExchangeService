using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using K2Field.ExchangeService;
using System.Data;

namespace ExchangeServiceTests
{
    /// <summary>
    /// Description résumée pour PowerShellTest
    /// </summary>
    [TestClass]
    public class PowerShellTest
    {
        public PowerShellTest()
        {
            //
            // TODO: ajoutez ici la logique du constructeur
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Obtient ou définit le contexte de test qui fournit
        ///des informations sur la série de tests active, ainsi que ses fonctionnalités.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Attributs de tests supplémentaires
        //
        // Vous pouvez utiliser les attributs supplémentaires suivants lorsque vous écrivez vos tests :
        //
        // Utilisez ClassInitialize pour exécuter du code avant d'exécuter le premier test de la classe
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Utilisez ClassCleanup pour exécuter du code une fois que tous les tests d'une classe ont été exécutés
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Utilisez TestInitialize pour exécuter du code avant d'exécuter chaque test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Utilisez TestCleanup pour exécuter du code après que chaque test a été exécuté
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void PSGetPermissionTest()
        {
            DataTable dtResults = new DataTable();
            foreach (string colName in new String[] { ""})
            {
                dtResults.Columns.Add(colName);
            }
            ExchangeManagementHelper mgtHelper = new ExchangeManagementHelper("", "");
           dtResults = mgtHelper.GetMailBoxPermissions("", true, dtResults);
        }
    }
}
