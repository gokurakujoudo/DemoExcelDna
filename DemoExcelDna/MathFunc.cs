using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using MathFuncConsole.MathObjects;
using MathFuncConsole.MathObjects.Applications;
using MathFuncConsole.MathObjects.Helper;
using Microsoft.Office.Interop.Excel;

namespace DemoExcelDna {

    public static class MathFunc {
        public static readonly Dictionary<string, MathObject> Environment = new Dictionary<string, MathObject>();
        private static readonly ObjectViewer ObjectViewer = new ObjectViewer(Environment);

        [ExcelFunction("Add or modify a stock", Category = "MathFunc.Asset", IsVolatile = true)]
        public static string Stock([ExcelArgument("MathObject Id")] string id,
                                   [ExcelArgument("Current price")] double price,
                                   [ExcelArgument("Sigma of return")] double sigma,
                                   [ExcelArgument("Drift of return")] double mu = 0,
                                   [ExcelArgument("Dividend yield")] double divd = 0) {
            Environment[id] = new Stock(id, price, sigma, mu, divd);
            return id;
        }

        [ExcelFunction("Add or modify an exchange option", Category = "MathFunc.Asset", IsVolatile = true)]
        public static string ExchangeOption([ExcelArgument("MathObject Id")] string id,
                                            [ExcelArgument("Id of stock 1")] string stock1Id,
                                            [ExcelArgument("Id of stock 2")] string stock2Id,
                                            [ExcelArgument("Correlation of two stocks")] double rho,
                                            [ExcelArgument("Option maturity")] double maturity) {
            var stock1 = Environment[stock1Id].To<Stock>();
            var stock2 = Environment[stock2Id].To<Stock>();
            Environment[id] = new ExchangeOption(id, stock1, stock2, rho, maturity);
            return id;
        }

        [ExcelFunction("Open Object Viewer")]
        public static string ShowForm() {
            ObjectViewer.Show();
            return "Open Object Viewer";
        }



        [ExcelFunction("Show a MathObject", Category = "MathFunc.MathObject", IsVolatile = true)]
        public static string ShowObj([ExcelArgument("MathObject Id")] string id) {
            if (Environment.ContainsKey(id))
                return Environment[id].ToString();
            return $"{id} cannot be found";
        }

        [ExcelFunction("A template for stock (6x1)", Category = "MathFunc.Asset_Template")]
        public static object[,] StockTemp() {
            return new object[] {"Id", "Price", "Sigma", "[Mu]", "[Divd]", "Stock =>"}.ToColumn();
        }




        [ExcelFunction("A template for exchange option (6x1)", Category = "MathFunc.Asset_Template")]
        public static object[,] ExchangeOptionTemp() =>
            new object[] {"Id", "Stock1", "Stock2", "Rho", "T", "Option =>"}.ToColumn();


        [ExcelFunction("Show a property of a MathObject", Category = "MathFunc.MathObject", IsVolatile = true)]
        public static object ShowProperty([ExcelArgument("MathObject Id")] string id,
                                          [ExcelArgument("Property name")] string property) {
            if (!Environment.ContainsKey(id))
                return $"{id} cannot be found";
            return Environment[id].RemoteGetter(property)();
        }



    }
}
