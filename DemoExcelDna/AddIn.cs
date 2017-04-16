using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace DemoExcelDna {
    public class AddIn : IExcelAddIn {
        public void AutoOpen() {
            IntelliSenseServer.Register();
        }

        public void AutoClose() {
        }
    }
}