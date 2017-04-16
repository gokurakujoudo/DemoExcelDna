using ExcelDna.Integration;
using MathFuncConsole.MathObjects.Helper;

namespace DemoExcelDna
{
    internal static class ExcelFunc
    {

        internal static double NormCdf(double z) =>
            XlCall.Excel(XlCall.xlfNormsdist, z).To<double>();

        internal static object[,] ToColumn(this object[] input)
        {
            var l = input.Length;
            var output = new object[l, 1];
            for (var i = 0; i < l; i++)
                output[i, 0] = input[i];
            return output;
        }

        internal static object Transpose([ExcelArgument(AllowReference = true)] object input) =>
            XlCall.Excel(XlCall.xlfTranspose, input);
    }
}
