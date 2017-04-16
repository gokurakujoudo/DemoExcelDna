using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using MathFuncConsole.MathObjects;
using MathFuncConsole.MathObjects.Helper;

namespace DemoExcelDna
{
    public partial class ObjectViewer : Form
    {
        public ObjectViewer(Dictionary<string, MathObject> env) {
            _env = env;
            this.InitializeComponent();
            this.Activated += (sender, e) => this.RefreshObjects();
        }

        private void button1_Click(object sender, System.EventArgs e) {
            var temp = listBox1.SelectedItem.To<MathObject>();
            if (temp == null) return;
            _env.Remove(temp.Name);
            this.RefreshObjects();
        }

        private void RefreshObjects() {
            listBox1.Items.Clear();
            listBox1.Items.AddRange(_env.Values.ToArray());
        }
    }
}
