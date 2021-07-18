using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace expert_opinion
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var helper = new WordHelper("expert_opinion_template.docx");
                var items = new Dictionary<string, string>
                {
                    {"<type_of_work>", comboBox_type_of_work.Text },
                    {"<title>", textBox_title.Text },
                    {"<gost_type_of_flowmeter>", comboBox_gost_type_of_flowmeter.Text },
                    {"<main_source>", textBox_main_source.Text },
                    {"<second_source>", textBox_second_source.Text },
                    {"<developer_company>", textBox_developer_company.Text },
                    {"<flow_meter>", comboBox_flow_meter.Text },
                    {"<dn>", textBox_dn.Text }
                };
                helper.Process(items);
                
                
            }
            catch (IOException ex)
            {
                //Console.WriteLine("The file could not be read:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
