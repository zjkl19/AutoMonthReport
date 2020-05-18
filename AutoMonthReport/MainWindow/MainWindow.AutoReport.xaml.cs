using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace AutoMonthReport
{
    public partial class MainWindow : Window
    {
        //免责声明
        private void AutoReport_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show("测试");

            string templateFile = @"Template\月报模板.docx";
            string outputFile = @"OutputReport\自动生成的月报.docx";

            //TODO：加入验证
            DateTime startDate = StartDate.SelectedDate ?? DateTime.Now;
            DateTime endDate = EndDate.SelectedDate ?? DateTime.Now;

            try
            {
                var doc = new Document(templateFile);
                doc.Variables["StartDate"] = startDate.ToString("D");
                doc.Variables["EndDate"] = endDate.ToString("D");
                doc.UpdateFields();
                doc.Save(outputFile, SaveFormat.Docx);
                MessageBox.Show("成功生成月报！");
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
    }
}
