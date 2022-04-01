using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

/*
 * Формат xml файлов алкодекларации
 * https://fsrar.gov.ru/files/structure/2248.pdf
 * https://wiki.sftserv.ru/index.php/%D0%A4%D0%B0%D0%B9%D0%BB%D1%8B_%D0%94%D0%B5%D0%BA%D0%BB%D0%B0%D1%80%D0%B0%D0%BD%D1%82-%D0%90%D0%BB%D0%BA%D0%BE
 * 
 * Проверить алкодекларацию - https://alko.kontur.ru/
 * 
 */


namespace AlcoDec
{
    public partial class Form1 : Form
    {
        XDocument xdoc;
        string xmlFile = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string xmlFile = ".\\" + "R2_472004353670_090_15102020_A2868624-3D48-4F96-A4E7-35F7C0BC6527" + ".xml";
            xdoc = XDocument.Load(xmlFile);
            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Elements("Справочники").Elements("ПроизводителиИмпортеры");
            //MessageBox.Show(xdoc.Element("Файл").Elements("Справочники").Elements("ПроизводителиИмпортеры").Count().ToString());
            string attrName_ID = "ИДПроизвИмп";
            string attrName_Name = "П000000000004";
            string attrName_INN = "П000000000005";

            GetElements(elementFind, attrName_ID, attrName_Name, attrName_INN);   // Ищем производителей
            attrName_ID = "ИдПостав";
            attrName_Name = "П000000000007";
            attrName_INN = "П000000000009";
            elementFind = xdoc.Element("Файл").Elements("Справочники").Elements("Поставщики");  // Ищем поставщиков
            GetElements(elementFind, attrName_ID, attrName_Name, attrName_INN);
        }

        private void GetElements(IEnumerable<XElement> _elementFind, string _attrName_ID, string _attrName_Name, string _attrName_INN)
        {
            foreach (XElement xElem in _elementFind)
            {
                XAttribute attProviderId = xElem.Attribute(_attrName_ID);
                XAttribute attProviderName = xElem.Attribute(_attrName_Name);
                XAttribute attProviderINN = null;
                string attProviderINNstr = string.Empty;
                foreach (XElement xElemUL in xElem.Elements("ЮЛ"))  // Заходим глубще (в тег "ЮЛ") и ищем ИНН
                {
                    //MessageBox.Show(xxe.ToString());
                    attProviderINN = xElemUL.Attribute(_attrName_INN);

                }
                if (attProviderINN == null)
                {
                    attProviderINNstr = "null";
                }
                else
                {
                    attProviderINNstr = attProviderINN.Value;
                }
                richTextBox1.AppendText(attProviderId.Value + " " + attProviderName.Value + " " + attProviderINNstr + "\n");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox2.Clear();
            dataGridView1.Rows.Clear();
            label5.Visible = false;
            label6.Visible = false;
            label7.Visible = false;
            label8.Visible = false;
            OpenFileDialog ofd = new OpenFileDialog();            
            ofd.Filter = "xml files (*.xml)|*.xml";            
            if (ofd.ShowDialog() == DialogResult.Cancel)
                return;
            xmlFile = ofd.FileName;
            label5.Text = "...загружен";
            label5.Visible = true;
            xdoc = XDocument.Load(xmlFile);
            OrgRequisites();   // Проверяем заполнены ли реквизиты организации
            panel1.Enabled = true;
            Importer();         // Перебираем импортные сорта
            if (dataGridView1.Rows.Count == 0)
            {
                button3.Enabled = false;
                button4.Enabled = true;
            }
            else
            {
                button4.Enabled = false;
                button3.Enabled = true;
            }
        }

        private bool OrgRequisites()
        {
            bool ret = true;
            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Element("Документ").Elements("Организация").Elements("Реквизиты");
            foreach (XElement xElem in elementFind)
            {
                XAttribute orgName = xElem.Attribute("Наим");
                XAttribute orgPhone = xElem.Attribute("ТелОрг");
                XAttribute orgMail = xElem.Attribute("EmailОтпр");
                textBox1.Text = orgName.Value;
                textBox2.Text = orgPhone.Value;
                textBox3.Text = orgMail.Value;
                if (orgName.Value == "")
                    MessageBox.Show("Не заполненны реквизит:\n" + "Название организации"); ret = false;
                if (orgPhone.Value == "")
                    MessageBox.Show("Не заполненны реквизит:\n" + "Телефон организации"); ret = false;
                if (orgMail.Value == "")
                    MessageBox.Show("Не заполненны реквизит:\n" + "Эл. почта организации"); ret = false;
            }
            return ret;
        }
        private void Importer()     // Перебираем импортные сорта
        {
            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Elements("Справочники").Elements("ПроизводителиИмпортеры").Elements("ФЛ");
            foreach (XElement xElem in elementFind)
            {
                XAttribute attManufacturerName = xElem.Parent.Attribute("П000000000004");
                XAttribute attManufacturerID = xElem.Parent.Attribute("ИДПроизвИмп");
                /*xElem.SetAttributeValue("П000000000005", "888");    // Добавить атрибут
                xElem.SetAttributeValue("П000000000006", "999");    // Добавить атрибут
                xElem.Name = "ЮЛ";  // Переименовать элемент*/
                //MessageBox.Show("Обнаружен импортный производитель.\n" + "Название: " + attManufacturerName.Value + "\nID: " + attManufacturerID.Value);
                dataGridView1.Rows.Add(attManufacturerID.Value, attManufacturerName.Value);
                ColumnNameImporter.DataSource = ImportersList();    // Загружаем список поставщиков в combobox'ы
                //ImportersList(xdoc)
            }
        }

        private List<String> ImportersList()  // Список поставщиков
        {
            List<String> ret = new List<string>();
            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Elements("Справочники").Elements("Поставщики").Elements("ЮЛ");
            foreach (XElement xElem in elementFind)
            {
                XAttribute attManufacturerName = xElem.Parent.Attribute("П000000000007");
                //XAttribute attManufacturerID = xElem.Parent.Attribute("ИДПроизвИмп");
                XAttribute attManufacturerINN = xElem.Attribute("П000000000009");
                XAttribute attManufacturerKPP = xElem.Attribute("П000000000010");
                ret.Add(attManufacturerName.Value + "(" + attManufacturerINN.Value + "/" + attManufacturerKPP.Value + ")");
            }
            return ret;
        }

        private void Turnover()     // Заполнение оборота
        {
            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Elements("Документ").Elements("ОбъемОборота").Elements("Оборот").Elements("СведПроизвИмпорт").Elements("Движение");
            foreach (XElement xElem in elementFind)
            {
                XAttribute attAllDecal = xElem.Attribute("П100000000008");  // Оборот в декалитрах (от организаций оптовой торговли)
                if (attAllDecal.Value == "0.00000" && xElem.Attribute("П100000000007").Value != "0.00000")  // Если оборот от производителя (от оптовиков нет, а от производителей есть)
                {
                    //MessageBox.Show("Производитель!");
                    attAllDecal = xElem.Attribute("П100000000007");
                }
                richTextBox2.AppendText(xElem.ToString() + "\n");
                xElem.Attribute("П100000000014").Value = attAllDecal.Value; // Расход
                xElem.Attribute("П100000000017").Value = attAllDecal.Value; // Продано
                xElem.Attribute("П100000000018").Value = "0.00000";         // Остаток
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string s = "";
            string inn = "";
            string kpp = "";
            int errors = 0;


            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Elements("Справочники").Elements("ПроизводителиИмпортеры").Elements("ФЛ");
            for (int i=0; i < dataGridView1.Rows.Count; i++)
            {
                foreach (XElement xElem in elementFind)
                {
                    XAttribute attManufacturerID = xElem.Parent.Attribute("ИДПроизвИмп");
                    string dataGrinRowID = dataGridView1.Rows[i].Cells["ColumnID"].Value.ToString();
                    if (attManufacturerID.Value == dataGrinRowID)
                    {
                        //MessageBox.Show("Совпадение " + attManufacturerID.Value);
                        if (dataGridView1.Rows[i].Cells["ColumnNameImporter"].Value == null)
                        {
                            MessageBox.Show("Не выбраны Поставщики - Импортёры!\n" + "Заполните все значения.");
                            errors++;
                            //break;
                            return;
                        }
                        s = dataGridView1.Rows[i].Cells["ColumnNameImporter"].Value.ToString();
                        inn = Regex.Match(s, @"(?<=\().+?(?=\/)").ToString();
                        kpp = Regex.Match(s, @"(?<=\/).+?(?=\))").ToString();
                        xElem.SetAttributeValue("П000000000005", inn);    // Добавить атрибут
                        xElem.SetAttributeValue("П000000000006", kpp);    // Добавить атрибут
                        xElem.Name = "ЮЛ";  // Переименовать элемент*/
                    }
                }
            }
            if (errors == 0)
            {
                richTextBox1.Text = xdoc.ToString();
                button4.Enabled = true;
                label7.Text = "...готово";
                label7.Visible = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            string xmlFileNameOnly = Path.GetFileNameWithoutExtension(xmlFile);
            sfd.FileName = xmlFileNameOnly;
            sfd.Filter = "xml files (*.xml)|*.xml";
            Turnover();   // Заполнение оборота
            label8.Text = "...готово";
            label8.Visible = true;
            MessageBox.Show("Готово!\nВыберите место сохранения файла\nНе меняйте имя файла:\n" + xmlFileNameOnly);
            if (sfd.ShowDialog() == DialogResult.Cancel)
                return;
            xdoc.Save(sfd.FileName);
            string testSite = @"https://alko.kontur.ru/";
            DialogResult result = MessageBox.Show(
                "Файл сохранён!\n" + "Открыть сайт " + testSite + " для проверки декларации?",
                "Информация",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            if (result == DialogResult.Yes)
                Process.Start(testSite);
            else
            {
                this.TopMost = true;
                this.TopMost = false;
            }

        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            IEnumerable<XElement> elementFind = xdoc.Element("Файл").Element("Документ").Elements("Организация").Elements("Реквизиты");
            foreach (XElement xElem in elementFind)
            {
                xElem.SetAttributeValue("Наим", textBox1.Text);    // Добавить атрибут
                xElem.SetAttributeValue("ТелОрг", textBox2.Text);    // Добавить атрибут
                xElem.SetAttributeValue("EmailОтпр", textBox3.Text);    // Добавить атрибут
            }
            label6.Text = "...сохранено";
            label6.Visible = true;
        }
    }
}
