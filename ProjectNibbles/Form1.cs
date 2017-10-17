using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using Newtonsoft.Json;

namespace ProjectNibbles {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            InitStuff();
        }

        OrderedDictionary currentCar = new OrderedDictionary();
        string[] attributes = { "vrn", "vehicle_make", "vehicle_model", "vehicle_registration_year", "seller_type", "vehicle_mileage", "vehicle_colour", "price", "vehicle_not_writeoff", "vehicle_vhc_checked", "URL", "Location", "MOT expiry" };

        public void InitStuff() {
            //Initialise the FBD and the ordered dictionary because a Car class would suck
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = "HTML file browser";
            openFileDialog1.Filter = "HTML (*.html;*htm)|*.html;*htm|" + "All files (*.*)|*.*";
            
            foreach (string attr in attributes)
                currentCar.Add(attr, "");
        }



        private void button_add_to_list_Click(object sender, EventArgs e) {
            M1();
        }

        public void M1() {
            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                foreach (string fileName in openFileDialog1.FileNames) {
                    string fullHTML = File.ReadAllText(fileName);
                    string dataLayer = ExtractRegex(fullHTML, "var dataLayer = ", "}];") + "}]";
                    tempBox.AppendText(dataLayer);
                    dynamic obj = JsonConvert.DeserializeObject(dataLayer);

                    currentCar[attributes[0]] = obj[0].a.attr.vrn;
                    currentCar[attributes[1]] = obj[0].a.attr.vehicle_make;
                    currentCar[attributes[2]] = obj[0].a.attr.vehicle_model;
                    currentCar[attributes[3]] = obj[0].a.attr.vehicle_registration_year;
                    currentCar[attributes[4]] = obj[0].a.attr.seller_type;
                    currentCar[attributes[5]] = obj[0].a.attr.vehicle_mileage;
                    currentCar[attributes[6]] = obj[0].a.attr.vehicle_colour;
                    currentCar[attributes[7]] = obj[0].a.attr.price;
                    currentCar[attributes[8]] = obj[0].a.attr.vehicle_not_writeoff;
                    currentCar[attributes[9]] = obj[0].a.attr.vehicle_vhc_checked;

                    currentCar[attributes[10]] = ExtractRegex(fullHTML, "https://", " -->"); //Saved Chrome files have the URL in the second line
                    currentCar[attributes[11]] = ExtractLocation(fullHTML).Replace(",", "");

                    ListViewItem car = new ListViewItem(currentCar[attributes[0]].ToString());
                    for (int i = 1; i < currentCar.Count; i++) {
                        car.SubItems.Add(currentCar[attributes[i]].ToString());
                    }

                    listView1.Items.AddRange(new ListViewItem[] { car });

                }
            }
        }

        private void button_commit_to_db_Click(object sender, EventArgs e) {

        }

        public string ExtractRegex(string source, string start, string end) {
            Regex rx = new Regex(start + "(.*?)" + end);
            var match = rx.Match(source);
            return match.Groups[1].ToString();
        }

        public string ExtractLocation(string source) {
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(source);
            var singleNode = doc.DocumentNode.SelectSingleNode("//body/div/div/div/main/div/header/strong/span");
            return singleNode.InnerText.ToString();
        }

    }
}
