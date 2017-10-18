using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;
using Newtonsoft.Json;
using System.Data.SqlClient;

namespace ProjectNibbles {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            InitStuff();
        }

        OrderedDictionary currentCar = new OrderedDictionary();
        string[] attributes = { "vrn", "vehicle_make", "vehicle_model", "vehicle_registration_year", "seller_type", "vehicle_mileage", "vehicle_colour", "price", "vehicle_not_writeoff", "vehicle_vhc_checked", "url", "location", "mot_expiry" };
        
        public void InitStuff() {
            //Initialise the FBD and the ordered dictionary with empty strings because a Car class would suck
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = "HTML file browser";
            openFileDialog1.Filter = "HTML (*.html;*htm)|*.html;*htm|" + "All files (*.*)|*.*";
 
            foreach (string attr in attributes)
                currentCar.Add(attr, "");
        }

        private void button_add_to_list_Click(object sender, EventArgs e) {
            AddToListView();
        }

        public void AddToListView() {
            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                foreach (string fileName in openFileDialog1.FileNames) {
                    string fullHTML = File.ReadAllText(fileName);
                    string dataLayer = ExtractRegex(fullHTML, "var dataLayer = ", "}];") + "}]";
                    //tempBox.AppendText(dataLayer);
                    dynamic obj = JsonConvert.DeserializeObject(dataLayer);

                    //Extracting the JSON data - if only I could loop this...
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

                    //Extracting the non-JSON data
                    currentCar[attributes[10]] = ExtractRegex(fullHTML, "https://", " -->"); //Saved Chrome files have the URL in the second line
                    currentCar[attributes[11]] = ExtractLocation(fullHTML).Replace(",", "");

                    //Add VRN to column 0 inside the declaration/init and add the rest of the attributes to ListView preparation
                    ListViewItem car = new ListViewItem(currentCar[attributes[0]].ToString());
                    for (int i = 1; i < currentCar.Count; i++) 
                        car.SubItems.Add(currentCar[attributes[i]].ToString());
                    
                    //Add to actual ListView
                    listView1.Items.AddRange(new ListViewItem[] { car });
                }
            }
        }

        SqlConnection connect = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename=D:\Projects\Visual Studio 2017\ProjectNibbles\ProjectNibbles\MyCars.mdf;Integrated Security = True");

        private void button_commit_to_db_Click(object sender, EventArgs e) {
            CommitToDB();
        }

        public bool VerifyRecord(string vrn) {           
            connect.Open();
            SqlCommand command = connect.CreateCommand();
            command.CommandText = "SELECT COUNT(*) FROM MyCars WHERE vrn = @vrn";
            command.Parameters.AddWithValue("@vrn", vrn);
            bool x = ((int)command.ExecuteScalar() <= 0) ? false : true; //It may return a capitalised True/False
            connect.Close();
            return x;
        }

        public void CommitToDB() {
            foreach (ListViewItem row in listView1.Items) {
                if (VerifyRecord(row.SubItems[0].Text.ToString()) == false) {
                    connect.Open();
                    SqlCommand command = connect.CreateCommand();
                    //Concatenators are bad apparently
                    command.CommandText = "INSERT INTO MyCars (" + string.Join(", ", attributes) + ") VALUES (@" + string.Join(", @", attributes) + ")";

                    //Implicit conversions from ListViewItem.ToString to the value types in the DB
                    for (int i = 0; i < attributes.Length; i++) 
                        command.Parameters.AddWithValue("@" + attributes[i], row.SubItems[i].Text.ToString());
                    
                    command.ExecuteNonQuery();
                    connect.Close();
                }
            }
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

        //Temp button
        private void button1_Click(object sender, EventArgs e) {
            //tempBox.AppendText(VerifyRecord("TEST002").ToString());
            //tempBox.AppendText("INSERT INTO MyCars (" + string.Join(", ", attributes) + ")");
            string[] temp = attributes.Select(x => "@" + x).ToArray();
            //foreach (string x in temp) tempBox.AppendText(x + "\n");
            tempBox.AppendText(string.Join(", ", attributes.Select(x => "@" + x).ToArray())); //Hilarious
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e) {
            if (Keys.Delete == e.KeyCode) {
                foreach (ListViewItem item in listView1.SelectedItems)
                    listView1.Items.Remove(item);
            }
        }
    }
}
