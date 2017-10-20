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

        //Could replace the OD with class if I can figure out how to loop through a class's properties
        OrderedDictionary currentCar = new OrderedDictionary();
        //using the same fields in the DB as the JSON will make things easier to loop later
        string[] attributes = { "vrn", "vehicle_make", "vehicle_model", "vehicle_registration_year", "seller_type", "vehicle_mileage", "vehicle_colour", "price", "vehicle_not_writeoff", "vehicle_vhc_checked", "url", "location", "mot_expiry" };
        
        //Init the SQL connection with connection string (single DB)
        SqlConnection connect = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename=D:\Projects\Visual Studio 2017\ProjectNibbles\ProjectNibbles\MyCars.mdf;Integrated Security = True");

        public void InitStuff() {
            //Initialise the FBD and the ordered dictionary with empty strings to determine order.
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = "HTML file browser";
            openFileDialog1.Filter = "HTML (*.html;*htm)|*.html;*htm|" + "All files (*.*)|*.*";

            saveFileDialog1.Filter = "CSV (*.csv)|*csv";
            saveFileDialog1.DefaultExt = "csv";
            saveFileDialog1.AddExtension = true;
 
            foreach (string attr in attributes)
                currentCar.Add(attr, "");
        }

        //Preview cars in ListView before commiting to DB
        private void button_add_to_list_Click(object sender, EventArgs e) {
            AddToListView();
        }

        public void AddToListView() {
            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                foreach (string fileName in openFileDialog1.FileNames) {
                    string fullHTML = File.ReadAllText(fileName);
                    string dataLayer = ExtractRegex(fullHTML, "var dataLayer = ", "}];") + "}]"; //Still a pain to extract
                    //tempBox.AppendText(dataLayer);
                    dynamic obj = JsonConvert.DeserializeObject(dataLayer);

                    //Extracting the JSON data strings and assigning to currentCar OD - if only I could loop this...
                    string reg = obj[0].a.attr.vrn;
                    reg = reg.Replace(" ", ""); //Ensure no white space

                    string price = obj[0].a.attr.price;
                    price = price.Substring(0, price.Length - 2); //Chop off the last two digits of the string price

                    currentCar[attributes[0]] = reg;
                    currentCar[attributes[1]] = obj[0].a.attr.vehicle_make;
                    currentCar[attributes[2]] = obj[0].a.attr.vehicle_model;
                    currentCar[attributes[3]] = obj[0].a.attr.vehicle_registration_year;
                    currentCar[attributes[4]] = obj[0].a.attr.seller_type;
                    currentCar[attributes[5]] = obj[0].a.attr.vehicle_mileage;
                    currentCar[attributes[6]] = obj[0].a.attr.vehicle_colour;
                    currentCar[attributes[7]] = price;
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

        //Commit contents of ListView to DB, Method names are self-explanatory
        private void button_commit_to_db_Click(object sender, EventArgs e) {
            CommitToDB();
        }

        public bool RecordExists(string vrn) {           
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
                if (RecordExists(row.SubItems[0].Text.ToString()) == false) {
                    connect.Open();
                    SqlCommand command = connect.CreateCommand();
                    //Concatenators are bad apparently. Using string.Join to join the elements of the attributes array, first is with comma, second is with @
                    command.CommandText = "INSERT INTO MyCars (" + string.Join(", ", attributes) + ") VALUES (@" + string.Join(", @", attributes) + ")";

                    for (int i = 0; i < attributes.Length; i++) 
                        command.Parameters.AddWithValue("@" + attributes[i], row.SubItems[i].Text.ToString());

                    //Implicit conversions from ListViewItem.ToString to the value types in the DB
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

        //Delete part - easy enough, extract reg number from selected ListViewItem
        private void button_delete_Click(object sender, EventArgs e) {
            string selectedVRN = listView2.SelectedItems[0].Text;
            DeleteRecord(selectedVRN);
        }

        //Delete via VRN. Not going to delete by a criteria just yet, e.g. delete all cars with price >5000.
        public void DeleteRecord(string vrn) {
            connect.Open();
            SqlCommand command = connect.CreateCommand();
            command.CommandText = "DELETE FROM MyCars WHERE vrn = @vrn";
            command.Parameters.AddWithValue("@vrn", vrn);
            command.ExecuteNonQuery();
            connect.Close();
        }

        private void button_display_all_Click(object sender, EventArgs e) {
            DisplayRecords("SELECT * FROM MyCars");
        }

        //Display All
        public void DisplayRecords(string commandText) {
            connect.Open();
            SqlCommand command = connect.CreateCommand();
            command.CommandText = commandText;
            SqlDataReader reader = command.ExecuteReader();

            listView2.Items.Clear();

            while (reader.Read()) {
                ListViewItem car = new ListViewItem(reader[attributes[0]].ToString()); //VRN first in column 0
                for (int i = 1; i < attributes.Length; i++) {
                    car.SubItems.Add(reader[attributes[i]].ToString());
                }
                listView2.Items.Add(car); //Alternative the method done in AddToListView();
            }
            connect.Close();
        }

        //To use Selenium or not, that is the question
        private void button_update_record_Click(object sender, EventArgs e) {
            DateTime dt = dateTimePicker1.Value;
            string selectedVRN = listView2.SelectedItems[0].Text;
            UpdateMOT(selectedVRN,dt);
        }

        //Update MOT 
        public void UpdateMOT(string vrn, DateTime dt) {
            connect.Open();
            SqlCommand command = connect.CreateCommand();
            command.CommandText = "UPDATE MyCars SET mot_expiry = @dt WHERE vrn = @vrn";
            command.Parameters.AddWithValue("@dt", dt);
            command.Parameters.AddWithValue("@vrn", vrn);
            command.ExecuteNonQuery();
            connect.Close();
        }

        //Temp button
        private void button1_Click(object sender, EventArgs e) {
            //tempBox.AppendText(VerifyRecord("TEST002").ToString());
            //tempBox.AppendText("INSERT INTO MyCars (" + string.Join(", ", attributes) + ")");
            //string[] temp = attributes.Select(x => "@" + x).ToArray();
            //foreach (string x in temp) tempBox.AppendText(x + "\n");
            //tempBox.AppendText(string.Join(", ", attributes.Select(x => "@" + x).ToArray())); //Hilarious
            //DeleteRecord("LS07JWP"); //Goodnight Passat - absolutely nothing happens if you try to delete a non-existent record. Could do the verify though.

            //Get selected item in ListView, output that to box
            string text = listView1.SelectedItems[0].Text; //Returns column 0, i.e. the VRN
            tempBox.AppendText(text);
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e) {
            if (Keys.Delete == e.KeyCode) {
                foreach (ListViewItem item in listView1.SelectedItems)
                    listView1.Items.Remove(item);
            }
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e) {
            Application.Exit();
        }

        private void button_export_csv_Click(object sender, EventArgs e) {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                string saveFile = saveFileDialog1.FileName;
                if (!File.Exists(saveFile)) {
                    using (StreamWriter writer = File.CreateText(saveFile)) {
                        foreach (string a in attributes) {
                            writer.Write(a + ",");
                        }
                        writer.Write("\n");
                    }
                }

                using (StreamWriter writer = File.AppendText(saveFile)) {
                    foreach (ListViewItem itemRow in listView2.Items) {
                        for (int i = 0; i < itemRow.SubItems.Count; i++) {
                            //tempBox.AppendText(itemRow.SubItems[i].Text + "\n");
                            writer.Write(itemRow.SubItems[i].Text + ",");
                        }
                        writer.Write("\n");
                    }
                }
            }



        }
    }

    // For reference
    public class Car {
        public string vrn, make, model;
        public int year;

        public Car(string _vrn, string _make, string _model, int _year) {
            this.vrn = _vrn;
            this.make = _make;
            this.model = _model;
            this.year = _year;
        }

        public bool CarExistsInDB() {
            SqlConnection connect = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename=D:\Projects\Visual Studio 2017\ProjectNibbles\ProjectNibbles\MyCars.mdf;Integrated Security = True");
            connect.Open();
            SqlCommand command = connect.CreateCommand();
            command.CommandText = "SELECT COUNT(*) FROM MyCars WHERE vrn = @vrn";
            command.Parameters.AddWithValue("@vrn", vrn);
            bool x = ((int)command.ExecuteScalar() <= 0) ? false : true;
            connect.Close();
            return x;
        }

        public int CarAge() {
            return int.Parse(DateTime.Now.Year.ToString()) - year;
        }

        public string VRN { get; set; }
        public string Make { get; set; }
        public string Model { get; set; }
        public int Year { get; set; }
    }

}
