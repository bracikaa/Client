using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace OpenPop.TestApplication
{
    public partial class UserAccountsForm : Form
    {
        List<string> _items = new List<string>();
        String connectionString = String.Empty;
        Int32 selectedUserId = -1;

        public UserAccountsForm(String connectionString)
        {
            InitializeComponent();

            passwordTxt.PasswordChar = '*';

            this.connectionString = connectionString;

            var connection = new SqlConnection(connectionString);

            //reader
            try
            {
                //open connection
                connection.Open();

                //command1
                String sqlGetUsers = "select Username from Users";
                SqlCommand command = new SqlCommand(sqlGetUsers, connection);
                using (SqlDataReader rdr = command.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        if (!rdr.IsDBNull(0))
                            _items.Add(rdr.GetString(0));
                    }
                }
                command.Dispose();

                //close connection
                connection.Close();
            }
            catch (Exception ex)
            {
            }

            listBox1.DataSource = _items;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            if(_items.Count > 0)
                 SetUser();
            else
                MessageBox.Show("No users in the db.");
        }

        private void listBox1_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            SetUser();
        }

        private void SetUser()
        {
            String curItem = listBox1.SelectedItem.ToString();

            var connection = new SqlConnection(this.connectionString);

            //reader
            try
            {
                //open connection
                connection.Open();

                //command1
                String sqlGetUserData = "select Id, Username, Password, UserGroup, Enabled from Users where Username = '" + curItem + "'";
                SqlCommand command = new SqlCommand(sqlGetUserData, connection);
                using (SqlDataReader rdr = command.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        if (!rdr.IsDBNull(0))
                            this.selectedUserId = rdr.GetInt32(0);
                        if (!rdr.IsDBNull(1))
                            usernameTxt.Text = rdr.GetString(1);
                        if (!rdr.IsDBNull(2))
                            passwordTxt.Text = rdr.GetString(2);
                        if (!rdr.IsDBNull(3))
                            groupTxt.Text = rdr.GetString(3);
                        if (!rdr.IsDBNull(4))
                            enabledCb.Checked = rdr.GetBoolean(4);
                    }
                }
                command.Dispose();

                //close connection
                connection.Close();
            }
            catch (Exception ex)
            {
            }
        }

        private void updateUserBtn_Click(object sender, EventArgs e)
        {
            var connection = new SqlConnection(this.connectionString);

            //reader
            try
            {
                //open connection
                connection.Open();

                int cbChecked = 0;

                if (enabledCb.Checked)
                    cbChecked = 1;

                //command1
                String sqlGetUserData = "update Users Set Username = '" + usernameTxt.Text + "', Password = '" + passwordTxt.Text + "', UserGroup = '" + groupTxt.Text + "', Enabled = " + cbChecked + " where Id = " + this.selectedUserId + "";
                SqlCommand command = new SqlCommand(sqlGetUserData, connection);
                command.ExecuteNonQuery();
                command.Dispose();
                MessageBox.Show(this, "User account successfully updated");

                //close connection
                connection.Close();
            }
            catch (Exception ex)
            {
            }
        }
    }
}
