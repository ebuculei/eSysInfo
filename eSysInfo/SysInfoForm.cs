using System;
using System.Collections;
using System.Management;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace eSysInfo
{
    public partial class SysInfoForm : Form
    {
        private string fnSys;
        private string fnHard;
        private string skSys;
        private string skHard;
        public SysInfoForm()
        {
            InitializeComponent();
            treeViewSystem.SelectedNode = treeViewSystem.Nodes.Find("Node89", true)[0];
        }

        private void InsertInfo(string SearchKey, ref ListView lstView, bool AddNullValue)
        {
            lstView.Items.Clear();
            bool pFlgDataExist = false;

            ManagementObjectSearcher Searcher = new ManagementObjectSearcher("SELECT * FROM " + SearchKey);

            try
            {
                foreach (ManagementObject mngObj in Searcher.Get())
                {
                    pFlgDataExist = true;
                    ListViewGroup Group;
                    try
                    {
                        Group = lstView.Groups.Add(mngObj["Name"].ToString(), mngObj["Name"].ToString());
                    }
                    catch
                    {
                        Group = lstView.Groups.Add(mngObj.ToString(), mngObj.ToString());
                    }

                    if (mngObj.Properties.Count <= 0)
                    {
                        //Interaction.MsgBox("No Information Currently Available", MsgBoxStyle.Information, "No Info to Display");
                        MessageBox.Show("No Information Available", "No Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (PropertyData PD in mngObj.Properties)
                    {
                        ListViewItem item = new ListViewItem(Group);
                        if (lstView.Items.Count % 2 != 0)
                            item.BackColor = Color.White;
                        else
                            item.BackColor = Color.WhiteSmoke;

                        item.Text = PD.Name;

                        if (PD.Value != null && PD.Value.ToString() != "")
                        {
                            switch (PD.Value.GetType().ToString())
                            {
                                case "System.String[]":
                                    {
                                        string[] StrMain = (string[])PD.Value;

                                        string strDisplay = "";
                                        foreach (string st in StrMain)
                                            strDisplay += st + " ";

                                        item.SubItems.Add(strDisplay);
                                        break;
                                    }

                                case "System.UInt16[]":
                                    {
                                        ushort[] ShortMain = (ushort[])PD.Value;

                                        string strDisplay = "";
                                        foreach (ushort st in ShortMain)
                                            strDisplay += st.ToString() + " ";

                                        item.SubItems.Add(strDisplay);
                                        break;
                                    }

                                default:
                                    {
                                        item.SubItems.Add(PD.Value.ToString());
                                        break;
                                    }
                            }
                        }
                        else if (AddNullValue == true)
                            //item.SubItems.Add("No Information Available to Display");
                            item.SubItems.Add("");
                        else
                            continue;
                        lstView.Items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                //Interaction.MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error on Getting Information");
                MessageBox.Show("Error on Getting Information", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            if (pFlgDataExist == false)
            {
                ListViewGroup Group;
                Group = lstView.Groups.Add("No Information Available to Display", "No Information Available to Display");
                ListViewItem item = new ListViewItem(Group);
                lstView.Items.Add(item);
            }
        }

        private void RemoveNullValue(ref ListView lstView)
        {
            foreach (ListViewItem item in lstView.Items)
            {
                //if (item.SubItems[1].Text == "No Information Available to Display")
                if (item.SubItems[1].Text == "")
                    item.Remove();
            }
        }

        private void SizeLastColumn(ListView lv)
        {
            lv.Columns[lv.Columns.Count - 1].Width = -2;
        }

        private void listViewSystem_Resize(object sender, EventArgs e)
        {
            SizeLastColumn((ListView)sender);
        }

        private void listViewHardware_Resize(object sender, EventArgs e)
        {
            SizeLastColumn((ListView)sender);
        }

        private void treeViewSystem_AfterSelect(object sender, TreeViewEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            lblStatus.Text = "Operation in progress, please wait...";
            toolStripProgressBar1.Visible = true;
            {
                var withBlock = treeViewSystem;
                if (!(e.Node.Parent == null))
                {
                    toolStripLabelSys.Text = withBlock.SelectedNode.Text;
                    this.Refresh();
                    // Child Node
                    fnSys = withBlock.SelectedNode.Text.Replace(" ", "");
                    skSys = "Win32_" + withBlock.SelectedNode.Text.Replace(" ", "");
                    InsertInfo(skSys, ref listViewSystem, chkShowNullValueSystem.Checked);
                }
            }
            Cursor = Cursors.Default;
            lblStatus.Text = "Ready";
            toolStripProgressBar1.Visible = false;
        }

        private void treeViewHardware_AfterSelect(object sender, TreeViewEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            lblStatus.Text = "Operation in progress, please wait...";
            toolStripProgressBar1.Visible = true;
            {
                var withBlock = treeViewHardware;
                if (!(e.Node.Parent == null))
                {
                    toolStripLabelHard.Text = withBlock.SelectedNode.Text;
                    //this.Refresh();
                    // Child Node
                    fnHard = withBlock.SelectedNode.Text.Replace(" ", "");
                    skHard = "Win32_" + withBlock.SelectedNode.Text.Replace(" ", "");
                    InsertInfo(skHard, ref listViewHardware, chkShowNullValueHardware.Checked);
                }
            }
            Cursor = Cursors.Default;
            lblStatus.Text = "Ready";
            toolStripProgressBar1.Visible = false;
        }

        private void chkShowNullValueSystem_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            lblStatus.Text = "Operation in progress, please wait...";
            toolStripProgressBar1.Visible = true;
            if (listViewSystem.Items.Count > 1)
            {
                if (chkShowNullValueSystem.Checked == false)
                    RemoveNullValue(ref listViewSystem);
                else
                    InsertInfo(skSys, ref listViewSystem, chkShowNullValueSystem.Checked);
            }
            Cursor = Cursors.Default;
            lblStatus.Text = "Ready";
            toolStripProgressBar1.Visible = false;
        }

        private void chkShowNullValueHardware_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            lblStatus.Text = "Operation in progress, please wait...";
            toolStripProgressBar1.Visible = true;
            if (listViewHardware.Items.Count > 1)
            {
                if (chkShowNullValueHardware.Checked == false)
                    RemoveNullValue(ref listViewHardware);
                else
                    InsertInfo(skHard, ref listViewHardware, chkShowNullValueHardware.Checked);
            }
            Cursor = Cursors.Default;
            lblStatus.Text = "Ready";
            toolStripProgressBar1.Visible = false;
        }

        SaveFileDialog saveFileDialog = new SaveFileDialog
        {
            CheckFileExists = false,
            DefaultExt = "txt",
            Filter = "Text File|*.txt",
            SupportMultiDottedExtensions = true
        };

        private void btnSavetoTxtSystem_Click(object sender, EventArgs e)
        {
            if (this.listViewSystem.Items.Count > 1)
            {
                saveFileDialog.FileName = fnSys;
                //saveFileDialog.FileName = skSys;
                //saveFileDialog.Filter = "Text File|*.txt";

                DialogResult dr = saveFileDialog.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    string txtFileName = saveFileDialog.FileName;   // Path   
                    if (!string.IsNullOrEmpty(txtFileName))
                    {
                        using (StreamWriter myWriter = new StreamWriter(txtFileName))
                        {
                            foreach (ListViewItem myItem in listViewSystem.Items)
                                myWriter.WriteLine(myItem.Text + "   :   " + myItem.SubItems[1].Text);
                            myWriter.Close();
                            MessageBox.Show("Saved Data to " + txtFileName, "Saved Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

            }
        }

        private void btnSavetoTxtHardware_Click(object sender, EventArgs e)
        {
            if (this.listViewHardware.Items.Count > 1)
            {
                saveFileDialog.FileName = fnHard;
                //saveFileDialog.FileName = skHard;
                //saveFileDialog.Filter = "Text File|*.txt";

                DialogResult dr = saveFileDialog.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    string txtFileName = saveFileDialog.FileName;   // Path   
                    if (!string.IsNullOrEmpty(txtFileName))
                    {
                        using (StreamWriter myWriter = new StreamWriter(txtFileName))
                        {
                            foreach (ListViewItem myItem in listViewHardware.Items)
                                myWriter.WriteLine(myItem.Text + "   :   " + myItem.SubItems[1].Text);
                            myWriter.Close();
                            MessageBox.Show("Saved Data to " + txtFileName, "Saved Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

            }
        }

        private static void ShowAboutBox()
        {
            using (AboutBox ab = new AboutBox())
            {
                ab.ShowDialog();
            }
        }
        private void btnAboutSys_Click(object sender, EventArgs e)
        {
            ShowAboutBox();
        }

        private void btnAboutHard_Click(object sender, EventArgs e)
        {
            ShowAboutBox();
        }
    }
}
