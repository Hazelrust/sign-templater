using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using sign_templater.Services;
using Word = Microsoft.Office.Interop.Word;

namespace sign_templater
{
    public partial class SidebarControl : UserControl
    {
        private ComboBox comboQuickDest;
        private TextBox txtAddress;
        private TextBox txtAdmissionNo;
        private TextBox txtExtraID;
        private Label lblExtraID;
        private ComboBox comboDegree;
        private CheckedListBox listDocuments;
        private Button btnGenerate;

        public SidebarControl()
        {
            InitializeComponent();
            SetupCustomUI();
        }

        private void SetupCustomUI()
        {
            this.AutoScroll = true;
            int currentY = 20;
            int padding = 20;
            int controlWidth = this.Width - (padding * 2);

            // 0. Degree Level
            Label lblDegree = new Label { Text = "ระดับการศึกษา:", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth };
            currentY += 25;
            comboDegree = new ComboBox { Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            comboDegree.Items.AddRange(new[] { "(None)", "Bachelor Degree", "Master Degree", "Graduate Diploma" });
            currentY += 40;

            // 1. Quick Select
            Label lblQuick = new Label { Text = "เลือกที่อยู่ด่วน (Destinations):", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth };
            currentY += 25;
            comboQuickDest = new ComboBox { Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            comboQuickDest.Items.AddRange(new[] { 
                "Manual Input (พิมพ์เอง)",
                "LSAC (Credential Assembly Service)",
                "New York State BOLE",
                "The Evaluation Company (Texas)",
                "Harvard Law School",
                "Michigan State University",
                "George Mason University",
                "University of Southern California (USC)",
                "University of San Francisco",
                "Hofstra University",
                "University of Groningen (Netherlands)",
                "Maastricht University (Netherlands)",
                "Universiteit Leiden (Netherlands)",
                "Maurer School of Law (Indiana)",
                "East-West Center (Hawaii)",
                "Federation of Law Societies of Canada"
            });
            comboQuickDest.SelectedIndexChanged += (s, e) => ApplyQuickAddress();
            currentY += 40;

            // 2. Admission No
            Label lblId = new Label { Text = "Admission No:", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth };
            currentY += 25;
            txtAdmissionNo = new TextBox { Location = new System.Drawing.Point(padding, currentY), Width = controlWidth };
            currentY += 40;

            // 2.1 Extra ID (Dynamic)
            lblExtraID = new Label { Text = "Extra ID:", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, Visible = false };
            currentY += 25;
            txtExtraID = new TextBox { Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, Visible = false };
            currentY += 40;

            // 3. Address Textbox
            Label lblAddr = new Label { Text = "ที่อยู่ผู้รับ (Recipient Address):", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth };
            currentY += 25;
            txtAddress = new TextBox { Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, Height = 100, Multiline = true, ScrollBars = ScrollBars.Vertical };
            currentY += 110;

            // 4. Document Checklist
            Label lblDocs = new Label { Text = "รายการเอกสาร (Checklist):", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth };
            currentY += 25;
            listDocuments = new CheckedListBox { Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, Height = 180 };
            listDocuments.Items.AddRange(new[] { 
                "Official Transcript of Record (English)",
                "Official Transcript of Record (Thai)",
                "Academic Transcript (English ver.)",
                "Certificate of Academic Achievement (English)",
                "Certificate of Academic Achievement (Thai)",
                "Certificate of Student Status (English)",
                "Certificate of Student Status (Thai)",
                "Certificate of class ranking (English)",
                "Translation of Certificate (English)",
                "Verification of Names",
                "LSAC Form",
                "Replacement Diploma (Thai)",
                "Unfulfilled Transcript"
            });
            currentY += 200;

            // 5. Generate Button
            btnGenerate = new Button { Text = "Fill / Update Current Template", Location = new System.Drawing.Point(padding, currentY), Width = controlWidth, Height = 50, BackColor = System.Drawing.Color.LightGreen, Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold) };
            btnGenerate.Click += btnGenerate_Click;
            
            this.Controls.AddRange(new Control[] { lblDegree, comboDegree, lblQuick, comboQuickDest, lblId, txtAdmissionNo, lblExtraID, txtExtraID, lblAddr, txtAddress, lblDocs, listDocuments, btnGenerate });
        }

        private void ApplyQuickAddress()
        {
            string selected = comboQuickDest.Text;
            txtExtraID.Visible = false;
            lblExtraID.Visible = false;

            if (selected.Contains("LSAC"))
            {
                txtAddress.Text = "Law School Admission Council\r\nLLM Credential Assembly Service\r\n662 Penn Street, Box 8511\r\nNewtown, PA 18940-8511 USA";
                lblExtraID.Text = "LSAC Number:";
                lblExtraID.Visible = true;
                txtExtraID.Visible = true;
            }
            else if (selected.Contains("BOLE"))
            {
                txtAddress.Text = "The New York State Board of Law Examiners\r\nCorporate Plaza, Building 3\r\n254 Washington Ave. Extension\r\nAlbany, NY 12203-5195 U.S.A.";
                lblExtraID.Text = "BOLE ID:";
                lblExtraID.Visible = true;
                txtExtraID.Visible = true;
            }
            else if (selected.Contains("Evaluation Company"))
            {
                txtAddress.Text = "The Evaluation Company\r\n2400 Augusta Drive, #451\r\nHouston, TX 77057\r\nTel. +1 713-266-8805";
                lblExtraID.Text = "Evaluation ID:";
                lblExtraID.Visible = true;
                txtExtraID.Visible = true;
            }
            else if (selected.Contains("Harvard"))
            {
                txtAddress.Text = "Harvard Law School\r\nGraduate Program & International Legal Studies\r\n1585 Massachusetts Avenue, Suite 5005\r\nCambridge, MA 02138 U.S.A";
            }
            else if (selected.Contains("Michigan State"))
            {
                txtAddress.Text = "Michigan State University\r\nOffice of Admissions\r\nHannah Administration Building\r\n426 Auditorium Road, Rm 250\r\nEast Lansing, MI 48824 U.SA";
            }
            else if (selected.Contains("George Mason"))
            {
                txtAddress.Text = "Office of Admissions MS 4C8\r\nGeorge Mason University\r\n4400 University Drive\r\nFairfax, VA 22030";
            }
            else if (selected.Contains("Southern California"))
            {
                txtAddress.Text = "University of Southern California\r\nUSC office of Graduate Admission\r\n3601 South Flower Street, Room 112\r\nLos Angeles, CA 90089-0915 U.S.A.";
            }
            else if (selected.Contains("Groningen"))
            {
                txtAddress.Text = "University of Groningen\r\nAdmission Office-Student Information & Administration (SI&A)\r\nBroerstraat 5, 9712 CP Groningen, The Netherlands";
            }
            else if (selected.Contains("Maastricht"))
            {
                txtAddress.Text = "Maastricht University\r\nStudent Services Centre, UM Admissions\r\nP.O. BOX 616, 6200 MD Maastricht, The Netherlands";
            }
            else if (selected.Contains("Canada"))
            {
                txtAddress.Text = "Federation of Law Societies of Canada\r\nc/o National Committee on Accreditation\r\nWorld Exchange Plaza, 1810-45 O’Conner Street\r\nOttawa, Ontario, Canada K1P A4";
            }
            else if (selected.Contains("East-West Center"))
            {
                txtAddress.Text = "East-West Center\r\nRegistrar’s office / Graduate Degree Fellowship\r\n1601 East-West Road, John A. Burns Hall, Room 2066\r\nHonolulu, Hawaii 96848-1601 USA.";
            }
            else
            {
                txtAddress.Clear();
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                // 1. If document is empty or doesn't have our controls, load template content
                if (doc.ContentControls.Count == 0 && !doc.Content.Text.Contains("{{address}}"))
                {
                    LoadTemplateToActiveDocument(doc);
                }
                
                // 2. Page Setup
                doc.PageSetup.PageWidth = doc.Application.CentimetersToPoints(22.9f);
                doc.PageSetup.PageHeight = doc.Application.CentimetersToPoints(16.2f);
                doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

                string degree = comboDegree.Text != "(None)" ? $"({comboDegree.Text})" : "";
                string formattedDocs = string.Join("\r", listDocuments.CheckedItems.Cast<string>().Select(d => d + "\t1"));
                string extraIdText = txtExtraID.Visible ? txtExtraID.Text : "";

                // 3. Update or Replace placeholders
                UpdateOrReplace(doc, "TUREG_Address", "{{address}}", txtAddress.Text);
                UpdateOrReplace(doc, "TUREG_ID", "{{id}}", txtAdmissionNo.Text);
                UpdateOrReplace(doc, "TUREG_Degree", "{{degree}}", degree);
                UpdateOrReplace(doc, "TUREG_ExtraID", "{{extra_id}}", extraIdText);
                UpdateOrReplace(doc, "TUREG_Docs", "{{documents}}", formattedDocs);

                MessageBox.Show("อัปเดตข้อมูลเรียบร้อยแล้ว!", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void LoadTemplateToActiveDocument(Word.Document doc)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "TuregTemplate_Temp.docx");
            
            // Extract embedded resource to temp file
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = "sign_templater.Resources.EnvelopeTemplate.docx";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new Exception("ไม่พบไฟล์ Template ในโปรเจกต์ (Resource not found)");
                using (FileStream fileStream = new FileStream(tempPath, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
            }

            // Insert file content into active document
            doc.Range().InsertFile(tempPath);
            
            // Cleanup temp file
            try { File.Delete(tempPath); } catch { }
        }

        private void UpdateOrReplace(Word.Document doc, string tag, string placeholder, string newValue)
        {
            Word.ContentControls controls = doc.SelectContentControlsByTag(tag);
            if (controls.Count > 0)
            {
                controls[1].Range.Text = newValue;
                return;
            }

            Word.Range range = doc.Content;
            Word.Find findObject = range.Find;
            findObject.ClearFormatting();
            findObject.Text = placeholder;
            
            if (findObject.Execute())
            {
                Word.ContentControl cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlText, range);
                cc.Tag = tag;
                cc.Title = tag.Replace("TUREG_", "");
                cc.Range.Text = newValue;
            }
        }

        private void SidebarControl_Load(object sender, EventArgs e)
        {
        }
    }
}
