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
            comboDegree.SelectedIndex = 0; // Set default to (None)
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
                txtAddress.Text = "Law School Admission Council\vLLM Credential Assembly Service\v662 Penn Street, Box 8511\vNewtown, PA 18940-8511 USA";
                lblExtraID.Text = "LSAC Number:";
                lblExtraID.Visible = true;
                txtExtraID.Visible = true;
            }
            else if (selected.Contains("BOLE"))
            {
                txtAddress.Text = "The New York State Board of Law Examiners\vCorporate Plaza, Building 3\v254 Washington Ave. Extension\vAlbany, NY 12203-5195 U.S.A.";
                lblExtraID.Text = "BOLE ID:";
                lblExtraID.Visible = true;
                txtExtraID.Visible = true;
            }
            else if (selected.Contains("Evaluation Company"))
            {
                txtAddress.Text = "The Evaluation Company\v2400 Augusta Drive, #451\vHouston, TX 77057\vTel. +1 713-266-8805";
                lblExtraID.Text = "Evaluation ID:";
                lblExtraID.Visible = true;
                txtExtraID.Visible = true;
            }
            else if (selected.Contains("Harvard"))
            {
                txtAddress.Text = "Harvard Law School\vGraduate Program & International Legal Studies\v1585 Massachusetts Avenue, Suite 5005\vCambridge, MA 02138 U.S.A";
            }
            else if (selected.Contains("Michigan State"))
            {
                txtAddress.Text = "Michigan State University\vOffice of Admissions\vHannah Administration Building\v426 Auditorium Road, Rm 250\vEast Lansing, MI 48824 U.SA";
            }
            else if (selected.Contains("George Mason"))
            {
                txtAddress.Text = "Office of Admissions MS 4C8\vGeorge Mason University\v4400 University Drive\vFairfax, VA 22030";
            }
            else if (selected.Contains("Southern California"))
            {
                txtAddress.Text = "University of Southern California\vUSC office of Graduate Admission\v3601 South Flower Street, Room 112\vLos Angeles, CA 90089-0915 U.S.A.";
            }
            else if (selected.Contains("Groningen"))
            {
                txtAddress.Text = "University of Groningen\vAdmission Office-Student Information & Administration (SI&A)\vBroerstraat 5, 9712 CP Groningen, The Netherlands";
            }
            else if (selected.Contains("Maastricht"))
            {
                txtAddress.Text = "Maastricht University\vStudent Services Centre, UM Admissions\vP.O. BOX 616, 6200 MD Maastricht, The Netherlands";
            }
            else if (selected.Contains("Canada"))
            {
                txtAddress.Text = "Federation of Law Societies of Canada\vc/o National Committee on Accreditation\vWorld Exchange Plaza, 1810-45 O’Conner Street\vOttawa, Ontario, Canada K1P A4";
            }
            else if (selected.Contains("East-West Center"))
            {
                txtAddress.Text = "East-West Center\vRegistrar’s office / Graduate Degree Fellowship\v1601 East-West Road, John A. Burns Hall, Room 2066\vHonolulu, Hawaii 96848-1601 USA.";
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
                
                // 2. Page Setup (Forced C5 Landscape)
                foreach (Word.Section section in doc.Sections)
                {
                    section.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                    section.PageSetup.PageWidth = doc.Application.CentimetersToPoints(22.9f);
                    section.PageSetup.PageHeight = doc.Application.CentimetersToPoints(16.2f);
                    section.PageSetup.TopMargin = doc.Application.CentimetersToPoints(1.0f);
                    section.PageSetup.BottomMargin = doc.Application.CentimetersToPoints(1.0f);
                    section.PageSetup.LeftMargin = doc.Application.CentimetersToPoints(1.5f);
                    section.PageSetup.RightMargin = doc.Application.CentimetersToPoints(1.0f);
                }

                string degree = comboDegree.Text != "(None)" ? $"({comboDegree.Text.ToLower()})" : "";
                
                // Add custom prefixes for LSAC/BOLE or default to parentheses for others
                string extraIdText = "";
                if (txtExtraID.Visible && !string.IsNullOrWhiteSpace(txtExtraID.Text))
                {
                    string selected = comboQuickDest.Text;
                    if (selected.Contains("LSAC"))
                        extraIdText = $"LSAC number : {txtExtraID.Text}";
                    else if (selected.Contains("BOLE"))
                        extraIdText = $"BOLE ID: {txtExtraID.Text}";
                    else
                        extraIdText = $"({txtExtraID.Text})";
                }

                // Prepare Combined Block for ID + ExtraID + Docs to ensure Shift+Enter (\v) works correctly
                string admissionPrefix = "Admission No. ";
                string admissionLine = !string.IsNullOrWhiteSpace(txtAdmissionNo.Text) ? $"{admissionPrefix}{txtAdmissionNo.Text}" : "";
                
                // Construct the block with \v
                List<string> mainBlockParts = new List<string>();
                if (!string.IsNullOrEmpty(admissionLine)) mainBlockParts.Add(admissionLine);
                if (!string.IsNullOrEmpty(degree)) mainBlockParts.Add(degree);
                if (!string.IsNullOrEmpty(extraIdText)) mainBlockParts.Add(extraIdText);
                
                string formattedDocs = string.Join("\v", listDocuments.CheckedItems.Cast<string>().Select(d => d + "\t1"));
                if (!string.IsNullOrEmpty(formattedDocs)) mainBlockParts.Add(formattedDocs);

                string finalMainBlock = string.Join("\v", mainBlockParts);

                // 3. Update or Replace placeholders
                UpdateOrReplace(doc, "TUREG_Address", "{{address}}", txtAddress.Text);
                
                // We use TUREG_Docs as the main anchor for the combined block to ensure single paragraph formatting
                UpdateOrReplace(doc, "TUREG_Docs", "{{documents}}", finalMainBlock);
                
                // Apply Bold to "Admission No. "
                Word.ContentControls docsCCs = doc.SelectContentControlsByTag("TUREG_Docs");
                if (docsCCs.Count > 0)
                {
                    Word.ContentControl cc = docsCCs[1];
                    cc.Range.Font.Bold = 0; // Clear all bold first
                    
                    if (!string.IsNullOrEmpty(admissionLine))
                    {
                        Word.Range searchRange = cc.Range;
                        Word.Find findBold = searchRange.Find;
                        findBold.ClearFormatting();
                        findBold.Text = admissionLine;
                        
                        if (findBold.Execute())
                        {
                            searchRange.Font.Bold = 1; // This now refers to the found range only
                        }
                    }
                }
                
                // Clear other individual placeholders if they exist to avoid duplication
                UpdateOrReplace(doc, "TUREG_ID", "{{id}}", "");
                UpdateOrReplace(doc, "TUREG_Degree", "{{degree}}", "");
                UpdateOrReplace(doc, "TUREG_ExtraID", "{{extra_id}}", "");

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
            Word.ContentControl cc = null;

            if (controls.Count > 0)
            {
                cc = controls[1];
            }
            else
            {
                Word.Range range = doc.Content;
                Word.Find findObject = range.Find;
                findObject.ClearFormatting();
                findObject.Text = placeholder;

                if (findObject.Execute())
                {
                    // Use Rich Text Content Control to allow mixed formatting (Bold + Normal)
                    cc = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, range);
                    cc.Tag = tag;
                    cc.Title = tag.Replace("TUREG_", "");
                }
            }

            if (cc != null)
            {
                // Important: Use .Range.Text for Content Control to preserve and apply formatting
                cc.Range.Text = newValue;

                // Set Line Spacing to 1.0 for all placeholders
                cc.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                cc.Range.ParagraphFormat.SpaceBefore = 0;
                cc.Range.ParagraphFormat.SpaceAfter = 0;
                
                // Force Right Tab alignment for specific tags
                if (tag == "TUREG_Docs" || tag == "TUREG_ExtraID" || tag == "TUREG_ID")
                {
                    // Apply formatting to the entire range of the content control
                    Word.Range ccRange = cc.Range;
                    ccRange.ParagraphFormat.TabStops.ClearAll();
                    
                    float rightPos = doc.Application.CentimetersToPoints(20.0f); // Default fallback
                    
                    if (Convert.ToBoolean(ccRange.Information[Word.WdInformation.wdWithInTable]))
                    {
                        Word.Cell cell = ccRange.Cells[1];
                        rightPos = cell.Width - cell.RightPadding - cell.LeftPadding - 5; // Extra buffer
                    }
                    
                    ccRange.ParagraphFormat.TabStops.Add(rightPos, Word.WdTabAlignment.wdAlignTabRight);
                }
            }
        }

        private void SidebarControl_Load(object sender, EventArgs e)
        {
        }
    }
}
