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
        private FlowLayoutPanel flowDocuments;
        private Button btnGenerate;

        public SidebarControl()
        {
            InitializeComponent();
            SetupCustomUI();
        }

        private void SetupCustomUI()
        {
            this.AutoScroll = false;
            int padding = 20;
            int controlWidth = this.Width - (padding * 2) - 15;

            // 1. Bottom Panel for Button (Always Visible)
            Panel pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 75,
                Padding = new Padding(padding, 5, padding, 15)
            };
            btnGenerate = new Button
            {
                Text = "Fill / Update Current Template",
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.LightGreen,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            btnGenerate.Click += btnGenerate_Click;
            pnlBottom.Controls.Add(btnGenerate);

            // 2. Top Flow Panel for Inputs (Auto-adjusting height)
            FlowLayoutPanel flowTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(padding, 10, padding, 10)
            };

            // Degree Level
            flowTop.Controls.Add(new Label { Text = "ระดับการศึกษา:", Width = controlWidth });
            comboDegree = new ComboBox { Width = controlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            comboDegree.Items.AddRange(new[] { "(None)", "Bachelor Degree", "Master Degree", "Graduate Diploma" });
            comboDegree.SelectedIndex = 0; 
            flowTop.Controls.Add(comboDegree);

            // Quick Select
            flowTop.Controls.Add(new Label { Text = "เลือกที่อยู่ด่วน (Destinations):", Width = controlWidth, Margin = new Padding(0, 10, 0, 0) });
            comboQuickDest = new ComboBox { Width = controlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            comboQuickDest.Items.AddRange(new[] { 
                "Manual Input (พิมพ์เอง)", "LSAC (Credential Assembly Service)", "New York State BOLE",
                "The Evaluation Company (Texas)", "Harvard Law School", "Michigan State University",
                "George Mason University", "University of Southern California (USC)", "University of San Francisco",
                "Hofstra University", "University of Groningen (Netherlands)", "Maastricht University (Netherlands)",
                "Universiteit Leiden (Netherlands)", "Maurer School of Law (Indiana)", "East-West Center (Hawaii)",
                "Federation of Law Societies of Canada"
            });
            comboQuickDest.SelectedIndexChanged += (s, e) => ApplyQuickAddress();
            flowTop.Controls.Add(comboQuickDest);

            // Admission No
            flowTop.Controls.Add(new Label { Text = "Admission No:", Width = controlWidth, Margin = new Padding(0, 10, 0, 0) });
            txtAdmissionNo = new TextBox { Width = controlWidth };
            flowTop.Controls.Add(txtAdmissionNo);

            // Extra ID (LSAC/BOLE)
            lblExtraID = new Label { Text = "Extra ID:", Width = controlWidth, Visible = false, Margin = new Padding(0, 10, 0, 0) };
            txtExtraID = new TextBox { Width = controlWidth, Visible = false };
            flowTop.Controls.Add(lblExtraID);
            flowTop.Controls.Add(txtExtraID);

            // Address
            flowTop.Controls.Add(new Label { Text = "ที่อยู่ผู้รับ (Recipient Address):", Width = controlWidth, Margin = new Padding(0, 10, 0, 0) });
            txtAddress = new TextBox { Width = controlWidth, Height = 100, Multiline = true, ScrollBars = ScrollBars.Vertical };
            flowTop.Controls.Add(txtAddress);

            // Checklist Header
            flowTop.Controls.Add(new Label { Text = "รายการเอกสาร (Checklist):", Width = controlWidth, Margin = new Padding(0, 10, 0, 5) });

            // 3. Middle Flow Panel for Documents
            flowDocuments = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = System.Drawing.Color.White,
                Padding = new Padding(5)
            };

            string[] docItems = new[] { 
                // Current Students
                "Official Transcript (Current Students) (English)\v(ใบรายงานผลการศึกษา ภาษาอังกฤษ)",
                "Student Status Certificate (English)\v(ใบรับรองการเป็นนักศึกษา ภาษาอังกฤษ)",
                "Student Status Certificate (Thai)\v(ใบรับรองการเป็นนักศึกษา ภาษาไทย)",
                "Certificate of Expected Graduation (English)\v(ใบรับรองคาดว่าจะสำเร็จการศึกษา ภาษาอังกฤษ)",
                "Certificate of Expected Graduation (Thai)\v(ใบรับรองคาดว่าจะสำเร็จการศึกษา ภาษาไทย)",
                "Certificate of Completion of All Coursework (English)\v(ใบรับรองสอบไล่ได้ครบทุกวิชา ภาษาอังกฤษ)",
                "Certificate of Completion of All Coursework (Thai)\v(ใบรับรองสอบไล่ได้ครบทุกวิชา ภาษาไทย)",
                "Associate degree Certificate (Thai)\v(ใบรับรองอนุปริญญา ภาษาไทย)",
                
                // Alumni
                "Certified Copies of Educational Documents\v(สำเนาเอกสารทางการศึกษา)",
                "Official Transcript (English)\v(ใบรายงานผลการศึกษา ภาษาอังกฤษ)",
                "Official Transcript (Thai)\v(ใบรายงานผลการศึกษา ภาษาไทย)",
                "Certificate of Academic Achievement (English)\v(หนังสือรับรองคุณวุฒิ ภาษาอังกฤษ)",
                "Certificate of Academic Achievement (Thai)\v(หนังสือรับรองคุณวุฒิ ภาษาไทย)",
                "Certified Translation of Degree Certificate (English)\v(ใบแปลปริญญาบัตร ภาษาอังกฤษ)",
                "Sealed Envelopes Service\v(ซองปิดผนึกบรรจุสำเนาเอกสารทางการศึกษา)",
                "Name Verification (English)\v(หนังสือรับรองตัวสะกด ชื่อ-นามสกุล)",
                "Certificate of Former Student Status (English)\v(ใบรับรองเคยเป็นนักศึกษา ภาษาอังกฤษ)",
                "Certificate of Former Student Status (Thai)\v(ใบรับรองเคยเป็นนักศึกษา ภาษาไทย)",
                "Transcript for a Non-Graduated Former Student (English)\v(Transcript หมดสภาพ)",
                "Replacement Diploma (Thai)\v(ใบแทนปริญญาบัตร)",
                
                // Others / Existing
                "Certificate of class ranking (English)\v(ใบรับรองลำดับที่ผลการเรียน ภาษาอังกฤษ)",
                "LSAC Form\v(แบบฟอร์ม LSAC)",
            };

            foreach (var item in docItems)
            {
                AddDocRow(item.Replace("\v", Environment.NewLine), controlWidth);
            }

            // 3.1 Manual Input Row (Add at the end of checklist)
            AddManualInputRow(controlWidth);

            // Wrapper for checklist
            Panel pnlCenter = new Panel { Dock = DockStyle.Fill, Padding = new Padding(padding, 0, padding, 0) };
            pnlCenter.Controls.Add(flowDocuments);

            // Add in order: Bottom first, then Top, then Fill
            this.Controls.Add(pnlCenter);
            this.Controls.Add(flowTop);
            this.Controls.Add(pnlBottom);
        }

        private void AddDocRow(string text, int controlWidth)
        {
            Panel row = new Panel { Width = controlWidth - 10, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
            
            CheckBox cb = new CheckBox
            {
                Text = text,
                AutoSize = true,
                Location = new System.Drawing.Point(0, 0),
                MaximumSize = new System.Drawing.Size(controlWidth - 20, 0),
                Tag = "DOC_CB"
            };

            row.Controls.Add(cb);
            flowDocuments.Controls.Add(row);
        }

        private void AddManualInputRow(int controlWidth)
        {
            Panel row = new Panel { Width = controlWidth - 10, Height = 45, Margin = new Padding(0, 10, 0, 5) };
            
            CheckBox cb = new CheckBox
            {
                Text = "อื่นๆ (Manual Input):",
                AutoSize = true,
                Location = new System.Drawing.Point(0, 0),
                Tag = "MANUAL_CB"
            };

            TextBox txt = new TextBox
            {
                Width = controlWidth - 40,
                Location = new System.Drawing.Point(20, 22),
                Tag = "MANUAL_TXT"
            };

            row.Controls.Add(cb);
            row.Controls.Add(txt);
            flowDocuments.Controls.Add(row);
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
                
                // Collect checked documents from FlowLayoutPanel
                List<string> selectedDocs = new List<string>();
                foreach (Control row in flowDocuments.Controls)
                {
                    if (row is Panel pnl)
                    {
                        CheckBox cb = pnl.Controls.OfType<CheckBox>().FirstOrDefault();
                        if (cb != null && cb.Checked && cb.Tag?.ToString() == "DOC_CB")
                        {
                            // Keep only English part (first line)
                            string englishPart = cb.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None)[0];
                            selectedDocs.Add(englishPart + "\t1");
                        }
                        
                        // Check for Manual Input TextBox
                        TextBox manualTxt = pnl.Controls.OfType<TextBox>().Where(t => t.Tag?.ToString() == "MANUAL_TXT").FirstOrDefault();
                        if (manualTxt != null && !string.IsNullOrWhiteSpace(manualTxt.Text))
                        {
                            // We need to find if the associated checkbox is checked
                            CheckBox manualCb = pnl.Controls.OfType<CheckBox>().Where(c => c.Tag?.ToString() == "MANUAL_CB").FirstOrDefault();
                            if (manualCb != null && manualCb.Checked)
                            {
                                selectedDocs.Add(manualTxt.Text + "\t1");
                            }
                        }
                    }
                }

                string formattedDocs = string.Join("\v", selectedDocs);
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
