using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;
using PROJE1;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

using DocumentFormat.OpenXml.Wordprocessing;

namespace a1
{
    public partial class Form1 : Form
    {
        public Form1() => this.InitializeComponent();
        void ekleTree()
        {
            string connectionStringe = "Data Source =.; Initial Catalog = master; Integrated Security = True; Persist Security Info = False";
            string cmdTexte = "select name from sys.databases ORDER BY create_date ASC";
            SqlConnection connectione = new SqlConnection(connectionStringe);
            SqlCommand sqlCommande = new SqlCommand(cmdTexte, connectione);
            connectione.Open();
            SqlDataAdapter sqlDataAdaptere = new SqlDataAdapter(sqlCommande);
            DataTable dt1 = new DataTable();
            sqlDataAdaptere.Fill(dt1);
            foreach (DataRow dr1 in dt1.Rows)
            {

                TreeNode treeNode = new TreeNode(dr1["name"].ToString());
                treeView1.Nodes.Add(treeNode);
            }
            connectione.Close();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            treeView2.Visible = true;
            button1.Visible = true;
            string connectionString4 = "Data Source = .; Initial Catalog = " + treeView1.SelectedNode.Text.ToString() + "; Integrated Security = True; Persist Security Info = False";
            string cmdText4 = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"; 
            SqlConnection connection4 = new SqlConnection(connectionString4);
            SqlCommand sqlCommand4 = new SqlCommand(cmdText4, connection4);
            connection4.Open();
            SqlDataAdapter sqlDataAdaptere = new SqlDataAdapter(sqlCommand4);
            DataTable dt2 = new DataTable();
            sqlDataAdaptere.Fill(dt2);
            foreach (DataRow dr1 in dt2.Rows)
            {
                TreeNode treeNode = new TreeNode(dr1["TABLE_NAME"].ToString());
                treeView2.Nodes.Add(treeNode);
            }
            connection4.Close();
            label3.Text = treeView1.SelectedNode.Text.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string tabloadi = null;
            tabloadi = treeView2.SelectedNode.Text;
            string cmdText5 = "select * from " + tabloadi;
            string connectionString5 = "Data Source = .; Initial Catalog = " + label3.Text + "; Integrated Security = True; Persist Security Info = False";
            SqlConnection connection5 = new SqlConnection(connectionString5);
            SqlCommand selectCommand = new SqlCommand(cmdText5, connection5);
            connection5.Open();
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand);
            DataTable dt = new DataTable();
            sqlDataAdapter.Fill(dt);
            string AccountName = System.Environment.UserName.ToLower();
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create(@"C:/Users/" + AccountName + "/Desktop/RAPOR.docx", (WordprocessingDocumentType)0, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = new Body();
                Paragraph paragraphyazi = body.AppendChild(new Paragraph());
                Paragraph paragraph1 = body.AppendChild(new Paragraph());
                ParagraphProperties paragraphProperties1 = new ParagraphProperties(
                    new Justification() { Val = Wordprocessing.JustificationValues.Center });
                ParagraphProperties paragraphPropertiesyazi = new ParagraphProperties(
                    new Justification() { Val = Wordprocessing.JustificationValues.Both });
                paragraphyazi.AppendChild(paragraphPropertiesyazi);
                paragraph1.AppendChild(paragraphProperties1);
                Run runyazi = paragraphyazi.AppendChild(new Run());
                Run run1 = paragraph1.AppendChild(new Run());
                if (tabloadi == "Tabloadi")
                {
                    run1.Append(new Text("istenilen tablo açıklaması"));
                }
                if (tabloadi == "Tabloadi2")
                {
                   run1.Append(new Text("istenilen tablo açıklaması2"));
                }
                RunProperties runProperties = new RunProperties();
                runProperties.Append(new Bold(), new RunFonts() { Ascii = "Consolas", HighAnsi = "Consolas", ComplexScript = "Consolas" }, new FontSize() { Val = "30" });
                run1.RunProperties = runProperties;
                RunProperties runPropertiesyazi = new RunProperties();
                runPropertiesyazi.Append(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" }, new FontSize() { Val = "22" });
                runyazi.RunProperties = runPropertiesyazi;
                Table table = new Table();
                TableProperties tableProperties = new TableProperties();
                TableStyle tableStyle = new TableStyle() { Val = "LightShading-Accent1" };
                tableProperties.TableStyle = tableStyle;
                TableWidth tableWidth = new TableWidth();
             
                table.AppendChild(tableProperties);
                TableRow row = null;
                row = new TableRow();
                for (int sutun1 = 0; sutun1 < dt.Columns.Count; ++sutun1)
                {
                    TableCell tableCell = new TableCell();
                    Paragraph paragraphheader = new Paragraph();
                    ParagraphProperties paragraphPropertiesheader = new ParagraphProperties();
                    Run runheader = new Run();
                    RunProperties runPropertiesheader = new RunProperties();
                    RunFonts runFontsheader = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
                    Bold boldheader = new Bold();
                    FontSize fontSizeheader = new FontSize() { Val = "14" };
                    runPropertiesheader.Append(runFontsheader);
                    runPropertiesheader.Append(boldheader);
                    runPropertiesheader.Append(fontSizeheader);
                    Text texttablo = new Text();
                    texttablo.Text = dt.Columns[sutun1].ToString();
                    runheader.Append(runPropertiesheader);
                    runheader.Append(texttablo);
                    paragraphheader.Append(paragraphPropertiesheader);
                    paragraphheader.Append(runheader);
                  
                    tableCell.Append(paragraphheader);
               


                    TableCellProperties tableCellProperties = new TableCellProperties(
                            new TableCellWidth() { Width = "3", Type = TableWidthUnitValues.Auto },
                            new GridSpan() { Val = 6 },
                            new TopBorder() { Val = BorderValues.Single, Color = "0000" },
                            new LeftBorder() { Val = BorderValues.Single, Color = "0000" },
                            new RightBorder() { Val = BorderValues.Single, Color = "0000" },
                            new BottomBorder() { Val = BorderValues.Single, Color = "0000" }
                            );

                    TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                    tableCellProperties.Append(tableCellVerticalAlignment);
                    tableCell.Append(tableCellProperties);
                    row.Append(tableCell);
                }
                table.Append(row);
                for (int satir = 0; satir < dt.Rows.Count; satir++)
                {
                    row = new TableRow();
                    for (int sutun = 0; sutun < dt.Columns.Count; sutun++)
                    {
                        TableCell tableCell = new TableCell();
                        TableCellProperties tableCellProperties = new TableCellProperties(
                          
                            new TableCellWidth() { Width = "3", Type = TableWidthUnitValues.Auto },
                            new GridSpan() { Val = 6 },
                            new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                            new TopBorder() { Val = BorderValues.Single, Color = "0000" },
                            new LeftBorder() { Val = BorderValues.Single, Color = "0000" },
                            new RightBorder() { Val = BorderValues.Single, Color = "0000" },
                            new BottomBorder() { Val = BorderValues.Single, Color = "0000" });
                      
                        Paragraph paragraphtablo = new Paragraph();
                        ParagraphProperties paragraphPropertiestablo = new ParagraphProperties();
                        Run runtablo = new Run();
                        RunProperties runPropertiestablo = new RunProperties();
                        RunFonts runFontstablo = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
                        FontSize fontSizetablo = new FontSize() { Val = "14" };
                        runPropertiestablo.Append(runFontstablo);
                        runPropertiestablo.Append(fontSizetablo);
                        Text texttablo = new Text();
                        texttablo.Text = dt.Rows[satir][sutun].ToString();
                        runtablo.Append(runPropertiestablo);
                        runtablo.Append(texttablo);
                        paragraphtablo.Append(paragraphPropertiestablo);
                        paragraphtablo.Append(runtablo);
                       
                        tableCell.Append(paragraphtablo);

                        tableCell.Append(tableCellProperties);
                        row.Append(tableCell);
                    }
                    table.Append(row);
                }
                body.Append(table);
                wordprocessingDocument.MainDocumentPart.Document.AppendChild(body);
                wordprocessingDocument.MainDocumentPart.Document.Save();
                wordprocessingDocument.Close();
            }
            dt.Clear();
            connection5.Close();
            sqlDataAdapter.Dispose();
         
            MessageBox.Show("Tablo Oluşturma işlemi tamamlandı.");
          
        }
     


    }
}
