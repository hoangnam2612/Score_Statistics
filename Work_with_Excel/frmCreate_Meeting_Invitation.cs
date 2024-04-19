using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Work_with_Excel
{
    public partial class frmCreate_Meeting_Invitation : Form
    {
        List<(string, string)> newListStudent;
        public frmCreate_Meeting_Invitation(List<(string, string)> listStudent)
        {
            InitializeComponent();
            newListStudent = listStudent;
        }

        string folderPath = @"C:\Users\hoang\OneDrive\Documents";

        private bool create_Docx(string studentName, string classRooom, int minute, int hour, int day, int month, int year, string address)
        {
            try
            {
                string docxFilePath = Path.Combine(folderPath, "Thư mời của " + studentName + ".docx");
                using (WordprocessingDocument doc = WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document))
                {
                    // tạo main document part
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

                    // tạo document và body
                    Document docxDocument = new Document();
                    Body docxBody = new Body();

                    //thêm tiêu đề và đoạn văn bản vào tài liệu
                    Paragraph paraTitle = new Paragraph(new Run(new Text("Thư mời họp phụ huynh")));
                    paraTitle.ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                    Run runTitle = paraTitle.GetFirstChild<Run>();
                    runTitle.RunProperties = new RunProperties(new Bold());
                    DocumentFormat.OpenXml.Wordprocessing.FontSize fontSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "50" }; // Đặt kích thước font là 50 half-points
                    runTitle.RunProperties.Append(fontSize);
                    Paragraph paraContent = new Paragraph(new Run(new Text("Kính gửi quý phụ huynh của học sinh: " + studentName +
                                                                            ". Học sinh lớp: " + classRooom + " tới dự buổi họp phụ huynh đầu năm.")));
                    paraContent.Append(new Break());
                    paraContent.Append(new Run(new Text($"Thời gian: {hour}:{minute}, ngày {day} tháng {month} năm {year}")));
                    paraContent.Append(new Break());
                    paraContent.Append(new Run(new Text($"Địa điểm: {address}")));

                    // thêm các phần tử vào Body
                    docxBody.Append(paraTitle);
                    docxBody.Append(paraContent);

                    //thêm body vào Document
                    docxDocument.Append(docxBody);

                    //Đặt Document vào main document part
                    mainPart.Document = docxDocument;

                    return true;

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private void btnPrintInvitation_Click(object sender, EventArgs e)
        {
            bool exportSuccess = false;
            int minute = dtpTimeMeeting.Value.Minute;
            int hour = dtpTimeMeeting.Value.Hour;
            int day = dtpTimeMeeting.Value.Day;
            int month = dtpTimeMeeting.Value.Month;
            int year = dtpTimeMeeting.Value.Year;
            string address = txtAddress.Text;
            foreach (var item in newListStudent)
            {
                if (create_Docx(item.Item1, item.Item2, minute, hour, day, month, year, address))
                {
                    exportSuccess = true;
                }
            }

            if (exportSuccess)
            {
                MessageBox.Show("Successfully created invitation letter in " + folderPath);
            }
            else
            {
                MessageBox.Show("Unable to create invitation");
            }
        }
    }
}
