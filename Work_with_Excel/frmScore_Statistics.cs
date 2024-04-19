using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;
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
    public partial class frmScore_Statistics : Form
    {
        private List<Avarage_Score> avarage_Scores = new List<Avarage_Score>();
        public string StudentName { get; private set; }
        public string Classes { get; private set; }
        public frmScore_Statistics()
        {
            InitializeComponent();
        }

        #region Logic

        private List<Avarage_Score> GetDataFromExcel()
        {
            OpenFileDialog fileExcel = new OpenFileDialog();
            string path = string.Empty;
            if (fileExcel.ShowDialog() == DialogResult.OK)
            {
                path = fileExcel.FileName;
            }
            try
            {
                // mo file excel
                var package = new ExcelPackage(new FileInfo(path));

                // lấy sheet đầu tiên để thao tác
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file, lưu ý file excel bắt đầu từ số 1 không phải số 0
                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    try
                    {
                        // biến j biểu thị cho một column trong file
                        int j = 1;

                        //lấy ra cột tương ứng giá trị tại vị trí [i,j]
                        // tăng j lên 1 đơn vị sau khi thực hiện xong
                        string name = worksheet.Cells[i, j++].Value.ToString();
                        string parent = worksheet.Cells[i, j++].Value.ToString();
                        string grade = worksheet.Cells[i, j++].Value.ToString();
                        string classRoom = worksheet.Cells[i, j++].Value.ToString();
                        float math_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float literature_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float english_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float physics_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float chemistry_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float biology_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float history_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());
                        float geography_score = float.Parse(worksheet.Cells[i, j++].Value.ToString());

                        Avarage_Score score = new Avarage_Score()
                        {
                            Name = name,
                            Parent = parent,
                            Grade = grade,
                            Class = classRoom,
                            Math_Score = math_score,
                            Literature_Score = literature_score,
                            EngLish_Score = english_score,
                            Physics_Score = physics_score,
                            Chemistry_Score = chemistry_score,
                            Biology_Score = biology_score,
                            History_Score = history_score,
                            Geography_Score = geography_score
                        };

                        avarage_Scores.Add(score);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return avarage_Scores;
        }

        private void CalculateAvg()
        {
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                double mathScore = Convert.ToDouble(row.Cells["Math_Score"].Value);
                double literatureScore = Convert.ToDouble(row.Cells["Literature_Score"].Value);
                double englishScore = Convert.ToDouble(row.Cells["English_Score"].Value);
                double physicsScore = Convert.ToDouble(row.Cells["Physics_Score"].Value);
                double chemistryScore = Convert.ToDouble(row.Cells["Chemistry_Score"].Value);
                double biologyScore = Convert.ToDouble(row.Cells["Biology_Score"].Value);
                double historyScore = Convert.ToDouble(row.Cells["History_Score"].Value);
                double geographyScore = Convert.ToDouble(row.Cells["Geography_Score"].Value);


                double avg = (mathScore + literatureScore + englishScore + physicsScore + chemistryScore + biologyScore + historyScore + geographyScore) / 8;

                row.Cells["Avg"].Value = avg.ToString("0.0");

                string academicPerformance = GetAcademicPerformance(avg);
                row.Cells["AcademicPerformance"].Value = academicPerformance;
            }
        }

        private string GetAcademicPerformance(double avg)
        {
            if (avg >= 8.0)
            {
                return "Very Good";
            }
            else if (avg >= 6.5)
            {
                return "Good";
            }
            else if (avg >= 5.0)
            {
                return "Average";
            }
            else
            {
                return "Weak";
            }
        }


        private void ApplyFilters()
        {
            string name = txtName.Text.Trim();

            List<Avarage_Score> filteredListByName = avarage_Scores.Where(x => x.Name.Contains(name)).ToList();

            if (!string.IsNullOrWhiteSpace(name))
            {
                dgvData.DataSource = filteredListByName;
            }
            else
            {
                dgvData.DataSource = avarage_Scores;
            }

            List<Avarage_Score> filteredListByGrade = filteredListByName;
            string classes = cbClass.Text;
            string grade = cbGrade.Text;
            if (grade != "All")
            {
                if (classes != "All")
                {
                    filteredListByGrade = filteredListByName.Where(x => x.Class == classes && x.Grade == grade).ToList();
                }
                else
                {
                    filteredListByGrade = filteredListByName.Where(x => x.Grade == grade).ToList();
                }

            }
            else
            {
                if (classes != "All")
                {
                    filteredListByGrade = filteredListByName.Where(x => x.Class == classes).ToList();
                }
            }
            dgvData.DataSource = filteredListByGrade;
        }

        private void RefreshClassComboBox()
        {
            cbClass.Items.Clear();
            cbClass.Items.Insert(0, "All");
            string selectedGrade = cbGrade.Text;

            if (selectedGrade == "All")
            {
                cbClass.Items.AddRange(avarage_Scores.Select(s => s.Class).Distinct().ToArray());
                cbClass.Sorted = true;
            }
            else
            {
                cbClass.Items.AddRange(avarage_Scores.Where(s => s.Grade.ToString() == selectedGrade)
                                                            .Select(s => s.Class)
                                                            .Distinct()
                                                            .ToArray());
                cbClass.Sorted = true;
            }
            cbClass.SelectedIndex = 0;
        }

        private void TopStudentOption()
        {
            string[] topStudent = { "All Students", "Very Good Students", "Good Students", "Average Students", "Weak Students" };
            cbTopStudent.DataSource = topStudent;
        }

        private string GetTopStudent()
        {
            string selectedValue = cbTopStudent.Text;
            string rank = "";
            switch (selectedValue)
            {
                case "Very Good Students":
                    rank = "Very Good";
                    break;
                case "Good Students":
                    rank = "Good";
                    break;
                case "Average Students":
                    rank = "Average";
                    break;
                case "Weak Students":
                    rank = "Weak";
                    break;
            }
            return rank;
        }
        #endregion


        #region Event

        private void btnImport_Click(object sender, EventArgs e)
        {
            dgvData.DataSource = GetDataFromExcel();
            var listGrade = avarage_Scores.Select(x => x.Grade).Distinct().ToList();
            listGrade.Insert(0, "All");
            cbGrade.DataSource = listGrade;
        }

        private DataTable AddDataFromDataGridViewToDataTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();

            dt.Columns.Clear();
            // Định nghĩa cấu trúc cho DataTable
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Parent", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Class", typeof(string));
            dt.Columns.Add("Math_Score", typeof(float));
            dt.Columns.Add("Literature_Score", typeof(float));
            dt.Columns.Add("EngLish_Score", typeof(float)); // Chú ý đúng tên cột
            dt.Columns.Add("Physics_Score", typeof(float));
            dt.Columns.Add("Chemistry_Score", typeof(float));
            dt.Columns.Add("Biology_Score", typeof(float));
            dt.Columns.Add("History_Score", typeof(float));
            dt.Columns.Add("Geography_Score", typeof(float));
            dt.Columns.Add("Avg", typeof(float));
            dt.Columns.Add("AcademicPerformance", typeof(string));

            if (dgv != null)
            {
                foreach (DataGridViewRow dgvRow in dgv.Rows)
                {

                    DataRow row = dt.NewRow();
                    foreach (DataGridViewCell cell in dgvRow.Cells)
                    {
                        string columnName = dgv.Columns[cell.ColumnIndex].Name;
                        if (dt.Columns.Contains(columnName))
                        {
                            if (cell.Value != null)
                            {
                                if (float.TryParse(cell.Value.ToString(), out float cellValue))
                                {
                                    row[columnName] = cellValue;
                                }
                                else
                                {
                                    row[columnName] = cell.Value.ToString();
                                }
                            }
                            else
                            {
                                row[columnName] = DBNull.Value;
                            }
                        }
                    }
                    if (!dgvRow.IsNewRow)
                    {
                        dt.Rows.Add(row);
                    }
                }
            }
            else
            {
                MessageBox.Show("Dgv is null");
            }
            return dt;
        }

        DataTable originalData;

        private void btnCAS_Click(object sender, EventArgs e)
        {
            dgvData.Columns.Add("Avg", "Avg");
            dgvData.Columns.Add("AcademicPerformance", "AcademicPerformance");

            CalculateAvg();
            DataTable dtNew = AddDataFromDataGridViewToDataTable(dgvData);
            originalData = AddDataFromDataGridViewToDataTable(dgvData);
            if (dgvData.Columns.Contains("Avg"))
            {
                dgvData.Columns.Remove("Avg");
            }
            if (dgvData.Columns.Contains("AcademicPerformance"))
            {
                dgvData.Columns.Remove("AcademicPerformance");
            }
            dgvData.DataSource = dtNew;
            TopStudentOption();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            // tạo SaveFileDialog để lưu file excel
            SaveFileDialog dialog = new SaveFileDialog();

            // chỉ lọc các file có định dạng Excel
            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

            //Nếu mở file và chọn nơi lưu file thành công sẽ lưu đường dẫn lại dùng
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                path = dialog.FileName;
            }

            //nếu đường dẫn null rỗng thì báo không hợp lệ và return hàm
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Invalid report path");
                return;
            }

            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    // đặt tiêu đề cho file
                    p.Workbook.Properties.Title = "Báo cáo thống kê";

                    // tạo 1 sheet để làm việc trên đó
                    p.Workbook.Worksheets.Add("Điểm và học lực");

                    // lấy sheet vừa add ra để thao tác
                    ExcelWorksheet ws = p.Workbook.Worksheets[0];

                    // đặt tên cho sheet
                    ws.Name = "Điểm và học lực";
                    //fontsize mặc định cho cả sheet
                    ws.Cells.Style.Font.Size = 11;
                    // font family mặc định cho cả sheet
                    ws.Cells.Style.Font.Name = "Arial";

                    // Tạo danh sách các column header
                    string[] arrColumnHeader =
                    {
                        "Name", "Parent", "Grade", "ClassRoom", "Math Score", "Literature Score", "English Score", "Physics Score",
                        "Chemistry Score","Biology Score","History Score","Geography Score","Avg", "AcademicPerformance"
                    };

                    //lấy ra số lượng cột cần dùng dựa vào số lượng header
                    var countColHeader = arrColumnHeader.Count();

                    //merge các column lại từ column 1 đến số column header
                    //gán giá trị cho cell vừa merge là Thống kê thông tin nhân viên mới
                    ws.Cells[1, 1].Value = "Thống kê điểm và học lực của học sinh";
                    ws.Cells[1, 1, 1, countColHeader].Merge = true;
                    //in đậm
                    ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                    // căn giữa
                    ws.Cells[1, 1, 1, countColHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    int colIndex = 1;
                    int rowIndex = 2;

                    //tạo các header từ column header đã tạo từ bên trên
                    foreach (var item in arrColumnHeader)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];

                        cell.Value = item;

                        colIndex++;
                    }

                    foreach (DataGridViewRow row in dgvData.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = row.Cells["Name"].Value;
                            ws.Cells[rowIndex, colIndex++].Value = row.Cells["Parent"].Value;
                            ws.Cells[rowIndex, colIndex++].Value = row.Cells["Grade"].Value;
                            ws.Cells[rowIndex, colIndex++].Value = row.Cells["Class"].Value;
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Math_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Literature_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["English_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Physics_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Chemistry_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Biology_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["History_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Geography_Score"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = Math.Round(float.Parse(row.Cells["Avg"].Value.ToString()), 1).ToString();
                            ws.Cells[rowIndex, colIndex++].Value = row.Cells["AcademicPerformance"].Value;
                        }

                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    for (int i = 2; i <= dgvData.Rows.Count + 2; i++)
                    {
                        for (int j = 1; j <= colIndex; j++)
                        {
                            ws.Cells[i, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    }


                    //Lưu file lại
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(path, bin);

                }
                MessageBox.Show("Export file excel successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        List<(string, string)> listStudent = new List<(string, string)>();

        private void btnCreateInvitation_Click(object sender, EventArgs e)
        {
            //bool exportSuccess = false;
            foreach (DataGridViewRow item in dgvData.Rows)
            {
                if (!item.IsNewRow)
                {
                    listStudent.Add((item.Cells["Name"].Value.ToString(), item.Cells["Class"].Value.ToString()));
                }
            }

            frmCreate_Meeting_Invitation createInvitation = new frmCreate_Meeting_Invitation(listStudent);
            createInvitation.Show();


        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void cbClass_SelectedValueChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void cbGrade_SelectedValueChanged(object sender, EventArgs e)
        {
            RefreshClassComboBox();
        }

        private void cbTopStudent_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedAcademicPerformance = GetTopStudent();
            DataView dv = originalData.DefaultView;
            switch (selectedAcademicPerformance)
            {
                case "Very Good":
                    dv.RowFilter = "AcademicPerformance = 'Very Good'";
                    break;
                case "Good":
                    dv.RowFilter = "AcademicPerformance = 'Good'";
                    break;
                case "Average":
                    dv.RowFilter = "AcademicPerformance = 'Average'";
                    break;
                case "Weak":
                    dv.RowFilter = "AcademicPerformance = 'Weak'";
                    break;
                default:
                    dv.RowFilter = "";
                    break;
            }
            dgvData.DataSource = dv;

        }

        private void Score_Statistics_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to exit the program?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
            }

        }

        #endregion

    }
}
