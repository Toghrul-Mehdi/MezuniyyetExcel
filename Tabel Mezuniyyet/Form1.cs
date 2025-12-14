namespace Tabel_Mezuniyyet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Dosyası|*.xlsx",
                Title = "Excel faylını seçin"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                // Uzun iş background-da
                await Task.Run(() => CreateSheetsFromNames(filePath));
                MessageBox.Show("İş tamamlandı");
            }
        }

        private void CreateSheetsFromNames(string filePath)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook(filePath))
            {
                var jurnalSheet = workbook.Worksheet("jurnal");
                var templateSheetName = "template";

                if (!workbook.Worksheets.Contains(templateSheetName))
                {
                    MessageBox.Show("Şablon sheet tapılmadı!");
                    return;
                }

                var templateSheet = workbook.Worksheet(templateSheetName);

                int lastRow = jurnalSheet.LastRowUsed().RowNumber();

                int adColumnIndex = 3;        // C sütunu → tam ad (Soyad Ad Ata adı)
                int vezifeColumnIndex = 4;    // D sütunu → Vəzifə
                int gunColumnIndex = 7;       // G sütunu → Məzuniyyət günləri
                int tarixColumnIndex = 9;     // I sütunu → İşə başlama tarixi
                int finColumnIndex = 10;      // J sütunu → Fin Kod
                int omarColumnIndex = 12;     // L sütunu → Unikal sheet adı üçün

                int startRow = 4; // İşçilər 4-cü sətirdən başlayır

                for (int row = startRow; row <= lastRow; row++)
                {
                    var fullName = jurnalSheet.Cell(row, adColumnIndex).GetString().Trim();
                    if (string.IsNullOrEmpty(fullName))
                        continue;

                    // Sheet adını L sütunundan götür (ömər sütunu)
                    var sheetName = jurnalSheet.Cell(row, omarColumnIndex).GetString().Trim();
                    if (string.IsNullOrEmpty(sheetName))
                        continue;

                    // Əgər sheet varsa, yeni sheet yaratma
                    if (workbook.Worksheets.Any(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    // Yeni sheet yarat və məlumatları yaz
                    var newSheet = templateSheet.CopyTo(sheetName);
                    newSheet.SetAutoFilter(false);
                    if (newSheet.AutoFilter != null)
                        newSheet.AutoFilter.Clear();

                    var vezife = jurnalSheet.Cell(row, vezifeColumnIndex).GetString().Trim();
                    var mezuniyetGun = jurnalSheet.Cell(row, gunColumnIndex).GetString().Trim();
                    var iseBaslamaTarixi = jurnalSheet.Cell(row, tarixColumnIndex).GetString().Trim();
                    var finKod = jurnalSheet.Cell(row, finColumnIndex).GetString().Trim();

                    newSheet.Cell("C2").Value = fullName;
                    newSheet.Cell("C4").Value = vezife;
                    newSheet.Cell("J27").Value = mezuniyetGun;
                    newSheet.Cell("I31").Value = $"{iseBaslamaTarixi}-ci ildən işdə hesab olunsun";
                    newSheet.Cell("M13").Value = finKod;
                }

                workbook.Save();

                MessageBox.Show("Yeni sheetlər yaradıldı və fayl üzərinə yazıldı.");
            }
        }
    }
}
