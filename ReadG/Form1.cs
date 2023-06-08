using AForge;
using AForge.Imaging;
using AForge.Imaging.Filters;
using AForge.Math.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;



using Tesseract;


using static System.Net.Mime.MediaTypeNames;


namespace ReadG
{
    public partial class Form1 : Form
    {
        private Button btnOpenFiles;
        private Button btnExport;
        private ListBox listBox1;

        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            // Initialization code
        }

        private void BtnOpenFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Arquivos de Imagem|*.bmp;*.jpg;*.png"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    listBox1.Items.Add(fileName);
                }
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0)
            {
                string[] horizontalLabels = { "A", "B", "C", "D", "E" };
                string[] verticalLabels = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20" };

                System.Threading.Tasks.Parallel.ForEach(listBox1.Items.Cast<string>(), (imageFile) =>
                {
                    using (Bitmap originalImage = new Bitmap(imageFile))
                    {
                        // Detectar número usando o Tesseract
                        Rectangle numberRegion = new Rectangle(999, 380, 348, 175);
                        string number = DetectNumber(originalImage, numberRegion);

                        List<string> cellResults = new List<string>();

                        string outputFileName = Path.GetFileNameWithoutExtension(imageFile);

                        ProcessRectangularRegion(originalImage, horizontalLabels, verticalLabels, 164, 784, 5, 20, cellResults);
                        ProcessRectangularRegion(originalImage, horizontalLabels, verticalLabels, 1002, 785, 5, 20, cellResults);
                        ProcessRectangularRegion(originalImage, horizontalLabels, verticalLabels, 580, 781, 5, 10, cellResults);


                        // Gerar arquivo de texto com os resultados
                        string outputFilePath = Path.Combine(@"C:\READG\PROCESSO", "matricula_" + number + ".txt");
                        string outputDirectory = Path.GetDirectoryName(outputFilePath);
                        if (!Directory.Exists(outputDirectory))
                        {
                            Directory.CreateDirectory(outputDirectory);
                        }

                        using (StreamWriter writer = new StreamWriter(outputFilePath))
                        {
                            writer.WriteLine(number);
                            writer.WriteLine(outputFileName);

                            foreach (string result in cellResults)
                            {
                                writer.WriteLine(result);
                            }
                        }
                    }
                });

                MessageBox.Show("       Processamento concluído.        ");
            }
            else
            {
                MessageBox.Show("   A lista está vazia. Adicione pelo menos uma imagem antes de verificar as células.   ");
            }
        }

        private void ProcessRectangularRegion(Bitmap originalImage, string[] horizontalLabels, string[] verticalLabels, int startX, int startY, int width, int height, List<string> cellResults)
        {
            for (int i = 0; i < width; i++)
            {
                for (int j = 0; j < height; j++)
                {
                    int cellX = startX + i * 58;
                    int cellY = startY + j * 46;

                    bool hasStain = CheckCellForStain(originalImage, cellX, cellY, 58, 46);

                    string cellLabel = horizontalLabels[i] + verticalLabels[j];

                    string result = cellLabel + ": " + (hasStain ? "S" : "N");
                    cellResults.Add(result);
                }
            }
        }

        private string DetectNumber(Bitmap image, Rectangle region)
        {
            // Recortar a região de interesse da imagem
            Bitmap croppedImage = image.Clone(region, image.PixelFormat);

            // Inicializar o Tesseract com o idioma adequado
            using (var engine = new TesseractEngine("./tessdata", "eng", EngineMode.Default))
            {
                // Configurar a resolução da imagem
                engine.SetVariable("user_defined_dpi", "300");

                // Iniciar o processo do Tesseract para extrair o texto da imagem
                using (var page = engine.Process(croppedImage))
                {
                    // Obter o texto extraído
                    string extractedText = page.GetText().Trim();

                    // Filtrar apenas os dígitos usando expressão regular
                    string numbersOnly = Regex.Replace(extractedText, "[^0-9]", "");

                    // Retornar os números detectados
                    return numbersOnly;
                }
            }
        }

        private bool CheckCellForStain(Bitmap image, int x, int y, int width, int height)
        {
            int stainThreshold = 10; // Limiar de tamanho mínimo da mancha
            int stainSize = 0; // Tamanho da mancha

            for (int i = x; i < x + width; i++)
            {
                for (int j = y; j < y + height; j++)
                {
                    Color pixelColor = image.GetPixel(i, j);

                    // Verificar se o pixel está dentro das cores citadas
                    if (pixelColor.R <= 100 && pixelColor.G <= 100 && pixelColor.B <= 190)
                    {
                        stainSize++; // Incrementar o tamanho da mancha

                        // Desenhar o "X" vermelho na coordenada da mancha
                        //image.SetPixel(i, j, Color.Red);
                    }
                }
            }

            if (stainSize > stainThreshold)
            {
                //MessageBox.Show("Tamanho da mancha: " + stainSize.ToString());

                // Salvar a imagem com o "X" vermelho
                //string outputDirectory = @"C:\Users\Leonel\3D Objects";
                //string outputFileName = Path.Combine(outputDirectory, "imagem_dificada.jpg");
                //image.Save(outputFileName, System.Drawing.Imaging.ImageFormat.Jpeg);

                return true;
            }

            return false; // Retorna false quando nenhuma mancha maior que o limiar for detectada
        }
        //MessageBox.Show("Tamanho da mancha: " + stainSize.ToString());
        private void ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                _ = listBox1.SelectedItem.ToString();
                // Do something with the selected file
            }
        }
        private void ListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.D)
            {
                if (listBox1.SelectedIndex != -1)
                {
                    int selectedIndex = listBox1.SelectedIndex;
                    listBox1.Items.RemoveAt(selectedIndex);
                    if (selectedIndex < listBox1.Items.Count)
                    {
                        listBox1.SelectedIndex = selectedIndex;
                    }
                    else if (listBox1.Items.Count > 0)
                    {
                        listBox1.SelectedIndex = listBox1.Items.Count - 1;
                    }
                }
            }
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
            {
                string[] horizontalLabels = { "A", "B", "C", "D", "E" };
                string[] verticalLabels = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20" };

                string outputDirectory = @"C:\READG\PROCESSO\";
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                for (int itemIndex = 0; itemIndex < listBox1.Items.Count; itemIndex++)
                {
                    string imageFile = listBox1.Items[itemIndex].ToString();
                    Bitmap originalImage = new Bitmap(imageFile);

                    // Detectar número usando o Tesseract
                    Rectangle numberRegion = new Rectangle(1001, 380, 368, 190);
                    string number = DetectNumber(originalImage, numberRegion);

                    List<string> cellResults = new List<string>();

                    // Processar primeira região retangular
                    int startX = 168;
                    int startY = 792;

                    for (int i = 0; i < 5; i++)
                    {
                        for (int j = 0; j < 20; j++)
                        {
                            int cellX = startX + i * 58;
                            int cellY = startY + j * 46;

                            bool hasStain = CheckCellForStain(originalImage, cellX, cellY, 58, 46);

                            string cellLabel = horizontalLabels[i] + verticalLabels[j];

                            string result = cellLabel + ": " + (hasStain ? "AS IN" : "NAO IN");
                            cellResults.Add(result);
                        }
                    }

                    // Processar segunda região retangular
                    int startX2 = 1014;
                    int startY2 = 789;

                    for (int i = 0; i < 5; i++)
                    {
                        for (int j = 0; j < 20; j++)
                        {
                            int cellX = startX2 + i * 58;
                            int cellY = startY2 + j * 46;

                            bool hasStain = CheckCellForStain(originalImage, cellX, cellY, 58, 46);

                            string cellLabel = horizontalLabels[i] + verticalLabels[j];

                            string result = cellLabel + ": " + (hasStain ? "AS IN" : "NAO IN");
                            cellResults.Add(result);
                        }
                    }

                    // Processar terceira região retangular
                    int startX3 = 593;
                    int startY3 = 787;

                    for (int i = 0; i < 5; i++)
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            int cellX = startX3 + i * 58;
                            int cellY = startY3 + j * 46;

                            bool hasStain = CheckCellForStain(originalImage, cellX, cellY, 58, 46);

                            string cellLabel = horizontalLabels[i] + verticalLabels[j];

                            string result = cellLabel + ": " + (hasStain ? "AS IN" : "NAO IN");
                            cellResults.Add(result);
                        }
                    }

                    string outputFilePath = Path.Combine(outputDirectory, "matricula" + textBox2.Text + ".txt");
                    using (StreamWriter writer = new StreamWriter(outputFilePath))
                    {
                        writer.WriteLine(textBox2.Text);

                        string outputFileName = Path.GetFileNameWithoutExtension(imageFile);
                        writer.WriteLine(outputFileName);

                        foreach (string result in cellResults)
                        {
                            writer.WriteLine(result);

                        }

                        originalImage.Dispose();
                    }
                }

                listBox1.Items.RemoveAt(listBox1.SelectedIndex);

                MessageBox.Show("       Processamento concluído.        ");
            }
            else
            {
                MessageBox.Show("       A lista está vazia. Adicione pelo menos uma imagem antes de verificar as células.       ");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
                 
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Arquivos Excel (*.xlsx)|*.xlsx";
                openFileDialog.Title = "Selecione um arquivo Excel";

                string excelFilePath = openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : string.Empty;

                excelFilePath = excelFilePath ?? string.Empty;

                string folderPath = @"C:\READG\PROCESSO";
                string[] txtFiles = Directory.GetFiles(folderPath, "*.txt");

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                // Process each text file in listbox2
                for (int i = 0; i < txtFiles.Length; i++)
                {
                    string filePath = txtFiles[i];

                    string[] lines = File.ReadAllLines(filePath);

                    if (lines.Length > 0)
                    {
                        string firstLineContent = lines[0];
                        string secondLineContent = lines.Length > 1 ? lines[1] : string.Empty;
                        ModifyExcelCell(worksheet, "A" + (i + 2), firstLineContent);
                        ModifyExcelCell(worksheet, "B" + (i + 2), secondLineContent);
                    }

                    // Check if lines 3, 23, 43, 63, and 83 contain the letter 'S'
                    int sCount = 0;
                    int[] lineNumbers = { 3, 23, 43, 63, 83 };
                    int lineWithSingleS = -1;

                    foreach (int lineNumber in lineNumbers)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount++;
                                lineWithSingleS = lineNumber;
                            }
                        }
                    }

                    // Display the result based on the number of occurrences of 'S' and the position of the single 'S'
                    if (sCount == 1)
                    {
                        if (lineWithSingleS == 3)
                        {
                            ModifyExcelCell(worksheet, "C" + (i + 2), "A");
                        }
                        else if (lineWithSingleS == 23)
                        {
                            ModifyExcelCell(worksheet, "C" + (i + 2), "B");
                        }
                        else if (lineWithSingleS == 43)
                        {
                            ModifyExcelCell(worksheet, "C" + (i + 2), "C");
                        }
                        else if (lineWithSingleS == 63)
                        {
                            ModifyExcelCell(worksheet, "C" + (i + 2), "D");
                        }
                        else if (lineWithSingleS == 83)
                        {
                            ModifyExcelCell(worksheet, "C" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "C" + (i + 2), "X");
                    }

                    // Check if lines 4, 24, 44, 64, and 84 contain the letter 'S'
                    int sCount2 = 0;
                    int[] lineNumbers2 = { 4, 24, 44, 64, 84 };
                    int lineWithSingleS2 = -1;

                    foreach (int lineNumber in lineNumbers2)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2++;
                                lineWithSingleS2 = lineNumber;
                            }
                        }
                    }

                    // Display the result based on the number of occurrences of 'S' and the position of the single 'S'
                    if (sCount2 == 1)
                    {
                        if (lineWithSingleS2 == 4)
                        {
                            ModifyExcelCell(worksheet, "D" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2 == 24)
                        {
                            ModifyExcelCell(worksheet, "D" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2 == 44)
                        {
                            ModifyExcelCell(worksheet, "D" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2 == 64)
                        {
                            ModifyExcelCell(worksheet, "D" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2 == 84)
                        {
                            ModifyExcelCell(worksheet, "D" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "D" + (i + 2), "X");
                    }
                    // Check if lines 4, 24, 44, 64, and 84 contain the letter 'S'
                    int sCount3 = 0;
                    int[] lineNumbers3 = { 5, 25, 45, 65, 85 };
                    int lineWithSingleS3 = -1;

                    foreach (int lineNumber in lineNumbers3)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount3++;
                                lineWithSingleS3 = lineNumber;
                            }
                        }
                    }

                    // Display the result based on the number of occurrences of 'S' and the position of the single 'S'
                    if (sCount3 == 1)
                    {
                        if (lineWithSingleS3 == 5)
                        {
                            ModifyExcelCell(worksheet, "E" + (i + 2), "A");
                        }
                        else if (lineWithSingleS3 == 25)
                        {
                            ModifyExcelCell(worksheet, "E" + (i + 2), "B");
                        }
                        else if (lineWithSingleS3 == 45)
                        {
                            ModifyExcelCell(worksheet, "E" + (i + 2), "C");
                        }
                        else if (lineWithSingleS3 == 65)
                        {
                            ModifyExcelCell(worksheet, "E" + (i + 2), "D");
                        }
                        else if (lineWithSingleS3 == 85)
                        {
                            ModifyExcelCell(worksheet, "E" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "E" + (i + 2), "X");
                    }
                    // Check if lines 4, 24, 44, 64, and 84 contain the letter 'S'
                    int sCount4 = 0;
                    int[] lineNumbers4 = { 6, 26, 46, 66, 86 };
                    int lineWithSingleS4 = -1;

                    foreach (int lineNumber in lineNumbers4)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount4++;
                                lineWithSingleS4 = lineNumber;
                            }
                        }
                    }

                    // Display the result based on the number of occurrences of 'S' and the position of the single 'S'
                    if (sCount4 == 1)
                    {
                        if (lineWithSingleS4 == 6)
                        {
                            ModifyExcelCell(worksheet, "F" + (i + 2), "A");
                        }
                        else if (lineWithSingleS4 == 26)
                        {
                            ModifyExcelCell(worksheet, "F" + (i + 2), "B");
                        }
                        else if (lineWithSingleS4 == 46)
                        {
                            ModifyExcelCell(worksheet, "F" + (i + 2), "C");
                        }
                        else if (lineWithSingleS4 == 66)
                        {
                            ModifyExcelCell(worksheet, "F" + (i + 2), "D");
                        }
                        else if (lineWithSingleS4 == 86)
                        {
                            ModifyExcelCell(worksheet, "F" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "F" + (i + 2), "X");
                    }
                    // 5
                    int sCount5 = 0;
                    int[] lineNumbers5 = { 7, 27, 47, 67, 87 };
                    int lineWithSingleS5 = -1;

                    foreach (int lineNumber in lineNumbers5)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount5++;
                                lineWithSingleS5 = lineNumber;
                            }
                        }
                    }

                    // Display the result based on the number of occurrences of 'S' and the position of the single 'S'
                    if (sCount5 == 1)
                    {
                        if (lineWithSingleS5 == 7)
                        {
                            ModifyExcelCell(worksheet, "G" + (i + 2), "A");
                        }
                        else if (lineWithSingleS5 == 27)
                        {
                            ModifyExcelCell(worksheet, "G" + (i + 2), "B");
                        }
                        else if (lineWithSingleS5 == 47)
                        {
                            ModifyExcelCell(worksheet, "G" + (i + 2), "C");
                        }
                        else if (lineWithSingleS5 == 67)
                        {
                            ModifyExcelCell(worksheet, "G" + (i + 2), "D");
                        }
                        else if (lineWithSingleS5 == 87)
                        {
                            ModifyExcelCell(worksheet, "G" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "G" + (i + 2), "X");
                    }
                    //5
                    int sCount6 = 0;  //M
                    int[] lineNumbers6 = { 8, 28, 48, 68, 88 };    //M
                    int lineWithSingleS6 = -1;          //M

                    foreach (int lineNumber in lineNumbers6)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount6++;      //M
                                lineWithSingleS6 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount6 == 1)   //M
                    {
                        if (lineWithSingleS6 == 8)     //m
                        {
                            ModifyExcelCell(worksheet, "H" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS6 == 28)
                        {
                            ModifyExcelCell(worksheet, "H" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS6 == 48)
                        {
                            ModifyExcelCell(worksheet, "H" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS6 == 68)
                        {
                            ModifyExcelCell(worksheet, "H" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS6 == 88)
                        {
                            ModifyExcelCell(worksheet, "H" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "H" + (i + 2), "X");     //M
                    }
                    //5
                    int sCount7 = 0;  //M
                    int[] lineNumbers7 = { 9, 29, 49, 69, 89 };    //M
                    int lineWithSingleS7 = -1;          //M

                    foreach (int lineNumber in lineNumbers7)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount7++;      //M
                                lineWithSingleS7 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount7 == 1)   //M
                    {
                        if (lineWithSingleS7 == 9)          //m
                        {
                            ModifyExcelCell(worksheet, "I" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS7 == 29)
                        {
                            ModifyExcelCell(worksheet, "I" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS7 == 49)
                        {
                            ModifyExcelCell(worksheet, "I" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS7 == 69)
                        {
                            ModifyExcelCell(worksheet, "I" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS7 == 89)
                        {
                            ModifyExcelCell(worksheet, "I" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "I" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount8 = 0;  //M
                    int[] lineNumbers8 = { 10, 30, 50, 70, 90 };    //M
                    int lineWithSingleS8 = -1;          //M

                    foreach (int lineNumber in lineNumbers8)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount8++;      //M
                                lineWithSingleS8 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount8 == 1)   //M
                    {
                        if (lineWithSingleS8 == 10)          //m
                        {
                            ModifyExcelCell(worksheet, "J" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS8 == 30)    //N
                        {
                            ModifyExcelCell(worksheet, "J" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS8 == 50)  //N
                        {
                            ModifyExcelCell(worksheet, "J" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS8 == 70)  //NN
                        {
                            ModifyExcelCell(worksheet, "J" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS8 == 90)  //NN
                        {
                            ModifyExcelCell(worksheet, "J" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "J" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount9 = 0;  //M
                    int[] lineNumbers9 = { 11, 31, 51, 71, 91 };    //M
                    int lineWithSingleS9 = -1;          //M

                    foreach (int lineNumber in lineNumbers9)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount9++;      //M
                                lineWithSingleS9 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount9 == 1)   //M
                    {
                        if (lineWithSingleS9 == 11)          //m
                        {
                            ModifyExcelCell(worksheet, "K" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS9 == 31)    //N
                        {
                            ModifyExcelCell(worksheet, "K" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS9 == 51)  //N
                        {
                            ModifyExcelCell(worksheet, "K" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS9 == 71)  //NN
                        {
                            ModifyExcelCell(worksheet, "K" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS9 == 91)  //NN
                        {
                            ModifyExcelCell(worksheet, "K" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "K" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount10 = 0;  //M
                    int[] lineNumbers10 = { 12, 32, 52, 72, 92 };    //M
                    int lineWithSingleS10 = -1;          //M

                    foreach (int lineNumber in lineNumbers10)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount10++;      //M
                                lineWithSingleS10 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount10 == 1)   //M
                    {
                        if (lineWithSingleS10 == 12)          //m
                        {
                            ModifyExcelCell(worksheet, "L" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS10 == 32)    //N
                        {
                            ModifyExcelCell(worksheet, "L" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS10 == 52)  //N
                        {
                            ModifyExcelCell(worksheet, "L" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS10 == 72)  //NN
                        {
                            ModifyExcelCell(worksheet, "L" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS10 == 92)  //NN
                        {
                            ModifyExcelCell(worksheet, "L" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "L" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount11 = 0;  //M
                    int[] lineNumbers11 = { 13, 33, 53, 73, 93 };    //M
                    int lineWithSingleS11 = -1;          //M

                    foreach (int lineNumber in lineNumbers11)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount11++;      //M
                                lineWithSingleS11 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount11 == 1)   //M
                    {
                        if (lineWithSingleS11 == 13)          //m
                        {
                            ModifyExcelCell(worksheet, "M" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS11 == 33)    //N
                        {
                            ModifyExcelCell(worksheet, "M" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS11 == 53)  //N
                        {
                            ModifyExcelCell(worksheet, "M" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS11 == 73)  //NN
                        {
                            ModifyExcelCell(worksheet, "M" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS11 == 93)  //NN
                        {
                            ModifyExcelCell(worksheet, "M" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "M" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount12 = 0;  //M
                    int[] lineNumbers12 = { 14, 34, 54, 74, 94 };    //M
                    int lineWithSingleS12 = -1;          //M

                    foreach (int lineNumber in lineNumbers12)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount12++;      //M
                                lineWithSingleS12 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount12 == 1)   //M
                    {
                        if (lineWithSingleS12 == 14)          //m
                        {
                            ModifyExcelCell(worksheet, "N" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS12 == 34)    //N
                        {
                            ModifyExcelCell(worksheet, "N" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS12 == 54)  //N
                        {
                            ModifyExcelCell(worksheet, "N" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS12 == 74)  //NN
                        {
                            ModifyExcelCell(worksheet, "N" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS12 == 94)  //NN
                        {
                            ModifyExcelCell(worksheet, "N" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "N" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount13 = 0;  //M
                    int[] lineNumbers13 = { 15, 35, 55, 75, 95 };    //M
                    int lineWithSingleS13 = -1;          //M

                    foreach (int lineNumber in lineNumbers13)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13++;      //M
                                lineWithSingleS13 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount13 == 1)   //M
                    {
                        if (lineWithSingleS13 == 15)          //m
                        {
                            ModifyExcelCell(worksheet, "O" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS13 == 35)    //N
                        {
                            ModifyExcelCell(worksheet, "O" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS13 == 55)  //N
                        {
                            ModifyExcelCell(worksheet, "O" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS13 == 75)  //NN
                        {
                            ModifyExcelCell(worksheet, "O" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS13 == 95)  //NN
                        {
                            ModifyExcelCell(worksheet, "O" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "O" + (i + 2), "X");     //M
                    }
                    //8
                    int sCount14 = 0;  //M
                    int[] lineNumbers14 = { 16, 36, 56, 76, 96 };    //M
                    int lineWithSingleS14 = -1;          //M

                    foreach (int lineNumber in lineNumbers14)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount14++;      //M
                                lineWithSingleS14 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount14 == 1)   //M
                    {
                        if (lineWithSingleS14 == 16)          //m
                        {
                            ModifyExcelCell(worksheet, "P" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS14 == 36)    //N
                        {
                            ModifyExcelCell(worksheet, "P" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS14 == 56)  //N
                        {
                            ModifyExcelCell(worksheet, "P" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS14 == 76)  //NN
                        {
                            ModifyExcelCell(worksheet, "P" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS14 == 96)  //NN
                        {
                            ModifyExcelCell(worksheet, "P" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "P" + (i + 2), "X");     //M
                    }
                    int sCount15 = 0; //M
                    int[] lineNumbers15 = { 17, 37, 57, 77, 97 }; //M
                    int lineWithSingleS15 = -1; //M

                    foreach (int lineNumber in lineNumbers15) //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount15++; //M
                                lineWithSingleS15 = lineNumber; //M
                            }
                        }
                    }

                    if (sCount15 == 1) //M
                    {
                        if (lineWithSingleS15 == 17) //m
                        {
                            ModifyExcelCell(worksheet, "Q" + (i + 2), "A"); //M LETRA
                        }
                        else if (lineWithSingleS15 == 37) //N
                        {
                            ModifyExcelCell(worksheet, "Q" + (i + 2), "B"); //M LETRA
                        }
                        else if (lineWithSingleS15 == 57) //N
                        {
                            ModifyExcelCell(worksheet, "Q" + (i + 2), "C"); //M
                        }
                        else if (lineWithSingleS15 == 77) //NN
                        {
                            ModifyExcelCell(worksheet, "Q" + (i + 2), "D"); //M
                        }
                        else if (lineWithSingleS15 == 97) //NN
                        {
                            ModifyExcelCell(worksheet, "Q" + (i + 2), "E"); //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "Q" + (i + 2), "X"); //M
                    }
                    int sCount13_1 = 0;  //M
                    int[] lineNumbers13_1 = { 18, 38, 58, 78, 98 };    //M
                    int lineWithSingleS13_1 = -1;          //M

                    foreach (int lineNumber in lineNumbers13_1)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_1++;      //M
                                lineWithSingleS13_1 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount13_1 == 1)   //M
                    {
                        if (lineWithSingleS13_1 == 18)          //m
                        {
                            ModifyExcelCell(worksheet, "R" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS13_1 == 38)    //N
                        {
                            ModifyExcelCell(worksheet, "R" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS13_1 == 58)  //N
                        {
                            ModifyExcelCell(worksheet, "R" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS13_1 == 78)  //NN
                        {
                            ModifyExcelCell(worksheet, "R" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS13_1 == 98)  //NN
                        {
                            ModifyExcelCell(worksheet, "R" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "R" + (i + 2), "X");     //M
                    }
                    int sCount13_2 = 0;
                    int[] lineNumbers13_2 = { 19, 39, 59, 79, 99 };
                    int lineWithSingleS13_2 = -1;

                    foreach (int lineNumber in lineNumbers13_2)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_2++;
                                lineWithSingleS13_2 = lineNumber;
                            }
                        }
                    }

                    if (sCount13_2 == 1)
                    {
                        if (lineWithSingleS13_2 == 19)
                        {
                            ModifyExcelCell(worksheet, "S" + (i + 2), "A");
                        }
                        else if (lineWithSingleS13_2 == 39)
                        {
                            ModifyExcelCell(worksheet, "S" + (i + 2), "B");
                        }
                        else if (lineWithSingleS13_2 == 59)
                        {
                            ModifyExcelCell(worksheet, "S" + (i + 2), "C");
                        }
                        else if (lineWithSingleS13_2 == 79)
                        {
                            ModifyExcelCell(worksheet, "S" + (i + 2), "D");
                        }
                        else if (lineWithSingleS13_2 == 99)
                        {
                            ModifyExcelCell(worksheet, "S" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "S" + (i + 2), "X");
                    }
                    int sCount13_3 = 0;
                    int[] lineNumbers13_3 = { 20, 40, 60, 80, 100 };
                    int lineWithSingleS13_3 = -1;

                    foreach (int lineNumber in lineNumbers13_3)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_3++;
                                lineWithSingleS13_3 = lineNumber;
                            }
                        }
                    }

                    if (sCount13_3 == 1)
                    {
                        if (lineWithSingleS13_3 == 20)
                        {
                            ModifyExcelCell(worksheet, "T" + (i + 2), "A");
                        }
                        else if (lineWithSingleS13_3 == 40)
                        {
                            ModifyExcelCell(worksheet, "T" + (i + 2), "B");
                        }
                        else if (lineWithSingleS13_3 == 60)
                        {
                            ModifyExcelCell(worksheet, "T" + (i + 2), "C");
                        }
                        else if (lineWithSingleS13_3 == 80)
                        {
                            ModifyExcelCell(worksheet, "T" + (i + 2), "D");
                        }
                        else if (lineWithSingleS13_3 == 100)
                        {
                            ModifyExcelCell(worksheet, "T" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "T" + (i + 2), "X");
                    }
                    int sCount13_4 = 0;  //M
                    int[] lineNumbers13_4 = { 21, 41, 61, 81, 101 };    //M
                    int lineWithSingleS13_4 = -1;          //M

                    foreach (int lineNumber in lineNumbers13_4)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_4++;      //M
                                lineWithSingleS13_4 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount13_4 == 1)   //M
                    {
                        if (lineWithSingleS13_4 == 21)          //m
                        {
                            ModifyExcelCell(worksheet, "U" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS13_4 == 41)    //N
                        {
                            ModifyExcelCell(worksheet, "U" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS13_4 == 61)  //N
                        {
                            ModifyExcelCell(worksheet, "U" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS13_4 == 81)  //NN
                        {
                            ModifyExcelCell(worksheet, "U" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS13_4 == 101)  //NN
                        {
                            ModifyExcelCell(worksheet, "U" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "U" + (i + 2), "X");     //M
                    }
                    int sCount13_5 = 0;  //M
                    int[] lineNumbers13_5 = { 22, 42, 62, 82, 102 };    //M
                    int lineWithSingleS13_5 = -1;          //M

                    foreach (int lineNumber in lineNumbers13_5)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_5++;      //M
                                lineWithSingleS13_5 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount13_5 == 1)   //M
                    {
                        if (lineWithSingleS13_5 == 22)          //m
                        {
                            ModifyExcelCell(worksheet, "V" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS13_5 == 42)    //N
                        {
                            ModifyExcelCell(worksheet, "V" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS13_5 == 62)  //N
                        {
                            ModifyExcelCell(worksheet, "V" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS13_5 == 82)  //NN
                        {
                            ModifyExcelCell(worksheet, "V" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS13_5 == 102)  //NN
                        {
                            ModifyExcelCell(worksheet, "V" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "V" + (i + 2), "X");     //M
                    }
                    int sCount13_190 = 0;  //M
                    int[] lineNumbers13_190 = { 203, 213, 223, 233, 243 };    //M
                    int lineWithSingleS13_190 = -1;          //M

                    foreach (int lineNumber in lineNumbers13_190)  //M
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_190++;      //M
                                lineWithSingleS13_190 = lineNumber;      //M
                            }
                        }
                    }

                    if (sCount13_190 == 1)   //M
                    {
                        if (lineWithSingleS13_190 == 203)          //m
                        {
                            ModifyExcelCell(worksheet, "W" + (i + 2), "A");        //M LETRA
                        }
                        else if (lineWithSingleS13_190 == 213)    //N
                        {
                            ModifyExcelCell(worksheet, "W" + (i + 2), "B");         //M LETRA
                        }
                        else if (lineWithSingleS13_190 == 223)  //N
                        {
                            ModifyExcelCell(worksheet, "W" + (i + 2), "C");     //M
                        }
                        else if (lineWithSingleS13_190 == 233)  //NN
                        {
                            ModifyExcelCell(worksheet, "W" + (i + 2), "D");     //M
                        }
                        else if (lineWithSingleS13_190 == 243)  //NN
                        {
                            ModifyExcelCell(worksheet, "W" + (i + 2), "E");     //M
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "W" + (i + 2), "X");     //M
                    }
                    int sCount13_191 = 0;
                    int[] lineNumbers13_191 = { 204, 214, 224, 234, 244 };
                    int lineWithSingleS13_191 = -1;

                    foreach (int lineNumber in lineNumbers13_191)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount13_191++;
                                lineWithSingleS13_191 = lineNumber;
                            }
                        }
                    }

                    if (sCount13_191 == 1)
                    {
                        if (lineWithSingleS13_191 == 204)
                        {
                            ModifyExcelCell(worksheet, "X" + (i + 2), "A");
                        }
                        else if (lineWithSingleS13_191 == 214)
                        {
                            ModifyExcelCell(worksheet, "X" + (i + 2), "B");
                        }
                        else if (lineWithSingleS13_191 == 224)
                        {
                            ModifyExcelCell(worksheet, "X" + (i + 2), "C");
                        }
                        else if (lineWithSingleS13_191 == 234)
                        {
                            ModifyExcelCell(worksheet, "X" + (i + 2), "D");
                        }
                        else if (lineWithSingleS13_191 == 244)
                        {
                            ModifyExcelCell(worksheet, "X" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "X" + (i + 2), "X");
                    }
                    int sCount192 = 0;
                    int[] lineNumbers192 = { 205, 215, 225, 235, 245 };
                    int lineWithSingleS192 = -1;

                    foreach (int lineNumber in lineNumbers192)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount192++;
                                lineWithSingleS192 = lineNumber;
                            }
                        }
                    }

                    if (sCount192 == 1)
                    {
                        if (lineWithSingleS192 == 205)
                        {
                            ModifyExcelCell(worksheet, "Y" + (i + 2), "A");
                        }
                        else if (lineWithSingleS192 == 215)
                        {
                            ModifyExcelCell(worksheet, "Y" + (i + 2), "B");
                        }
                        else if (lineWithSingleS192 == 225)
                        {
                            ModifyExcelCell(worksheet, "Y" + (i + 2), "C");
                        }
                        else if (lineWithSingleS192 == 235)
                        {
                            ModifyExcelCell(worksheet, "Y" + (i + 2), "D");
                        }
                        else if (lineWithSingleS192 == 245)
                        {
                            ModifyExcelCell(worksheet, "Y" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "Y" + (i + 2), "X");     //M
                    }
                    //z
                    int sCount1920 = 0;
                    int[] lineNumbers1920 = { 206, 216, 226, 236, 246 };
                    int lineWithSingleS1920 = -1;

                    foreach (int lineNumber in lineNumbers1920)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount1920++;
                                lineWithSingleS1920 = lineNumber;
                            }
                        }
                    }

                    if (sCount1920 == 1)
                    {
                        if (lineWithSingleS1920 == 206)
                        {
                            ModifyExcelCell(worksheet, "Z" + (i + 2), "A");
                        }
                        else if (lineWithSingleS1920 == 216)
                        {
                            ModifyExcelCell(worksheet, "Z" + (i + 2), "B");
                        }
                        else if (lineWithSingleS1920 == 226)
                        {
                            ModifyExcelCell(worksheet, "Z" + (i + 2), "C");
                        }
                        else if (lineWithSingleS1920 == 236)
                        {
                            ModifyExcelCell(worksheet, "Z" + (i + 2), "D");
                        }
                        else if (lineWithSingleS1920 == 246)
                        {
                            ModifyExcelCell(worksheet, "Z" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "Z" + (i + 2), "X");     //M
                    }
                    //GLORY
                    int sCount193 = 0;
                    int[] lineNumbersS193 = { 207, 217, 227, 237, 247 };
                    int lineWithSingleS193 = -1;

                    foreach (int lineNumber in lineNumbersS193)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount193++;
                                lineWithSingleS193 = lineNumber;
                            }
                        }
                    }

                    if (sCount193 == 1)
                    {
                        if (lineWithSingleS193 == 207)
                        {
                            ModifyExcelCell(worksheet, "AA" + (i + 2), "A");
                        }
                        else if (lineWithSingleS193 == 217)
                        {
                            ModifyExcelCell(worksheet, "AA" + (i + 2), "B");
                        }
                        else if (lineWithSingleS193 == 227)
                        {
                            ModifyExcelCell(worksheet, "AA" + (i + 2), "C");
                        }
                        else if (lineWithSingleS193 == 237)
                        {
                            ModifyExcelCell(worksheet, "AA" + (i + 2), "D");
                        }
                        else if (lineWithSingleS193 == 247)
                        {
                            ModifyExcelCell(worksheet, "AA" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AA" + (i + 2), "X");     //M
                    }
                    int sCount194 = 0;
                    int[] lineNumbersS194 = { 208, 218, 228, 238, 248 };
                    int lineWithSingleS194 = -1;

                    foreach (int lineNumber in lineNumbersS194)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount194++;
                                lineWithSingleS194 = lineNumber;
                            }
                        }
                    }

                    if (sCount194 == 1)
                    {
                        if (lineWithSingleS194 == 208)
                        {
                            ModifyExcelCell(worksheet, "AB" + (i + 2), "A");
                        }
                        else if (lineWithSingleS194 == 218)
                        {
                            ModifyExcelCell(worksheet, "AB" + (i + 2), "B");
                        }
                        else if (lineWithSingleS194 == 228)
                        {
                            ModifyExcelCell(worksheet, "AB" + (i + 2), "C");
                        }
                        else if (lineWithSingleS194 == 238)
                        {
                            ModifyExcelCell(worksheet, "AB" + (i + 2), "D");
                        }
                        else if (lineWithSingleS194 == 248)
                        {
                            ModifyExcelCell(worksheet, "AB" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AB" + (i + 2), "X");     //M
                    }
                    int sCount195 = 0;
                    int[] lineNumbersS195 = { 209, 219, 229, 239, 249 };
                    int lineWithSingleS195 = -1;

                    foreach (int lineNumber in lineNumbersS195)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount195++;
                                lineWithSingleS195 = lineNumber;
                            }
                        }
                    }

                    if (sCount195 == 1)
                    {
                        if (lineWithSingleS195 == 209)
                        {
                            ModifyExcelCell(worksheet, "AC" + (i + 2), "A");
                        }
                        else if (lineWithSingleS195 == 219)
                        {
                            ModifyExcelCell(worksheet, "AC" + (i + 2), "B");
                        }
                        else if (lineWithSingleS195 == 229)
                        {
                            ModifyExcelCell(worksheet, "AC" + (i + 2), "C");
                        }
                        else if (lineWithSingleS195 == 239)
                        {
                            ModifyExcelCell(worksheet, "AC" + (i + 2), "D");
                        }
                        else if (lineWithSingleS195 == 249)
                        {
                            ModifyExcelCell(worksheet, "AC" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AC" + (i + 2), "X");     //M
                    }
                    int sCount196 = 0;
                    int[] lineNumbersS196 = { 210, 220, 230, 240, 250 };
                    int lineWithSingleS196 = -1;

                    foreach (int lineNumber in lineNumbersS196)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount196++;
                                lineWithSingleS196 = lineNumber;
                            }
                        }
                    }

                    if (sCount196 == 1)
                    {
                        if (lineWithSingleS196 == 210)
                        {
                            ModifyExcelCell(worksheet, "AD" + (i + 2), "A");
                        }
                        else if (lineWithSingleS196 == 220)
                        {
                            ModifyExcelCell(worksheet, "AD" + (i + 2), "B");
                        }
                        else if (lineWithSingleS196 == 230)
                        {
                            ModifyExcelCell(worksheet, "AD" + (i + 2), "C");
                        }
                        else if (lineWithSingleS196 == 240)
                        {
                            ModifyExcelCell(worksheet, "AD" + (i + 2), "D");
                        }
                        else if (lineWithSingleS196 == 250)
                        {
                            ModifyExcelCell(worksheet, "AD" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AD" + (i + 2), "X");     //M
                    }
                    int sCount197 = 0;
                    int[] lineNumbersS197 = { 211, 221, 231, 241, 251 };
                    int lineWithSingleS197 = -1;

                    foreach (int lineNumber in lineNumbersS197)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount197++;
                                lineWithSingleS197 = lineNumber;
                            }
                        }
                    }

                    if (sCount197 == 1)
                    {
                        if (lineWithSingleS197 == 211)
                        {
                            ModifyExcelCell(worksheet, "AE" + (i + 2), "A");
                        }
                        else if (lineWithSingleS197 == 221)
                        {
                            ModifyExcelCell(worksheet, "AE" + (i + 2), "B");
                        }
                        else if (lineWithSingleS197 == 231)
                        {
                            ModifyExcelCell(worksheet, "AE" + (i + 2), "C");
                        }
                        else if (lineWithSingleS197 == 241)
                        {
                            ModifyExcelCell(worksheet, "AE" + (i + 2), "D");
                        }
                        else if (lineWithSingleS197 == 251)
                        {
                            ModifyExcelCell(worksheet, "AE" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AE" + (i + 2), "X");     //M
                    }
                    int sCount198 = 0;
                    int[] lineNumbersS198 = { 212, 222, 232, 242, 252 };
                    int lineWithSingleS198 = -1;

                    foreach (int lineNumber in lineNumbersS198)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount198++;
                                lineWithSingleS198 = lineNumber;
                            }
                        }
                    }

                    if (sCount198 == 1)
                    {
                        if (lineWithSingleS198 == 212)
                        {
                            ModifyExcelCell(worksheet, "AF" + (i + 2), "A");
                        }
                        else if (lineWithSingleS198 == 222)
                        {
                            ModifyExcelCell(worksheet, "AF" + (i + 2), "B");
                        }
                        else if (lineWithSingleS198 == 232)
                        {
                            ModifyExcelCell(worksheet, "AF" + (i + 2), "C");
                        }
                        else if (lineWithSingleS198 == 242)
                        {
                            ModifyExcelCell(worksheet, "AF" + (i + 2), "D");
                        }
                        else if (lineWithSingleS198 == 252)
                        {
                            ModifyExcelCell(worksheet, "AF" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AF" + (i + 2), "X");     //M
                    }

                    int sCount200 = 0;
                    int[] lineNumbersS200 = { 103, 123, 143, 163, 183 };
                    int lineWithSingleS200 = -1;

                    foreach (int lineNumber in lineNumbersS200)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount200++;
                                lineWithSingleS200 = lineNumber;
                            }
                        }
                    }

                    if (sCount200 == 1)
                    {
                        if (lineWithSingleS200 == 103)
                        {
                            ModifyExcelCell(worksheet, "AG" + (i + 2), "A");
                        }
                        else if (lineWithSingleS200 == 123)
                        {
                            ModifyExcelCell(worksheet, "AG" + (i + 2), "B");
                        }
                        else if (lineWithSingleS200 == 143)
                        {
                            ModifyExcelCell(worksheet, "AG" + (i + 2), "C");
                        }
                        else if (lineWithSingleS200 == 163)
                        {
                            ModifyExcelCell(worksheet, "AG" + (i + 2), "D");
                        }
                        else if (lineWithSingleS200 == 183)
                        {
                            ModifyExcelCell(worksheet, "AG" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AG" + (i + 2), "X");     //M
                    }
                    int sCount201 = 0;
                    int[] lineNumbersS201 = { 104, 124, 144, 164, 184 };
                    int lineWithSingleS201 = -1;

                    foreach (int lineNumber in lineNumbersS201)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount201++;
                                lineWithSingleS201 = lineNumber;
                            }
                        }
                    }

                    if (sCount201 == 1)
                    {
                        if (lineWithSingleS201 == 104)
                        {
                            ModifyExcelCell(worksheet, "AH" + (i + 2), "A");
                        }
                        else if (lineWithSingleS201 == 124)
                        {
                            ModifyExcelCell(worksheet, "AH" + (i + 2), "B");
                        }
                        else if (lineWithSingleS201 == 144)
                        {
                            ModifyExcelCell(worksheet, "AH" + (i + 2), "C");
                        }
                        else if (lineWithSingleS201 == 164)
                        {
                            ModifyExcelCell(worksheet, "AH" + (i + 2), "D");
                        }
                        else if (lineWithSingleS201 == 184)
                        {
                            ModifyExcelCell(worksheet, "AH" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AH" + (i + 2), "X");
                    }

                    int sCount204 = 0;
                    int[] lineNumbersS204 = { 105, 125, 145, 165, 185 };
                    int lineWithSingleS204 = -1;

                    foreach (int lineNumber in lineNumbersS204)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount204++;
                                lineWithSingleS204 = lineNumber;
                            }
                        }
                    }

                    if (sCount204 == 1)
                    {
                        if (lineWithSingleS204 == 105)
                        {
                            ModifyExcelCell(worksheet, "AI" + (i + 2), "A");
                        }
                        else if (lineWithSingleS204 == 125)
                        {
                            ModifyExcelCell(worksheet, "AI" + (i + 2), "B");
                        }
                        else if (lineWithSingleS204 == 145)
                        {
                            ModifyExcelCell(worksheet, "AI" + (i + 2), "C");
                        }
                        else if (lineWithSingleS204 == 165)
                        {
                            ModifyExcelCell(worksheet, "AI" + (i + 2), "D");
                        }
                        else if (lineWithSingleS204 == 185)
                        {
                            ModifyExcelCell(worksheet, "AI" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AI" + (i + 2), "X");
                    }
                    int sCount206 = 0;
                    int[] lineNumbersS206 = { 106, 126, 146, 166, 186 };
                    int lineWithSingleS206 = -1;

                    foreach (int lineNumber in lineNumbersS206)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount206++;
                                lineWithSingleS206 = lineNumber;
                            }
                        }
                    }

                    if (sCount206 == 1)
                    {
                        if (lineWithSingleS206 == 106)
                        {
                            ModifyExcelCell(worksheet, "AJ" + (i + 2), "A");
                        }
                        else if (lineWithSingleS206 == 126)
                        {
                            ModifyExcelCell(worksheet, "AJ" + (i + 2), "B");
                        }
                        else if (lineWithSingleS206 == 146)
                        {
                            ModifyExcelCell(worksheet, "AJ" + (i + 2), "C");
                        }
                        else if (lineWithSingleS206 == 166)
                        {
                            ModifyExcelCell(worksheet, "AJ" + (i + 2), "D");
                        }
                        else if (lineWithSingleS206 == 186)
                        {
                            ModifyExcelCell(worksheet, "AJ" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AJ" + (i + 2), "X");
                    }
                    int sCount2004 = 0;
                    int[] lineNumbersS2004 = { 107, 127, 147, 167, 187 };
                    int lineWithSingleS2004 = -1;

                    foreach (int lineNumber in lineNumbersS2004)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2004++;
                                lineWithSingleS2004 = lineNumber;
                            }
                        }
                    }

                    if (sCount2004 == 1)
                    {
                        if (lineWithSingleS2004 == 107)
                        {
                            ModifyExcelCell(worksheet, "AK" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2004 == 127)
                        {
                            ModifyExcelCell(worksheet, "AK" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2004 == 147)
                        {
                            ModifyExcelCell(worksheet, "AK" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2004 == 167)
                        {
                            ModifyExcelCell(worksheet, "AK" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2004 == 187)
                        {
                            ModifyExcelCell(worksheet, "AK" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AK" + (i + 2), "X");
                    }
                    int sCount2005 = 0;
                    int[] lineNumbersS2005 = { 108, 128, 148, 168, 188 };
                    int lineWithSingleS2005 = -1;

                    foreach (int lineNumber in lineNumbersS2005)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2005++;
                                lineWithSingleS2005 = lineNumber;
                            }
                        }
                    }

                    if (sCount2005 == 1)
                    {
                        if (lineWithSingleS2005 == 108)
                        {
                            ModifyExcelCell(worksheet, "AL" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2005 == 128)
                        {
                            ModifyExcelCell(worksheet, "AL" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2005 == 148)
                        {
                            ModifyExcelCell(worksheet, "AL" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2005 == 168)
                        {
                            ModifyExcelCell(worksheet, "AL" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2005 == 188)
                        {
                            ModifyExcelCell(worksheet, "AL" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AL" + (i + 2), "X");
                    }
                    int sCount2006 = 0;
                    int[] lineNumbersS2006 = { 109, 129, 149, 169, 189 };
                    int lineWithSingleS2006 = -1;

                    foreach (int lineNumber in lineNumbersS2006)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2006++;
                                lineWithSingleS2006 = lineNumber;
                            }
                        }
                    }

                    if (sCount2006 == 1)
                    {
                        if (lineWithSingleS2006 == 109)
                        {
                            ModifyExcelCell(worksheet, "AM" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2006 == 129)
                        {
                            ModifyExcelCell(worksheet, "AM" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2006 == 149)
                        {
                            ModifyExcelCell(worksheet, "AM" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2006 == 169)
                        {
                            ModifyExcelCell(worksheet, "AM" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2006 == 189)
                        {
                            ModifyExcelCell(worksheet, "AM" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AM" + (i + 2), "X");
                    }
                    int sCount207 = 0;
                    int[] lineNumbersS207 = { 110, 130, 150, 170, 190 };
                    int lineWithSingleS207 = -1;

                    foreach (int lineNumber in lineNumbersS207)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount207++;
                                lineWithSingleS207 = lineNumber;
                            }
                        }
                    }

                    if (sCount207 == 1)
                    {
                        if (lineWithSingleS207 == 110)
                        {
                            ModifyExcelCell(worksheet, "AN" + (i + 2), "A");
                        }
                        else if (lineWithSingleS207 == 130)
                        {
                            ModifyExcelCell(worksheet, "AN" + (i + 2), "B");
                        }
                        else if (lineWithSingleS207 == 150)
                        {
                            ModifyExcelCell(worksheet, "AN" + (i + 2), "C");
                        }
                        else if (lineWithSingleS207 == 170)
                        {
                            ModifyExcelCell(worksheet, "AN" + (i + 2), "D");
                        }
                        else if (lineWithSingleS207 == 190)
                        {
                            ModifyExcelCell(worksheet, "AN" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AN" + (i + 2), "X");
                    }
                    int sCount208 = 0;
                    int[] lineNumbersS208 = { 111, 131, 151, 171, 191 };
                    int lineWithSingleS208 = -1;

                    foreach (int lineNumber in lineNumbersS208)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount208++;
                                lineWithSingleS208 = lineNumber;
                            }
                        }
                    }

                    if (sCount208 == 1)
                    {
                        if (lineWithSingleS208 == 111)
                        {
                            ModifyExcelCell(worksheet, "AO" + (i + 2), "A");
                        }
                        else if (lineWithSingleS208 == 131)
                        {
                            ModifyExcelCell(worksheet, "AO" + (i + 2), "B");
                        }
                        else if (lineWithSingleS208 == 151)
                        {
                            ModifyExcelCell(worksheet, "AO" + (i + 2), "C");
                        }
                        else if (lineWithSingleS208 == 171)
                        {
                            ModifyExcelCell(worksheet, "AO" + (i + 2), "D");
                        }
                        else if (lineWithSingleS208 == 191)
                        {
                            ModifyExcelCell(worksheet, "AO" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AO" + (i + 2), "X");
                    }
                    int sCount209 = 0;
                    int[] lineNumbersS209 = { 112, 132, 152, 172, 192 };
                    int lineWithSingleS209 = -1;

                    foreach (int lineNumber in lineNumbersS209)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount209++;
                                lineWithSingleS209 = lineNumber;
                            }
                        }
                    }

                    if (sCount209 == 1)
                    {
                        if (lineWithSingleS209 == 112)
                        {
                            ModifyExcelCell(worksheet, "AP" + (i + 2), "A");
                        }
                        else if (lineWithSingleS209 == 132)
                        {
                            ModifyExcelCell(worksheet, "AP" + (i + 2), "B");
                        }
                        else if (lineWithSingleS209 == 152)
                        {
                            ModifyExcelCell(worksheet, "AP" + (i + 2), "C");
                        }
                        else if (lineWithSingleS209 == 172)
                        {
                            ModifyExcelCell(worksheet, "AP" + (i + 2), "D");
                        }
                        else if (lineWithSingleS209 == 192)
                        {
                            ModifyExcelCell(worksheet, "AP" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AP" + (i + 2), "X");
                    }
                    int sCount2010 = 0;
                    int[] lineNumbersS2010 = { 113, 133, 153, 173, 193 };
                    int lineWithSingleS2010 = -1;

                    foreach (int lineNumber in lineNumbersS2010)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2010++;
                                lineWithSingleS2010 = lineNumber;
                            }
                        }
                    }

                    if (sCount2010 == 1)
                    {
                        if (lineWithSingleS2010 == 113)
                        {
                            ModifyExcelCell(worksheet, "AQ" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2010 == 133)
                        {
                            ModifyExcelCell(worksheet, "AQ" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2010 == 153)
                        {
                            ModifyExcelCell(worksheet, "AQ" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2010 == 173)
                        {
                            ModifyExcelCell(worksheet, "AQ" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2010 == 193)
                        {
                            ModifyExcelCell(worksheet, "AQ" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AQ" + (i + 2), "X");
                    }
                    int sCount2011 = 0;
                    int[] lineNumbersS2011 = { 114, 134, 154, 174, 194 };
                    int lineWithSingleS2011 = -1;

                    foreach (int lineNumber in lineNumbersS2011)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2011++;
                                lineWithSingleS2011 = lineNumber;
                            }
                        }
                    }

                    if (sCount2011 == 1)
                    {
                        if (lineWithSingleS2011 == 114)
                        {
                            ModifyExcelCell(worksheet, "AR" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2011 == 134)
                        {
                            ModifyExcelCell(worksheet, "AR" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2011 == 154)
                        {
                            ModifyExcelCell(worksheet, "AR" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2011 == 174)
                        {
                            ModifyExcelCell(worksheet, "AR" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2011 == 194)
                        {
                            ModifyExcelCell(worksheet, "AR" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AR" + (i + 2), "X");
                    }
                    int sCount2012 = 0;
                    int[] lineNumbersS2012 = { 115, 135, 155, 175, 195 };
                    int lineWithSingleS2012 = -1;

                    foreach (int lineNumber in lineNumbersS2012)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2012++;
                                lineWithSingleS2012 = lineNumber;
                            }
                        }
                    }

                    if (sCount2012 == 1)
                    {
                        if (lineWithSingleS2012 == 115)
                        {
                            ModifyExcelCell(worksheet, "AS" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2012 == 135)
                        {
                            ModifyExcelCell(worksheet, "AS" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2012 == 155)
                        {
                            ModifyExcelCell(worksheet, "AS" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2012 == 175)
                        {
                            ModifyExcelCell(worksheet, "AS" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2012 == 195)
                        {
                            ModifyExcelCell(worksheet, "AS" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AS" + (i + 2), "X");
                    }
                    int sCount2013 = 0;
                    int[] lineNumbersS2013 = { 116, 136, 156, 176, 196 };
                    int lineWithSingleS2013 = -1;

                    foreach (int lineNumber in lineNumbersS2013)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2013++;
                                lineWithSingleS2013 = lineNumber;
                            }
                        }
                    }

                    if (sCount2013 == 1)
                    {
                        if (lineWithSingleS2013 == 116)
                        {
                            ModifyExcelCell(worksheet, "AT" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2013 == 136)
                        {
                            ModifyExcelCell(worksheet, "AT" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2013 == 156)
                        {
                            ModifyExcelCell(worksheet, "AT" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2013 == 176)
                        {
                            ModifyExcelCell(worksheet, "AT" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2013 == 196)
                        {
                            ModifyExcelCell(worksheet, "AT" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AT" + (i + 2), "X");
                    }
                    int sCount2014 = 0;
                    int[] lineNumbersS2014 = { 117, 137, 157, 177, 197 };
                    int lineWithSingleS2014 = -1;

                    foreach (int lineNumber in lineNumbersS2014)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2014++;
                                lineWithSingleS2014 = lineNumber;
                            }
                        }
                    }

                    if (sCount2014 == 1)
                    {
                        if (lineWithSingleS2014 == 117)
                        {
                            ModifyExcelCell(worksheet, "AU" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2014 == 137)
                        {
                            ModifyExcelCell(worksheet, "AU" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2014 == 157)
                        {
                            ModifyExcelCell(worksheet, "AU" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2014 == 177)
                        {
                            ModifyExcelCell(worksheet, "AU" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2014 == 197)
                        {
                            ModifyExcelCell(worksheet, "AU" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AU" + (i + 2), "X");
                    }
                    int sCount2015 = 0;
                    int[] lineNumbersS2015 = { 118, 138, 158, 178, 198 };
                    int lineWithSingleS2015 = -1;

                    foreach (int lineNumber in lineNumbersS2015)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2015++;
                                lineWithSingleS2015 = lineNumber;
                            }
                        }
                    }

                    if (sCount2015 == 1)
                    {
                        if (lineWithSingleS2015 == 118)
                        {
                            ModifyExcelCell(worksheet, "AV" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2015 == 138)
                        {
                            ModifyExcelCell(worksheet, "AV" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2015 == 158)
                        {
                            ModifyExcelCell(worksheet, "AV" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2015 == 178)
                        {
                            ModifyExcelCell(worksheet, "AV" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2015 == 198)
                        {
                            ModifyExcelCell(worksheet, "AV" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AV" + (i + 2), "X");
                    }
                    int sCount2016 = 0;
                    int[] lineNumbersS2016 = { 119, 139, 159, 179, 199 };
                    int lineWithSingleS2016 = -1;

                    foreach (int lineNumber in lineNumbersS2016)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2016++;
                                lineWithSingleS2016 = lineNumber;
                            }
                        }
                    }

                    if (sCount2016 == 1)
                    {
                        if (lineWithSingleS2016 == 119)
                        {
                            ModifyExcelCell(worksheet, "AW" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2016 == 139)
                        {
                            ModifyExcelCell(worksheet, "AW" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2016 == 159)
                        {
                            ModifyExcelCell(worksheet, "AW" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2016 == 179)
                        {
                            ModifyExcelCell(worksheet, "AW" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2016 == 199)
                        {
                            ModifyExcelCell(worksheet, "AW" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AW" + (i + 2), "X");
                    }
                    int sCount2017 = 0;
                    int[] lineNumbersS2017 = { 120, 140, 160, 180, 200 };
                    int lineWithSingleS2017 = -1;

                    foreach (int lineNumber in lineNumbersS2017)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2017++;
                                lineWithSingleS2017 = lineNumber;
                            }
                        }
                    }

                    if (sCount2017 == 1)
                    {
                        if (lineWithSingleS2017 == 120)
                        {
                            ModifyExcelCell(worksheet, "AX" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2017 == 140)
                        {
                            ModifyExcelCell(worksheet, "AX" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2017 == 160)
                        {
                            ModifyExcelCell(worksheet, "AX" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2017 == 180)
                        {
                            ModifyExcelCell(worksheet, "AX" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2017 == 200)
                        {
                            ModifyExcelCell(worksheet, "AX" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AX" + (i + 2), "X");
                    }
                    int sCount2018 = 0;
                    int[] lineNumbersS2018 = { 121, 141, 161, 181, 201 };
                    int lineWithSingleS2018 = -1;

                    foreach (int lineNumber in lineNumbersS2018)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2018++;
                                lineWithSingleS2018 = lineNumber;
                            }
                        }
                    }

                    if (sCount2018 == 1)
                    {
                        if (lineWithSingleS2018 == 121)
                        {
                            ModifyExcelCell(worksheet, "AY" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2018 == 141)
                        {
                            ModifyExcelCell(worksheet, "AY" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2018 == 161)
                        {
                            ModifyExcelCell(worksheet, "AY" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2018 == 181)
                        {
                            ModifyExcelCell(worksheet, "AY" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2018 == 201)
                        {
                            ModifyExcelCell(worksheet, "AY" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AY" + (i + 2), "X");
                    }
                    int sCount2019 = 0;
                    int[] lineNumbersS2019 = { 122, 142, 162, 182, 202 };
                    int lineWithSingleS2019 = -1;

                    foreach (int lineNumber in lineNumbersS2019)
                    {
                        if (lines.Length >= lineNumber)
                        {
                            string line = lines[lineNumber - 1];
                            if (line.Contains("S"))
                            {
                                sCount2019++;
                                lineWithSingleS2019 = lineNumber;
                            }
                        }
                    }

                    if (sCount2019 == 1)
                    {
                        if (lineWithSingleS2019 == 122)
                        {
                            ModifyExcelCell(worksheet, "AZ" + (i + 2), "A");
                        }
                        else if (lineWithSingleS2019 == 142)
                        {
                            ModifyExcelCell(worksheet, "AZ" + (i + 2), "B");
                        }
                        else if (lineWithSingleS2019 == 162)
                        {
                            ModifyExcelCell(worksheet, "AZ" + (i + 2), "C");
                        }
                        else if (lineWithSingleS2019 == 182)
                        {
                            ModifyExcelCell(worksheet, "AZ" + (i + 2), "D");
                        }
                        else if (lineWithSingleS2019 == 202)
                        {
                            ModifyExcelCell(worksheet, "AZ" + (i + 2), "E");
                        }
                    }
                    else
                    {
                        ModifyExcelCell(worksheet, "AZ" + (i + 2), "X");
                    }


                }

                // Save and close the Excel
                workbook.Save();
                workbook.Close();

                // Quit Excel application
                excelApp.Quit();
            MessageBox.Show("              Processamento Concluído com Sucesso!              ");
        }

        private void ModifyExcelCell(Excel.Worksheet worksheet, string cell, string value)
        {
            Excel.Range range = worksheet.Range[cell];
            range.Value = value;

            // Modify the border
            //range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders.Weight = Excel.XlBorderWeight.xlThick;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            range = null;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
