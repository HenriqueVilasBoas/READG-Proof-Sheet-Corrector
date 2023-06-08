using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using AForge.Imaging;
using AForge.Imaging.Filters;
using AForge.Math.Geometry;
using static System.Net.Mime.MediaTypeNames;


namespace ReadG
{
    public partial class Form2 : Form
    {
        // Variável para armazenar as coordenadas detectadas

        private System.Drawing.Point detectedCoordinates;


        public Form2()
        {
            InitializeComponent();
            listBox1.KeyDown += listBox1_KeyDown;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Abrir o explorador de arquivos para selecionar as imagens
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Imagens JPEG|*.jpg";
            openFileDialog.Title = "Selecione as imagens";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string imagePath in openFileDialog.FileNames)
                {
                    // Carregar a imagem selecionada
                    Bitmap image = new Bitmap(imagePath);

                    for (int y = 0; y < image.Height; y++)
                    {
                        for (int x = 0; x < image.Width; x++)
                        {
                            Color pixelColor = image.GetPixel(x, y);

                            if (pixelColor.R < 50 && pixelColor.G < 50 && pixelColor.B < 200)
                            {
                                image.SetPixel(x, y, Color.Black);
                            }
                        }
                    }

                    Grayscale grayscaleFilter = new Grayscale(0.90, 0.90, 0.8);
                    Bitmap grayImage = grayscaleFilter.Apply(image);

                    // Definir as coordenadas e dimensões da matriz
                    int startX = 167;
                    int startY = 786; 
                    int cellWidth = 58;
                    int cellHeight = 46;
                    int numRows = 20;
                    int numCols = 5;

                    // Desenhar a matriz vermelha com células
                    using (Graphics graphics = Graphics.FromImage(image))
                    {
                        using (Pen pen = new Pen(Color.Red, 1))
                        {
                            for (int row = 0; row < numRows; row++)
                            {
                                for (int col = 0; col < numCols; col++)
                                {
                                    int x = startX + col * cellWidth;
                                    int y = startY + row * cellHeight;
                                    Rectangle cellRect = new Rectangle(x, y, cellWidth, cellHeight);
                                    graphics.DrawRectangle(pen, cellRect);
                                }
                            }
                        }
                    }

                    // Salvar a imagem automaticamente em "C:\Users\Leonel\3D Objects"
                    string savePath = Path.Combine(@"C:\Users\Leonel\3D Objects", Path.GetFileName(imagePath));
                    image.Save(savePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                MessageBox.Show("Imagens processadas e salvas com sucesso.");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Percorrer os itens da listBox1
            foreach (var item in listBox1.Items)
            {
                string imagePath = item.ToString();

                // Definir as coordenadas de busca
                int startX = 60;
                int startY = 80;
                int endX = 205;
                int endY = 308;

                using (Bitmap originalImage = new Bitmap(imagePath))
                {
                    // Percorrer as coordenadas definidas
                    for (int y = startY; y <= endY; y++)
                    {
                        for (int x = startX; x <= endX; x++)
                        {
                            // Obter o valor dos componentes de cor
                            Color pixelColor = originalImage.GetPixel(x, y);
                            int red = pixelColor.R;
                            int green = pixelColor.G;
                            int blue = pixelColor.B;

                            // Verificar se os componentes de cor estão próximos de preto
                            if (red < 150 && green < 140 && blue < 140)
                            {
                                // Definir o pixel como preto
                                originalImage.SetPixel(x, y, Color.Black);
                            }
                        }
                    }

                    // Percorrer novamente as coordenadas definidas para encontrar as coordenadas desejadas
                    for (int y = startY; y <= endY; y++)
                    {
                        for (int x = startX; x <= endX; x++)
                        {
                            // Obter o valor dos componentes de cor
                            Color pixelColor = originalImage.GetPixel(x, y);
                            int red = pixelColor.R;
                            int green = pixelColor.G;
                            int blue = pixelColor.B;

                            // Verificar se os componentes de cor estão dentro do intervalo desejado
                            if (red <= 1 && green <= 1 && blue <= 1)
                            {
                                // Armazenar as coordenadas detectadas na variável
                                Point detectedCoordinates = new Point(x, y);

                                // Verificar se as coordenadas detectadas são diferentes de 75 e 141
                                if (detectedCoordinates.X != 750 || detectedCoordinates.Y != 1410)
                                {
                                    using (Bitmap croppedImage = new Bitmap(originalImage.Width, originalImage.Height))
                                    using (Graphics graphics = Graphics.FromImage(croppedImage))
                                    {
                                        // Copiar a parte desejada da imagem original
                                        Rectangle sourceRect = new Rectangle(detectedCoordinates.X, detectedCoordinates.Y, originalImage.Width, originalImage.Height);
                                        Rectangle destRect = new Rectangle(0, 0, originalImage.Width, originalImage.Height);
                                        graphics.DrawImage(originalImage, destRect, sourceRect, GraphicsUnit.Pixel);

                                        // Salvar a nova imagem no local especificado com o mesmo nome da imagem original
                                        string savePath = Path.Combine(@"C:\READG\IMAGEM\", Path.GetFileName(imagePath));
                                        croppedImage.Save(savePath, System.Drawing.Imaging.ImageFormat.Jpeg);

                                        //originalImage.Dispose();
                                    }
                                }

                                // Parar a verificação
                                break;
                            }
                        }
                    }
                }
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Arquivos de Imagem (*.jpg)|*.jpg";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    listBox1.Items.Add(fileName);
                }
            }
        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                _ = listBox1.SelectedItem.ToString();
                // Do something with the selected file
            }
        }
        private void listBox1_KeyDown(object sender, KeyEventArgs e)
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
    }
}