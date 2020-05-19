using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using Microsoft.Win32;
using NLog;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace MNPD_Documents_Recognition
{
    /// <summary>
    /// Логика взаимодействия для MNPD_Recognition.xaml
    /// </summary>
    public partial class MNPD_Recognition : Window
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private string filepath;
        BackgroundWorker worker;

        public MNPD_Recognition()
        {
            InitializeComponent();
            worker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            worker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
        }

        #region Функции для работы с элементами интерфейса

        /// <summary>
        /// Отработка выбора страницы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Pages_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string imagesDirectory = "cache";
            string currentDirectory = Directory.GetCurrentDirectory();
            var testImagePath = Path.Combine(currentDirectory, imagesDirectory, Pages.SelectedItem.ToString());
            PagePicture.Source = new BitmapImage(new Uri(testImagePath));

            //подрубаем тессеракт
            var imageFile = File.ReadAllBytes(testImagePath);
            logger.Debug(currentDirectory + @"\tesseract");
            var text = ParseText(currentDirectory + @"\tesseract", imageFile, "result_1000+Cyrillic");
            logger.Debug(text);
            RecognizedText.Clear();
            Regex pattern = new Regex("[ћђ]");
            for (int i = 0; i < text.Length; i++)
            {
                RecognizedText.Text += pattern.Replace(text[i], "ѣ") + Environment.NewLine;
            }
        }

        /// <summary>
        /// Обработка клика по кнопке выбора PDF файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Choose_PDF_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (dialog.ShowDialog() == true)
            {
                filepath = dialog.FileName;
                PagePicture.Source = null;
            }
            if (worker.IsBusy != true)
            {
                worker.RunWorkerAsync();
            }
                
        }

        #endregion

        #region Функции для работы с обработчиком background операций 
        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                var perc = e.ProgressPercentage;
                PDF_Convertion_Status.Value = e.ProgressPercentage;
            }
            catch (Exception msg)
            {
                logger.Debug(msg.ToString());
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                string imagesDirectory = "cache";
                string currentDirectory = Directory.GetCurrentDirectory();

                //заполняю листбокс
                string[] fileArray = Directory.GetFiles(Path.Combine(currentDirectory, imagesDirectory)).Select(Path.GetFileName).ToArray();
                var result = fileArray.OrderBy(x => x.Length);
                Pages.ItemsSource = result.ToArray();

                var testImagePath = Path.Combine(currentDirectory, imagesDirectory, "0.png");
                var imageFile = File.ReadAllBytes(testImagePath);
                logger.Debug(currentDirectory + @"\tesseract");
                var text = ParseText(currentDirectory + @"\tesseract", imageFile, "result_1000+Cyrillic");
                logger.Debug(text);
                for (int i = 0; i < text.Length; i++)
                {
                    RecognizedText.Text += text[i] + Environment.NewLine;
                }
            }
            else
            {
                var mess = e.Error;
                logger.Info(mess);
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                logger.Debug($"Поддержка отмены: {worker.WorkerSupportsCancellation}");
                logger.Debug($"Отмена воркера: {worker.CancellationPending}");
                using (var library = DocLib.Instance)
                {
                    using (var docReader = library.GetDocReader(filepath, new PageDimensions(768, 1024)))
                    {
                        string imagesDirectory = "cache";
                        Directory.CreateDirectory(imagesDirectory);
                        string currentDirectory = Directory.GetCurrentDirectory();
                        logger.Debug($"Количество страниц: {docReader.GetPageCount()}");
                        var pagesCount = docReader.GetPageCount();
                        for (int i = 0; i < pagesCount; i++)
                        {
                            using (var pageReader = docReader.GetPageReader(i))
                            {
                                var bytes = GetModifiedImage(pageReader);
                                File.WriteAllBytes(Path.Combine(currentDirectory, imagesDirectory, $"{i}.png"), bytes);
                            }
                            int progressPercentage = Convert.ToInt32(((double)i / pagesCount) * 100);
                            worker.ReportProgress(progressPercentage);
                        }
                        worker.CancelAsync();
                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                        }
                    }
                }

            }
            catch (Exception msg)
            {
                logger.Debug(msg.ToString());
            }
        }


        #endregion

        #region Различные вспомогательные функции
        /// <summary>
        /// Запускаем tesseract напрямую через exe файл, возвращаем распознанный текст
        /// TODO: Переделать уже с запуска из exe файла используя API (это оказалось очень очень сложно)
        /// </summary>
        /// <param name="tesseractPath"></param>
        /// <param name="imageFile"></param>
        /// <param name="lang"></param>
        /// <returns></returns>
        private static string[] ParseText(string tesseractPath, byte[] imageFile, params string[] lang)
        {
            string[] output;
            var tempOutputFile = Path.GetTempPath() + Guid.NewGuid();
            var tempImageFile = Path.GetTempFileName();

            try
            {
                File.WriteAllBytes(tempImageFile, imageFile);

                ProcessStartInfo info = new ProcessStartInfo();
                info.WorkingDirectory = tesseractPath;
                //info.WindowStyle = ProcessWindowStyle.Hidden;
                info.UseShellExecute = false;
                info.FileName = "cmd.exe";
                info.CreateNoWindow = true;
                info.Arguments =
                    "/c tesseract.exe " +
                    // Файл изображения
                    tempImageFile + " " +
                    // Временный выходной файл
                    tempOutputFile +
                    // Добавляем язык
                    " -l " + string.Join("+", lang);

                // Запускаем процесс Тессеракта
                Process process = Process.Start(info);
                process.WaitForExit();
                if (process.ExitCode == 0)
                    output = File.ReadAllLines(tempOutputFile + ".txt", Encoding.UTF8);
                else
                    throw new Exception("Error. Tesseract stopped with an error code = " + process.ExitCode);
            }
            finally
            {
                File.Delete(tempImageFile);
            }

            return output;
        }

        private byte[] GetModifiedImage(IPageReader pageReader)
        {
            var rawBytes = pageReader.GetImage();
            var width = pageReader.GetPageWidth();
            var height = pageReader.GetPageHeight();

            using (var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
            {
                AddBytes(bmp, rawBytes);
                Rectangle section = new Rectangle(new System.Drawing.Point(12, 50), new System.Drawing.Size(578, 876));
                Bitmap CroppedImage = CropImage(bmp, section);
                using (var stream = new MemoryStream())
                {
                    CroppedImage.Save(stream, ImageFormat.Png);
                    return stream.ToArray();
                }
            }
        }

        private void AddBytes(Bitmap bmp, byte[] rawBytes)
        {
            var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            var bmpData = bmp.LockBits(rect, ImageLockMode.WriteOnly, bmp.PixelFormat);
            var pNative = bmpData.Scan0;
            Marshal.Copy(rawBytes, 0, pNative, rawBytes.Length);
            bmp.UnlockBits(bmpData);
        }

        public Bitmap CropImage(Bitmap source, Rectangle section)
        {
            var bitmap = new Bitmap(section.Width, section.Height);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel);
                return bitmap;
            }
        }

        //SetGrayscale
        public Bitmap SetGrayscale(Bitmap img)
        {

            Bitmap temp = (Bitmap)img;
            Bitmap bmap = (Bitmap)temp.Clone();
            System.Drawing.Color c;
            for (int i = 0; i < bmap.Width; i++)
            {
                for (int j = 0; j < bmap.Height; j++)
                {
                    c = bmap.GetPixel(i, j);
                    byte gray = (byte)(.299 * c.R + .587 * c.G + .114 * c.B);

                    bmap.SetPixel(i, j, Color.FromArgb(gray, gray, gray));
                }
            }
            return (Bitmap)bmap.Clone();

        }
        //RemoveNoise
        public Bitmap RemoveNoise(Bitmap bmap)
        {

            for (var x = 0; x < bmap.Width; x++)
            {
                for (var y = 0; y < bmap.Height; y++)
                {
                    var pixel = bmap.GetPixel(x, y);
                    if (pixel.R < 162 && pixel.G < 162 && pixel.B < 162)
                        bmap.SetPixel(x, y, Color.Black);
                    else if (pixel.R > 162 && pixel.G > 162 && pixel.B > 162)
                        bmap.SetPixel(x, y, Color.White);
                }
            }

            return bmap;
        }
        #endregion
    }
}
