using System.Drawing;
using System.Text.RegularExpressions;
using PdfiumViewer;
using PdfSharp.Pdf.IO;
using Tesseract;

namespace PDFtoTEXT
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Программа запущена.");
            try
            {
                string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
                Console.WriteLine($"Каталог проекта: {projectDirectory}");

                // Папка с PDF-документами
                string pdfFolderPath = Path.Combine(projectDirectory, "PDFDocuments");
                Console.WriteLine($"Папка с PDF-документами: {pdfFolderPath}");

                // Папка с tessdata (ожидается, что файлы, например rus.traineddata, лежат в каталоге проекта)
                string tessdataPath = projectDirectory;

                string[] pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);
                Console.WriteLine($"Найдено {pdfFiles.Length} PDF файлов.");

                foreach (string pdfFile in pdfFiles)
                {
                    Console.WriteLine();
                    Console.WriteLine($"Начинаем обработку файла: {pdfFile}");
                    try
                    {
                        ExtractNamesFromPdf(pdfFile, tessdataPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при обработке файла {pdfFile}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Критическая ошибка: {ex.Message}");
            }

            Console.WriteLine("Обработка завершена. Нажмите любую клавишу для выхода...");
            Console.ReadKey();
        }

        /// <summary>
        /// Обрабатывает PDF: для каждой страницы выполняется OCR, извлекается имя,
        /// выводится в консоль, а затем соответствующая страница сохраняется в отдельный PDF-файл.
        /// </summary>
        public static void ExtractNamesFromPdf(string pdfPath, string tessdataPath)
        {

            string outputFolder = Path.Combine(Path.GetDirectoryName(pdfPath), "Pages");
            Directory.CreateDirectory(outputFolder);
            try
            {
                // Загружаем PDF для рендеринга (OCR) с помощью PdfiumViewer
                using (var pdfViewerDocument = PdfDocument.Load(pdfPath))
                // Открываем исходный PDF для извлечения страниц с помощью PdfSharp
                using (var inputDoc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
                {
                    Console.WriteLine($"Количество страниц: {pdfViewerDocument.PageCount}");
                    for (int i = 0; i < pdfViewerDocument.PageCount; i++)
                    {
                        Console.WriteLine($"Обработка страницы {i + 1}");
                        try
                        {
                            using (var bitmap = RenderPageToBitmap(pdfViewerDocument, i))
                            {
                                string pageText = PerformOcr(bitmap, tessdataPath);
                                // Выводим первые 50 символов OCR-текста для отладки
                                string name = ExtractName(pageText);
                                Console.Out.Flush();

                                // Извлекаем и сохраняем страницу в отдельный PDF-файл
                                SavePageAsPdf(inputDoc, i, name, outputFolder);
                                Console.WriteLine($"Страница {i + 1}: {name}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ошибка на странице {i + 1}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при открытии или обработке PDF файла: {ex.Message}");
            }
        }

        /// <summary>
        /// Рендерит указанную страницу PDF в Bitmap с разрешением 300 DPI.
        /// </summary>
        public static Bitmap RenderPageToBitmap(PdfiumViewer.PdfDocument document, int pageIndex)
        {
            try
            {
                const int dpi = 300;
                var size = document.PageSizes[pageIndex];
                int width = (int)(size.Width * dpi / 72);
                int height = (int)(size.Height * dpi / 72);
                Bitmap bitmap = (Bitmap)document.Render(pageIndex, width, height, dpi, dpi, PdfRenderFlags.Annotations);
                return bitmap;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при рендеринге страницы {pageIndex + 1}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Выполняет распознавание текста с изображения с помощью Tesseract.
        /// Перед OCR изображение поворачивается на 90 градусов.
        /// </summary>
        public static string PerformOcr(Bitmap image, string tessdataPath)
        {
            try
            {
                using (var engine = new TesseractEngine(tessdataPath, "rus", EngineMode.Default))
                {
                    // Поворачиваем изображение, если оно перевёрнуто на 90 градусов
                    image.RotateFlip(RotateFlipType.Rotate90FlipNone);

                    // Сохраняем изображение во временный файл
                    string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                    image.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);

                    using (var pix = Pix.LoadFromFile(tempFile))
                    {
                        using (var page = engine.Process(pix, PageSegMode.SingleBlock))
                        {
                            string text = page.GetText();
                            File.Delete(tempFile); // Удаляем временный файл
                            return text;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при выполнении OCR: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Извлекает имя и фамилию из OCR-текста.
        /// Предполагается, что после фразы "Настоящий сертификат удостоверяет, что" идёт текст с именем.
        /// </summary>
        // Порог допустимых различий (можно настроить)
        private const int AllowedDifferences = 3;
        // Ожидаемый маркер
        private const string ExpectedMarker = "Настоящий сертификат удостоверяет, что";

        public static string ExtractName(string ocrText)
        {

            try
            {
                if (string.IsNullOrWhiteSpace(ocrText))
                    return "Имя не найдено";

                // Нормализуем ожидаемый маркер: удаляем всё, кроме букв, и приводим к нижнему регистру
                string normalizedMarker = Normalize(ExpectedMarker);

                // Нормализуем весь OCR-текст аналогичным способом
                string normalizedOCR = Normalize(ocrText);

                // Пытаемся найти в нормализованном тексте подстроку, похожую на маркер
                int markerPos = FindMarkerPosition(normalizedOCR, normalizedMarker, AllowedDifferences);
                if (markerPos == -1)
                    return "Имя не найдено";

                // Определяем позицию в исходном тексте по количеству букв до markerPos
                int originalMarkerIndex = MapNormalizedIndexToOriginal(ocrText, markerPos);

                // Извлекаем текст после найденного маркера и приводим к одной строке
                string textAfterMarker = ocrText.Substring(originalMarkerIndex);
                textAfterMarker = textAfterMarker.Replace("\r", " ").Replace("\n", " ");

                // Оставляем в тексте только буквы и пробелы
                string cleanedText = Regex.Replace(textAfterMarker, @"[^\p{L}\s]", " ");
                cleanedText = Regex.Replace(cleanedText, @"\s+", " ").Trim();

                // Шаблон для имени: минимум два слова, каждое начинается с заглавной буквы и содержит минимум 3 буквы
                string pattern = @"\b([А-ЯЁ][а-яё]{2,}(?:\s+[А-ЯЁ][а-яё]{2,})+)\b";
                var match = Regex.Match(cleanedText, pattern);
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
                return "Имя не найдено";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при извлечении имени: {ex.Message}");
                return "Имя не найдено";
            }
        }

        // Нормализация: удаляем всё, кроме букв, и приводим к нижнему регистру
        private static string Normalize(string text)
        {
            return Regex.Replace(text, @"[^\p{L}]", "").ToLower();
        }

        // Преобразуем позицию в нормализованном тексте в позицию в исходном тексте, считая только буквы.
        private static int MapNormalizedIndexToOriginal(string originalText, int normIndex)
        {
            int letterCount = 0;
            for (int i = 0; i < originalText.Length; i++)
            {
                if (char.IsLetter(originalText[i]))
                {
                    if (letterCount == normIndex)
                        return i;
                    letterCount++;
                }
            }
            return 0;
        }

        // Поиск позиции маркера в нормализованном тексте с учетом допустимых отличий
        private static int FindMarkerPosition(string normalizedOCR, string normalizedMarker, int allowedDifferences)
        {
            int markerLength = normalizedMarker.Length;
            for (int i = 0; i <= normalizedOCR.Length - markerLength; i++)
            {
                string window = normalizedOCR.Substring(i, markerLength);
                int distance = LevenshteinDistance(window, normalizedMarker);
                if (distance <= allowedDifferences)
                {
                    return i;
                }
            }
            return -1;
        }

        // Классическая реализация алгоритма Левенштейна
        private static int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // Инициализация
            for (int i = 0; i <= n; i++)
                d[i, 0] = i;
            for (int j = 0; j <= m; j++)
                d[0, j] = j;

            // Вычисление расстояния
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (s[i - 1] == t[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1,      // удаление
                                 d[i, j - 1] + 1),     // вставка
                                 d[i - 1, j - 1] + cost // замена
                    );
                }
            }
            return d[n, m];
        }

        /// <summary>
        /// Сохраняет исходную страницу PDF в отдельный PDF-файл.
        /// Используется PdfSharp для извлечения страницы в оригинальном формате.
        /// </summary>
        private static void SavePageAsPdf(PdfSharp.Pdf.PdfDocument inputDoc, int pageIndex, string name, string outputFolder)
        {
            // Создаём новый PDF-документ и добавляем в него выбранную страницу
            PdfSharp.Pdf.PdfDocument outputDoc = new PdfSharp.Pdf.PdfDocument();
            outputDoc.AddPage(inputDoc.Pages[pageIndex]);

            // Очищаем имя от недопустимых символов для имени файла
            string sanitizedName = SanitizeFileName(name);
            // Формируем имя файла: номер страницы + "_" + имя + ".pdf"
            string fileName = $"{sanitizedName}.pdf";
            string filePath = Path.Combine(outputFolder, fileName);
            outputDoc.Save(filePath);
        }

        /// <summary>
        /// Удаляет из строки символы, недопустимые для имени файла.
        /// </summary>
        private static string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            return fileName;
        }
    }
}
