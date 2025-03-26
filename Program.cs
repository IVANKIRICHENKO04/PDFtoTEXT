using System.Drawing;
using System.Text.RegularExpressions;
using PdfiumViewer;
using Tesseract;

namespace PDFtoTEXT
{
    class Program
    {
        static void Main(string[] args)
        {
            string pdfFilePath = @"C:\Users\ivan3\Desktop\data\1.pdf";
            string tessdataPath = @"C:\Users\ivan3\Desktop\data"; // Папка должна содержать rus.traineddata

            ExtractNamesFromPdf(pdfFilePath, tessdataPath);
            Console.WriteLine("Обработка завершена.");
        }

        /// <summary>
        /// Обрабатывает PDF, извлекает OCR-текст каждой страницы, ищет имя и выводит номер страницы и имя.
        /// </summary>
        public static void ExtractNamesFromPdf(string pdfPath, string tessdataPath)
        {
            using (var document = PdfDocument.Load(pdfPath))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var bitmap = RenderPageToBitmap(document, i))
                    {
                        string pageText = PerformOcr(bitmap, tessdataPath);
                        string name = ExtractName(pageText);
                        Console.WriteLine($"Страница {i} : {name}");
                        Console.Out.Flush();
                    }
                }
            }
        }

        /// <summary>
        /// Рендерит указанную страницу PDF в Bitmap с разрешением 300 DPI.
        /// </summary>
        public static Bitmap RenderPageToBitmap(PdfDocument document, int pageIndex)
        {
            const int dpi = 300;
            var size = document.PageSizes[pageIndex];
            int width = (int)(size.Width * dpi / 72);
            int height = (int)(size.Height * dpi / 72);
            Bitmap bitmap = (Bitmap)document.Render(pageIndex, width, height, dpi, dpi, PdfRenderFlags.Annotations);
            return bitmap;
        }

        /// <summary>
        /// Выполняет распознавание текста с изображения с помощью Tesseract.
        /// Перед OCR изображение поворачивается на 90 градусов.
        /// </summary>
        public static string PerformOcr(Bitmap image, string tessdataPath)
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

        /// <summary>
        /// Извлекает имя и фамилию из OCR-текста.
        /// Предполагается, что после фразы "Настоящий сертификат удостоверяет, что" может идти до трех пустых строк, а затем 2–3 слова с заглавной буквы.
        /// </summary>
        // Порог допустимых различий (можно настроить)
        private const int AllowedDifferences = 2;

        // Ожидаемый маркер
        private const string ExpectedMarker = "Настоящий сертификат удостоверяет, что";

        public static string ExtractName(string ocrText)
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

            // Ищем имя – минимум два слова, каждое начинается с заглавной буквы, затем строчные
            string pattern = @"\b([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)+)\b";
            Match match = Regex.Match(cleanedText, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return "Имя не найдено";
        }

        // Нормализация: удаляем всё, кроме букв, и приводим к нижнему регистру
        private static string Normalize(string text)
        {
            return Regex.Replace(text, @"[^\p{L}]", "").ToLower();
        }

        // Преобразуем позицию в нормализованном тексте в позицию в исходном тексте,
        // считая только буквы.
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
    }
}
