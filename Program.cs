using System.Drawing;
using System.Text.RegularExpressions;
using PdfiumViewer;
using PdfSharp.Pdf.IO;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Tesseract;

namespace PDFtoTEXT
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Запуск в режиме Telegram-бота:
            await RunTelegramBotAsync();

            // Если требуется режим обработки папки, закомментируйте строку выше и раскомментируйте строку ниже:
            // RunFolderProcessing();
        }

        public static async Task RunTelegramBotAsync()
        {
            // Токен, полученный от BotFather (храните его безопасно!)
            var botClient = new TelegramBotClient("варлдоваопрвжыарповарпоыврапорывапрвыапоываполдрываролдп");
            var me = await botClient.GetMeAsync();
            Console.WriteLine($"Бот запущен: {me.FirstName} (ID: {me.Id})");

            using CancellationTokenSource cts = new CancellationTokenSource();

            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = Array.Empty<UpdateType>()
            };

            botClient.StartReceiving(
                updateHandler: HandleUpdateAsync,
                errorHandler: HandleErrorAsync,
                receiverOptions: receiverOptions,
                cancellationToken: cts.Token
            );

            Console.WriteLine("Бот запущен и ждёт сообщений. Нажмите любую клавишу для завершения...");
            Console.ReadKey();
            cts.Cancel();
        }

        static async Task HandleUpdateAsync(ITelegramBotClient bot, Update update, CancellationToken cancellationToken)
        {
            if (update.Type == UpdateType.Message && update.Message != null)
            {
                Message message = update.Message;
                if (message.Document != null && message.Document.MimeType == "application/pdf")
                {
                    try
                    {
                        // Получаем информацию о файле
                        var file = await bot.GetFileAsync(message.Document.FileId, cancellationToken);
                        string telegramFilePath = file.FilePath;
                        string localFilePath = Path.Combine(Path.GetTempPath(), message.Document.FileName);

                        // Формируем URL для скачивания файла:
                        string botToken = "варлдоваопрвжыарповарпоыврапорывапрвыапоываполдрываролдп";
                        string fileUrl = $"https://api.telegram.org/file/bot{botToken}/{telegramFilePath}";

                        using (var httpClient = new HttpClient())
                        using (var response = await httpClient.GetAsync(fileUrl, cancellationToken))
                        {
                            response.EnsureSuccessStatusCode();
                            using (var stream = new FileStream(localFilePath, FileMode.Create))
                            {
                                await response.Content.CopyToAsync(stream, cancellationToken);
                            }
                        }
                        Console.WriteLine($"Файл сохранён: {localFilePath}");

                        string tessdataPath = AppDomain.CurrentDomain.BaseDirectory;
                        // Отправляем страницы по мере их обработки
                        await ExtractNamesFromPdfAndSendAsync(localFilePath, tessdataPath, bot, message.Chat.Id, cancellationToken);

                        await bot.SendTextMessageAsync(message.Chat.Id, "PDF документ успешно обработан!", cancellationToken: cancellationToken);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка обработки PDF: {ex.Message}");
                        await bot.SendTextMessageAsync(message.Chat.Id, $"Ошибка обработки PDF: {ex.Message}", cancellationToken: cancellationToken);
                    }
                }
                else
                {
                    await bot.SendTextMessageAsync(message.Chat.Id, "Пожалуйста, отправьте PDF документ.", cancellationToken: cancellationToken);
                }
            }
        }

        static Task HandleErrorAsync(ITelegramBotClient bot, Exception exception, CancellationToken cancellationToken)
        {
            Console.WriteLine($"Ошибка: {exception.Message}");
            return Task.CompletedTask;
        }

        public static void RunFolderProcessing()
        {
            Console.WriteLine("Программа запущена в режиме обработки папки.");
            try
            {
                string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string pdfFolderPath = Path.Combine(projectDirectory, "PDFDocuments");
                string tessdataPath = projectDirectory;
                string[] pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);

                Console.WriteLine($"Найдено {pdfFiles.Length} PDF файлов.");
                foreach (string pdfFile in pdfFiles)
                {
                    Console.WriteLine($"\nНачинаем обработку файла: {pdfFile}");
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

        // Новый метод: обрабатывает PDF и отправляет каждую страницу сразу после обработки
        public static async Task ExtractNamesFromPdfAndSendAsync(string pdfPath, string tessdataPath, ITelegramBotClient bot, long chatId, CancellationToken cancellationToken)
        {
            string outputFolder = Path.Combine(Path.GetDirectoryName(pdfPath), "Pages");
            Directory.CreateDirectory(outputFolder);
            try
            {
                using (var pdfViewerDocument = PdfDocument.Load(pdfPath))
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
                                string name = ExtractName(pageText);
                                Console.Out.Flush();

                                string savedFilePath = SavePageAsPdf(inputDoc, i, name, outputFolder);
                                Console.WriteLine($"Страница {i + 1}: {name}");

                                using var fileStream = File.OpenRead(savedFilePath);
                                await bot.SendDocumentAsync(
                                    chatId: chatId,
                                    document: new InputFileStream(fileStream, Path.GetFileName(savedFilePath)),
                                    cancellationToken: cancellationToken
                                );
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

        // Оригинальный метод для обработки PDF (для режима папки)
        public static List<string> ExtractNamesFromPdf(string pdfPath, string tessdataPath)
        {
            List<string> processedFiles = new List<string>();
            string outputFolder = Path.Combine(Path.GetDirectoryName(pdfPath), "Pages");
            Directory.CreateDirectory(outputFolder);
            try
            {
                using (var pdfViewerDocument = PdfDocument.Load(pdfPath))
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
                                string name = ExtractName(pageText);
                                Console.Out.Flush();
                                string savedFilePath = SavePageAsPdf(inputDoc, i, name, outputFolder);
                                processedFiles.Add(savedFilePath);
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
            return processedFiles;
        }

        public static Bitmap RenderPageToBitmap(PdfDocument document, int pageIndex)
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

        public static string PerformOcr(Bitmap image, string tessdataPath)
        {
            try
            {
                using (var engine = new TesseractEngine(tessdataPath, "rus", EngineMode.Default))
                {
                    image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                    string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                    image.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);
                    using (var pix = Pix.LoadFromFile(tempFile))
                    {
                        using (var page = engine.Process(pix, PageSegMode.SingleBlock))
                        {
                            string text = page.GetText();
                            File.Delete(tempFile);
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

        private const int AllowedDifferences = 3;
        private const string ExpectedMarker = "Настоящий сертификат удостоверяет, что";

        public static string ExtractName(string ocrText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ocrText))
                    return "Имя не найдено";

                string normalizedMarker = Normalize(ExpectedMarker);
                string normalizedOCR = Normalize(ocrText);
                int markerPos = FindMarkerPosition(normalizedOCR, normalizedMarker, AllowedDifferences);
                if (markerPos == -1)
                    return "Имя не найдено";

                int originalMarkerIndex = MapNormalizedIndexToOriginal(ocrText, markerPos);
                string textAfterMarker = ocrText.Substring(originalMarkerIndex).Replace("\r", " ").Replace("\n", " ");
                string cleanedText = Regex.Replace(textAfterMarker, @"[^\p{L}\s]", " ");
                cleanedText = Regex.Replace(cleanedText, @"\s+", " ").Trim();
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

        private static string Normalize(string text)
        {
            return Regex.Replace(text, @"[^\p{L}]", "").ToLower();
        }

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

        private static int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++)
                d[i, 0] = i;
            for (int j = 0; j <= m; j++)
                d[0, j] = j;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (s[i - 1] == t[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1,
                                 d[i, j - 1] + 1),
                                 d[i - 1, j - 1] + cost
                    );
                }
            }
            return d[n, m];
        }

        private static string SavePageAsPdf(PdfSharp.Pdf.PdfDocument inputDoc, int pageIndex, string name, string outputFolder)
        {
            PdfSharp.Pdf.PdfDocument outputDoc = new PdfSharp.Pdf.PdfDocument();
            outputDoc.AddPage(inputDoc.Pages[pageIndex]);
            string sanitizedName = SanitizeFileName(name);
            string fileName = $"{sanitizedName}.pdf";
            string filePath = Path.Combine(outputFolder, fileName);
            outputDoc.Save(filePath);
            return filePath;
        }

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
