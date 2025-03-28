using PDFiumCore;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using System.Text.RegularExpressions;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Tesseract;

class Program
{
    private static readonly string BotToken = GetBotToken();

    private static string GetBotToken()
    {
            return "7986993779:AAHfJmEMVfsfDaRx22n0YH2ew5nLL9JwFjE";
    }

    static async Task Main(string[] args)
    {
        fpdfview.FPDF_InitLibrary();
        await RunTelegramBotAsync();
        fpdfview.FPDF_DestroyLibrary();
    }

    public static async Task RunTelegramBotAsync()
    {
        var botClient = new TelegramBotClient(BotToken);
        using var cts = new CancellationTokenSource();

        var receiverOptions = new ReceiverOptions
        {
            AllowedUpdates = new[] { UpdateType.Message }
        };

        botClient.StartReceiving(HandleUpdateAsync, HandleErrorAsync, receiverOptions, cts.Token);

        Console.WriteLine("🟢 Бот успешно запущен и ожидает сообщений...");

        AppDomain.CurrentDomain.ProcessExit += (_, _) => cts.Cancel();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cts.Cancel();
        };

        try
        {
            await Task.Delay(Timeout.Infinite, cts.Token);
        }
        catch (TaskCanceledException) { }

        Console.WriteLine("🔴 Бот остановлен.");
    }

    static async Task HandleUpdateAsync(ITelegramBotClient bot, Update update, CancellationToken token)
    {
        var chatId = update.Message?.Chat.Id;
        var userName = update.Message?.From?.Username ?? "неизвестный пользователь";

        if (update.Message?.Document?.MimeType == "application/pdf")
        {
            Console.WriteLine($"📥 PDF получен от {userName}. Начинается обработка.");

            try
            {
                var file = await bot.GetFileAsync(update.Message.Document.FileId, token);
                var localPath = Path.Combine(Path.GetTempPath(), update.Message.Document.FileName!);

                var fileUrl = $"https://api.telegram.org/file/bot{BotToken}/{file.FilePath}";
                using (var httpClient = new HttpClient())
                await using (var fs = new FileStream(localPath, FileMode.Create))
                {
                    var fileStream = await httpClient.GetStreamAsync(fileUrl, token);
                    await fileStream.CopyToAsync(fs, token);
                }

                Console.WriteLine($"💾 PDF скачан: {localPath}");

                await ExtractAndSendPagesAsync(localPath, bot, chatId!.Value, token);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Ошибка обработки PDF: {ex.Message}");
                await bot.SendTextMessageAsync(chatId!, $"⚠️ Ошибка обработки PDF: {ex.Message}", cancellationToken: token);
            }
        }
        else
        {
            Console.WriteLine($"⚠️ Получено не PDF от {userName}.");
            await bot.SendTextMessageAsync(chatId!, "⚠️ Пожалуйста, отправьте PDF файл.", cancellationToken: token);
        }
    }

    static Task HandleErrorAsync(ITelegramBotClient bot, Exception exception, CancellationToken token)
    {
        Console.WriteLine($"❌ Ошибка: {exception.Message}");
        return Task.CompletedTask;
    }

    static async Task ExtractAndSendPagesAsync(string pdfPath, ITelegramBotClient bot, long chatId, CancellationToken token)
    {
        Console.WriteLine($"🛠️ Начало обработки файла: {pdfPath}");
        var tessdataPath = AppDomain.CurrentDomain.BaseDirectory;
        using var inputDoc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
        var pdfDoc = fpdfview.FPDF_LoadDocument(pdfPath, null);

        int pageCount = fpdfview.FPDF_GetPageCount(pdfDoc);
        Console.WriteLine($"📄 Обнаружено страниц: {pageCount}");

        for (int i = 0; i < pageCount; i++)
        {
            try
            {
                Console.WriteLine($"🔍 Обработка страницы {i + 1}/{pageCount}...");
                using var bitmap = RenderPage(pdfDoc, i);
                var text = PerformOcr(bitmap, tessdataPath);
                var name = ExtractName(text);
                Console.WriteLine($"📝 Извлечённое имя: {name}");

                var pagePath = SavePageAsPdf(inputDoc, i, name);
                Console.WriteLine($"📤 Отправка страницы: {pagePath}");

                await using var fs = File.OpenRead(pagePath);
                await bot.SendDocumentAsync(chatId, new InputFileStream(fs, Path.GetFileName(pagePath)), cancellationToken: token);

                Console.WriteLine($"✅ Страница {i + 1} отправлена.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Ошибка на странице {i + 1}: {ex.Message}");
                await bot.SendTextMessageAsync(chatId, $"⚠️ Ошибка на странице {i + 1}: {ex.Message}", cancellationToken: token);
            }
        }

        fpdfview.FPDF_CloseDocument(pdfDoc);
        await bot.SendTextMessageAsync(chatId, "✅ PDF документ успешно обработан!", cancellationToken: token);

        Console.WriteLine("✅ Обработка файла завершена.");
    }


    static SKBitmap RenderPage(FpdfDocumentT pdfDoc, int pageIndex)
    {
        Console.WriteLine($"🔍 Начало рендеринга страницы {pageIndex + 1}...");
        var page = fpdfview.FPDF_LoadPage(pdfDoc, pageIndex);
        int width = (int)(fpdfview.FPDF_GetPageWidth(page) * 3);
        int height = (int)(fpdfview.FPDF_GetPageHeight(page) * 3);

        Console.WriteLine($"🖼️ Размер страницы (до поворота): {width}x{height}");

        var bitmapHandle = fpdfview.FPDFBitmapCreate(width, height, 0);
        fpdfview.FPDFBitmapFillRect(bitmapHandle, 0, 0, width, height, 0xFFFFFFFF);
        fpdfview.FPDF_RenderPageBitmap(bitmapHandle, page, 0, 0, width, height, 0, 0);

        int stride = fpdfview.FPDFBitmapGetStride(bitmapHandle);
        IntPtr bufferPtr = fpdfview.FPDFBitmapGetBuffer(bitmapHandle);
        int bufferSize = stride * height;

        SKBitmap skBitmap = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);

        unsafe
        {
            Buffer.MemoryCopy((void*)bufferPtr, (void*)skBitmap.GetPixels(), bufferSize, bufferSize);
        }

        fpdfview.FPDFBitmapDestroy(bitmapHandle);
        fpdfview.FPDF_ClosePage(page);

        Console.WriteLine($"🔄 Поворот страницы на 90° по часовой стрелке...");

        // Поворачиваем изображение на 90 градусов по часовой стрелке
        SKBitmap rotatedBitmap = new SKBitmap(height, width);
        using (var canvas = new SKCanvas(rotatedBitmap))
        {
            canvas.Translate(rotatedBitmap.Width, 0);
            canvas.RotateDegrees(90);
            canvas.DrawBitmap(skBitmap, 0, 0);
        }

        skBitmap.Dispose();

        Console.WriteLine($"✅ Страница {pageIndex + 1} успешно отрендерена и повернута.");

        return rotatedBitmap;
    }

    static string PerformOcr(SKBitmap bitmap, string tessdataPath)
    {
        Console.WriteLine($"🖥️ Выполняется OCR страницы...");
        try
        {
            using var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
            using var img = Pix.LoadFromMemory(data.ToArray());

            using var engine = new TesseractEngine(tessdataPath, "rus", EngineMode.Default);
            using var page = engine.Process(img);

            string text = page.GetText();

            Console.WriteLine($"✅ OCR выполнен успешно, длина извлеченного текста: {text.Length} символов.");
            return text;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка OCR: {ex.Message}");
            throw;
        }
    }

    static string ExtractName(string ocrText)
    {
        Console.WriteLine("🔎 Извлечение имени из распознанного текста...");
        var regex = new Regex(@"Настоящий сертификат удостоверяет, что\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)");
        var match = regex.Match(ocrText.Replace("\n", " "));
        var result = match.Success ? match.Groups[1].Value : "Неизвестный";

        Console.WriteLine($"📝 Извлечено имя: {result}");
        return result;
    }

    static string SavePageAsPdf(PdfDocument inputDoc, int index, string name)
    {
        Console.WriteLine($"💾 Сохранение страницы {index + 1} как отдельный PDF...");
        var outputDoc = new PdfDocument();
        outputDoc.AddPage(inputDoc.Pages[index]);
        var sanitized = string.Concat(name.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
        var filePath = Path.Combine(Path.GetTempPath(), $"{sanitized}_{index + 1}.pdf");
        outputDoc.Save(filePath);

        Console.WriteLine($"✅ Страница сохранена: {filePath}");
        return filePath;
    }
}
