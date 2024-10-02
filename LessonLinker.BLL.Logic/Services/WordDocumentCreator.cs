using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse;
using System.Linq;
using System.Text.RegularExpressions;

public class DocumentProcessor
{
    public async Task CreateAndMergeDocuments(string templatePath, string outputPath, IAsyncEnumerable<ApiScheduleResponse> schedule)
    {
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException("Шаблон не найден.", templatePath);
        }

        using (WordprocessingDocument mergedDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = mergedDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();

            await foreach (ApiScheduleResponse response in schedule)
            {
                try
                {
                    Console.WriteLine($"Расписание для группы {response.GroupName} получено...");
                    using (WordprocessingDocument templateDoc = WordprocessingDocument.Open(templatePath, false))
                    {
                        var sourceBody = templateDoc.MainDocumentPart.Document.Body.CloneNode(true);

                        ReplaceText(sourceBody, "<Group>", response.GroupName!);
                        SetDates(sourceBody, response);
                        ClearUnusedMarkers(sourceBody, "MonLess", "TuesLess", "WednesLess", "ThursLess", "FriLess", "SunLess");

                        body.Append(sourceBody);
                        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                        Console.WriteLine($"Страница группы {response.GroupName} сформирована...");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                }
            }

            mainPart.Document.Append(body);
            mainPart.Document.Save();
            Console.WriteLine("Документ готов.");
        }
    }

    private void SetDates(OpenXmlElement element, ApiScheduleResponse response)
    {
        foreach (var day in response.Response.Result.Days)
        {
            var formattedDate = FormatDate(day.Key);

            switch (day.Value.Title)
            {
                case "Понедельник":
                    ReplaceText(element, "MonDate", formattedDate);
                    SetLessons(element, "MonLess", day.Value);
                    break;
                case "Вторник":
                    ReplaceText(element, "TuesDate", formattedDate);
                    SetLessons(element, "TuesLess", day.Value);
                    break;
                case "Среда":
                    ReplaceText(element, "WednesDate", formattedDate);
                    SetLessons(element, "WednesLess", day.Value);
                    break;
                case "Четверг":
                    ReplaceText(element, "ThursDate", formattedDate);
                    SetLessons(element, "ThursLess", day.Value);
                    break;
                case "Пятница":
                    ReplaceText(element, "FriDate", formattedDate);
                    SetLessons(element, "FriLess", day.Value);
                    break;
                case "Суббота":
                    ReplaceText(element, "SunDate", formattedDate);
                    SetLessons(element, "SunLess", day.Value);
                    break;
                default:
                    break;
            }
        }
    }

    private string FormatDate(string date)
    {
        if (DateTime.TryParseExact(date, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
        {
            return parsedDate.ToString("dd.MM");
        }
        return date;
    }

    private void SetLessons(OpenXmlElement element, string dayMark, Day day)
    {
        if (day.Items is null) return;

        foreach (var lessonGroup in day.Items.GroupBy(l => l.Number))
        {
            var lessonItems = lessonGroup.OrderBy(l => GetGroupNumber(l.GroupShort)).ToList();
            var firstLesson = lessonItems.First();

            if (string.IsNullOrEmpty(firstLesson.GroupShort))
            {
                string lessNum = dayMark + GetSuffixForLessonNumber(firstLesson.Number);
                CreateLessonParagraph(element, lessNum, firstLesson.Name.ToUpper(), firstLesson.Teacher, firstLesson.Room, "");
            }
            else
            {
                if (lessonItems.Count == 1)
                {
                    string groupNum = GetGroupNumberFormatted(firstLesson.GroupShort);
                    CreateSingleGroupCell(element, dayMark, firstLesson, groupNum, groupNum == "ГР1");
                }
                else
                {
                    CreateMultipleGroupCells(element, dayMark, lessonItems);
                }
            }
        }
    }

    private void CreateMultipleGroupCells(OpenXmlElement element, string dayMark, List<LessonItem> lessonItems)
    {
        string lessNum = dayMark + GetSuffixForLessonNumber(lessonItems.First().Number);
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessNum));

        if (originalCell != null)
        {
            // Получаем значения topBorder и bottomBorder у originalCell
            uint topBorder = 4;
            uint bottomBorder = 4;

            if (originalCell.TableCellProperties?.TableCellBorders != null)
            {
                topBorder = originalCell.TableCellProperties.TableCellBorders.TopBorder?.Size ?? 4;
                bottomBorder = originalCell.TableCellProperties.TableCellBorders.BottomBorder?.Size ?? 4;
            }

            // Рассчитываем новую ширину ячейки
            int newWidthValue = AdjustCellWidth(originalCell, lessonItems.Count);

            // Применяем границы и устанавливаем ширину для оригинальной ячейки
            ApplyBorders(originalCell, topBorder, bottomBorder, 4, newWidthValue);

            // Клонируем ячейку для дальнейшего использования
            var buffCell = (TableCell)originalCell.Clone();

            // Заполняем первую ячейку
            CreateLessonParagraph(element, lessNum, lessonItems[0].Name.ToUpper(), lessonItems[0].Teacher, lessonItems[0].Room, GetGroupNumberFormatted(lessonItems[0].GroupShort));

            // Создаем остальные ячейки для каждой следующей группы
            for (int i = 1; i < lessonItems.Count; i++)
            {
                var newCell = (TableCell)buffCell.Clone();

                // Вставляем новую ячейку в строку
                var parentRow = originalCell.Ancestors<TableRow>().FirstOrDefault();
                parentRow?.Append(newCell);

                // Заполняем ячейку данными о следующей группе
                CreateLessonParagraph(newCell, lessNum, lessonItems[i].Name.ToUpper(), lessonItems[i].Teacher, lessonItems[i].Room, GetGroupNumberFormatted(lessonItems[i].GroupShort));

                // Применяем границы и ширину к новой ячейке
                ApplyBorders(newCell, topBorder, bottomBorder, (i == lessonItems.Count - 1) ? (uint)18 : 4, newWidthValue);
            }
        }
    }


    /// Метод для корректировки ширины всех ячеек в строке после их добавления
    private void AdjustRowCellWidths(TableRow row, int groupCount)
    {
        if (row != null)
        {
            var cells = row.Elements<TableCell>().ToList();

            // Рассчитываем новую ширину для каждой ячейки
            int totalWidth = cells.Sum(cell => int.Parse(cell.TableCellProperties.GetFirstChild<TableCellWidth>().Width));
            int newWidth = --totalWidth / groupCount;

            // Применяем новую ширину для каждой ячейки
            foreach (var cell in cells)
            {
                var cellWidth = cell.TableCellProperties.GetFirstChild<TableCellWidth>();
                if (cellWidth != null)
                {
                    cellWidth.Width = newWidth.ToString();
                }
            }
        }
    }

    // Метод для изменения ширины ячеек
    private int AdjustCellWidth(TableCell originalCell, int groupCount)
    {
        var originalWidth = originalCell.TableCellProperties?.GetFirstChild<TableCellWidth>();
        if (originalWidth != null)
        {
            // Извлечение значения ширины
            int originalWidthValue = int.Parse(originalWidth.Width);
            // Деление на количество групп
            return originalWidthValue / groupCount;
        }
        return 0;
    }

    // Применение границ и ширины к ячейке
    private void ApplyBorders(TableCell cell, uint topBorder, uint bottomBorder, uint rightBorderSize, int newWidthValue)
    {
        cell.TableCellProperties = new TableCellProperties(new TableCellBorders(
            new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = topBorder },
            new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = bottomBorder },
            new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = rightBorderSize }
        ));

        // Устанавливаем новую ширину ячейки
        cell.TableCellProperties.Append(new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa });
    }


    private void CreateSingleGroupCell(OpenXmlElement element, string dayMark, LessonItem lesson, string groupNum, bool firstCell)
    {
        string lessNum = dayMark + GetSuffixForLessonNumber(lesson.Number);

        if (firstCell)
        {
            CreateEmptySecondCell(element, lessNum);
            CreateLessonParagraph(element, lessNum, lesson.Name.ToUpper(), lesson.Teacher, lesson.Room, groupNum);
        }
        else
        {
            CreateEmptyFirstCell(element, lessNum);
            CreateLessonParagraph(element, lessNum, lesson.Name.ToUpper(), lesson.Teacher, lesson.Room, groupNum);
        }
    }

    private void ReplaceText(OpenXmlElement element, string oldValue, string newValue)
    {
        foreach (var text in element.Descendants<Text>())
        {
            if (text.Text.Contains(oldValue))
            {
                text.Text = text.Text.Replace(oldValue, newValue);
            }
        }
    }

    private void ClearUnusedMarkers(OpenXmlElement element, params string[] dayMarks)
    {
        string[] suffixes = { "One", "Two", "Three", "Four", "Five", "Six", "Seven" };

        foreach (var textElement in element.Descendants<Text>())
        {
            foreach (var dayMark in dayMarks)
            {
                foreach (var suffix in suffixes)
                {
                    string fullMarker = dayMark + suffix;
                    if (textElement.Text.Contains(fullMarker))
                    {
                        textElement.Text = textElement.Text.Replace(fullMarker, string.Empty);
                    }
                }
            }
        }
    }

    private string GetSuffixForLessonNumber(string number) => number switch
    {
        "1" => "One",
        "2" => "Two",
        "3" => "Three",
        "4" => "Four",
        "5" => "Five",
        "6" => "Six",
        "7" => "Seven",
        _ => "Unknown"
    };

    private int GetGroupNumber(string group) => int.TryParse(Regex.Match(group?.ToUpper() ?? string.Empty, @"\d+").Value, out int num) ? num : 0;

    private string GetGroupNumberFormatted(string group) => $"ГР{GetGroupNumber(group)}";

    private bool IsRoomValid(string room) => Regex.IsMatch(room, @"^\d+(-\d+)?$");

    private void CreateLessonParagraph(OpenXmlElement element, string lessonMark, string lessonName, string teacherInitials, string room, string group)
    {
        foreach (var textElement in element.Descendants<Text>())
        {
            if (textElement.Text.Contains(lessonMark))
            {
                var lessonText = $"{lessonName} {group}";
                var roomText = IsRoomValid(room) ? $" КАБ {room}" : $" {room}";
                teacherInitials = GetTeacherInitials(teacherInitials);

                textElement.Parent.InsertBeforeSelf(new Run(new RunProperties(new Bold(), new Italic(), new FontSize { Val = "18" }), new Text(lessonText)));
                textElement.Parent.InsertBeforeSelf(new Run(new RunProperties(new Bold(), new Italic(), new FontSize { Val = "18" }), new Break { Type = BreakValues.TextWrapping }, new Text($"{teacherInitials.ToUpper()}{roomText}")));

                textElement.Remove();
            }
        }
    }

    // Метод для создания пустой первой ячейки
    private void CreateEmptyCell(OpenXmlElement element, string lessonMark, bool isFirstCell)
    {
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessonMark));

        uint topBorder = 4;
        uint bottomBorder = 4;

        // Получение границ оригинальной ячейки
        if (originalCell?.TableCellProperties?.TableCellBorders != null)
        {
            topBorder = originalCell.TableCellProperties.TableCellBorders.TopBorder?.Size ?? 4;
            bottomBorder = originalCell.TableCellProperties.TableCellBorders.BottomBorder?.Size ?? 4;
        }

        if (originalCell != null)
        {
            var newCell = (TableCell)originalCell.Clone();
            newCell.RemoveAllChildren();
            newCell.Append(new Paragraph(new Run(new Text("")))); // Пустая ячейка

            // Установка границ
            ApplyBorders(originalCell, newCell, topBorder, bottomBorder, isFirstCell);

            // Получение ширины оригинальной ячейки и изменение ширины
            AdjustCellWidths(originalCell, newCell);

            var parentRow = originalCell.Ancestors<TableRow>().FirstOrDefault();
            parentRow?.Append(newCell);
        }
    }

    // Применение границ к ячейкам
    private void ApplyBorders(TableCell originalCell, TableCell newCell, uint topBorder, uint bottomBorder, bool isFirstCell)
    {
        originalCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
            new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = topBorder },
            new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = bottomBorder },
            new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
        ));
        newCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
            new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = topBorder },
            new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = bottomBorder },
            new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 18 }
        ));
    }

    // Изменение ширины ячеек
    private void AdjustCellWidths(TableCell originalCell, TableCell newCell)
    {
        var originalWidth = originalCell.TableCellProperties?.GetFirstChild<TableCellWidth>();
        if (originalWidth != null)
        {
            int originalWidthValue = int.Parse(originalWidth.Width);
            int newWidthValue = originalWidthValue / 2;

            originalCell.TableCellProperties.Append(new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa });
            newCell.TableCellProperties.Append(new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa });
        }
    }

    // Метод для создания пустой первой ячейки
    private void CreateEmptyFirstCell(OpenXmlElement element, string lessonMark)
    {
        CreateEmptyCell(element, lessonMark, true);
    }

    // Метод для создания пустой второй ячейки
    private void CreateEmptySecondCell(OpenXmlElement element, string lessonMark)
    {
        CreateEmptyCell(element, lessonMark, false);
    }

    // Преобразование полного ФИО в "Фамилия И.О."
    private string GetTeacherInitials(string fullName)
    {
        if (string.IsNullOrWhiteSpace(fullName))
            return string.Empty;

        var parts = fullName.Split(' ');
        if (parts.Length < 2)
            return fullName; // Если имя не в формате ФИО, возвращаем как есть

        string initials = $"{parts[0]} {parts[1][0]}.";
        if (parts.Length > 2)
        {
            initials += $" {parts[2][0]}.";
        }

        return initials;
    }
}