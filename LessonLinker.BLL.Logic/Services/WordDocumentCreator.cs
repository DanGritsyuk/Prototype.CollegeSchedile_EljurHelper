using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse;
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
            int newWidthValue = AdjustCellWidth(originalCell, lessonItems.Count);

            originalCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new TableCellWidth { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa }
            ));

            var buffCell = (TableCell)originalCell.Clone();

            CreateLessonParagraph(element, lessNum, lessonItems[0].Name.ToUpper(), lessonItems[0].Teacher, lessonItems[0].Room, GetGroupNumberFormatted(lessonItems[0].GroupShort));

            for (int i = 1; i < lessonItems.Count; i++)
            {
                var newCell = (TableCell)buffCell.Clone();

                originalCell.Ancestors<TableRow>().FirstOrDefault()?.Append(newCell);
                CreateLessonParagraph(element, lessNum, lessonItems[i].Name.ToUpper(), lessonItems[i].Teacher, lessonItems[i].Room, GetGroupNumberFormatted(lessonItems[i].GroupShort));

                newCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
                    i < lessonItems.Count - 1 ? new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                    : new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 18 },
                    new TableCellWidth { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa }
                ));


            }
        }
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

    private int AdjustCellWidth(TableCell cell, int groupCount)
    {
        if (int.TryParse(cell.TableCellProperties?.GetFirstChild<TableCellWidth>().Width, out int originalWidth))
        {
            int newWidthValue = originalWidth / groupCount;
            cell.TableCellProperties.Append(new TableCellWidth { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa });
            return newWidthValue;
        }
        return originalWidth;
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
        string[] suffixes = { "One", "Two", "Three", "Four", "Five", "Six" };

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
    private void CreateEmptyFirstCell(OpenXmlElement element, string lessonMark)
    {
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessonMark));
        if (originalCell != null)
        {
            var newCell = (TableCell)originalCell.Clone();
            originalCell.RemoveAllChildren();
            originalCell.Append(new Paragraph(new Run(new Text("")))); // Пустая ячейка

            // Установка границ
            originalCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            ));
            newCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 18 }
            ));

            // Получение ширины оригинальной ячейки
            var originalWidth = originalCell.TableCellProperties?.GetFirstChild<TableCellWidth>();
            if (originalWidth != null)
            {
                // Извлечение значения ширины
                int originalWidthValue = int.Parse(originalWidth.Width);
                // Деление на 2 для новой ширины
                int newWidthValue = originalWidthValue / 2;

                // Установка новой ширины ячейки
                originalCell.TableCellProperties.Append(
                        new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa }
                    );
                newCell.TableCellProperties.Append(
                        new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa }
                    );
            }

            originalCell.Ancestors<TableRow>().FirstOrDefault().Append(newCell);
        }
    }

    // Метод для создания пустой второй ячейки
    private void CreateEmptySecondCell(OpenXmlElement element, string lessonMark)
    {
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessonMark));
        if (originalCell != null)
        {
            var newCell = (TableCell)originalCell.Clone();
            newCell.RemoveAllChildren();
            newCell.Append(new Paragraph(new Run(new Text("")))); // Пустая ячейка

            // Установка границ
            originalCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 } // Последняя ячейка с границей 8
            ));
            newCell.TableCellProperties = new TableCellProperties(new TableCellBorders(
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 18 } // Последняя ячейка с границей 8
            ));

            // Получение ширины оригинальной ячейки
            var originalWidth = originalCell.TableCellProperties?.GetFirstChild<TableCellWidth>();
            if (originalWidth != null)
            {
                // Извлечение значения ширины
                int originalWidthValue = int.Parse(originalWidth.Width);
                // Деление на 2 для новой ширины
                int newWidthValue = originalWidthValue / 2;

                // Установка новой ширины ячейки
                originalCell.TableCellProperties.Append(
                        new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa }
                    );
                newCell.TableCellProperties.Append(
                        new TableCellWidth() { Width = newWidthValue.ToString(), Type = TableWidthUnitValues.Dxa }
                    );
            }

            var parentRow = originalCell.Ancestors<TableRow>().FirstOrDefault();
            parentRow?.Append(newCell);
        }
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