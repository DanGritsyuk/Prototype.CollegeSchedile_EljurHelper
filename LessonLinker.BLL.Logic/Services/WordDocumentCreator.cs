using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse;
using LessonLinker.Common.Entities.Schedule;
using LessonLinker.Common.Entities.Extensions;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Collections.Generic;

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
                        // Копируем содержимое из шаблона
                        var sourceBody = templateDoc.MainDocumentPart.Document.Body.CloneNode(true);

                        // Заменяем <group> на 'Ф 11'
                        ReplaceText(sourceBody, "<Group>", response.GroupName!);

                        // Устанавливаем даты и занятия
                        SetDates(sourceBody, response);

                        // Очищаем неиспользованные маркеры после замены
                        ClearUnusedMarkers(sourceBody, "MonLess", "TuesLess", "WednesLess", "ThursLess", "FriLess", "SunLess");

                        // Добавляем скопированное содержимое в body
                        body.Append(sourceBody);

                        // Добавляем разделитель между документами (например, страница)
                        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                        Console.WriteLine($"Страница группы {response.GroupName} сформирована...");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                }
            }

            // После завершения цикла добавляем body в mainPart.Document
            mainPart.Document.Append(body);
            mainPart.Document.Save();
            Console.WriteLine($"Документ готов.");
        }
    }

    private void SetDates(OpenXmlElement element, ApiScheduleResponse response)
    {
        foreach (var day in response.Response.Result.Days)
        {
            // Преобразуем ключ (дату) из формата yyyyMMdd в формат dd.MM
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

    // Метод для преобразования строки с датой из формата yyyyMMdd в формат dd.MM
    private string FormatDate(string date)
    {
        if (DateTime.TryParseExact(date, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
        {
            // Возвращаем только день и месяц в формате dd.MM
            return parsedDate.ToString("dd.MM");
        }
        // Если дата не в нужном формате, возвращаем пустую строку или оригинальную строку
        return date;
    }

    private void SetLessons(OpenXmlElement element, string dayMark, Day day)
    {
        foreach (var lessonGroup in day.Items.GroupBy(l => l.Number))
        {
            var lessonItems = lessonGroup.OrderBy(l => GetGroupNumber(l.GroupShort)).ToList();
            var firstLesson = lessonItems.First();

            // Если группа не указана (весь класс идет на урок), создаем одну ячейку
            if (string.IsNullOrEmpty(firstLesson.GroupShort))
            {
                string lessNum = dayMark + GetSuffixForLessonNumber(firstLesson.Number);
                CreateLessonParagraph(element, lessNum, firstLesson.Name.ToUpper(), firstLesson.Teacher, firstLesson.Room, ""); // Не указываем "ГР"
            }
            else
            {
                // Если указана только одна группа
                if (lessonItems.Count == 1)
                {
                    string groupNum = GetGroupNumberFormatted(firstLesson.GroupShort);

                    // Если группа ГР1, помещаем в первую ячейку, вторую оставляем пустой
                    if (groupNum == "ГР1")
                    {
                        CreateSingleGroupCell(element, dayMark, firstLesson, groupNum, true);
                    }
                    else
                    {
                        // Если группа не ГР1, первую ячейку оставляем пустой, а вторая содержит урок
                        CreateSingleGroupCell(element, dayMark, firstLesson, groupNum, false);
                    }
                }
                else
                {
                    // Несколько групп: уменьшаем ширину первой ячейки и создаем дополнительные
                    string lessNum = dayMark + GetSuffixForLessonNumber(firstLesson.Number);
                    CreateLessonParagraph(element, lessNum, firstLesson.Name.ToUpper(), firstLesson.Teacher, firstLesson.Room, GetGroupNumberFormatted(firstLesson.GroupShort));

                    AdjustCellWidth(element, lessNum, lessonItems.Count);

                    // Для каждой следующей группы создаем новую ячейку
                    for (int i = 1; i < lessonItems.Count; i++)
                    {
                        string groupNum = GetGroupNumberFormatted(lessonItems[i].GroupShort);
                        CreateAdditionalGroupCell(element, lessNum, groupNum, lessonItems[i], i == lessonItems.Count - 1);
                    }
                }
            }
        }
    }

    // Метод для создания одной группы с выбором, пустая ли первая или вторая ячейка
    private void CreateSingleGroupCell(OpenXmlElement element, string dayMark, LessonItem lesson, string groupNum, bool firstCell)
    {
        string lessNum = dayMark + GetSuffixForLessonNumber(lesson.Number);

        if (firstCell)
        {
            // ГР1, помещаем в первую ячейку, вторую оставляем пустой
            CreateLessonParagraph(element, lessNum, lesson.Name.ToUpper(), lesson.Teacher, lesson.Room, groupNum);
            CreateEmptySecondCell(element, lessNum);
        }
        else
        {
            // Не ГР1, первую ячейку оставляем пустой, вторую заполняем
            CreateEmptyFirstCell(element, lessNum);
            CreateLessonParagraphForGroup(element, lessNum + "G2", lesson.Name.ToUpper(), lesson.Teacher, lesson.Room);
        }
    }

    // Унификация и извлечение номера группы
    private int GetGroupNumber(string group)
    {
        if (string.IsNullOrEmpty(group)) return 0; // Если группа не указана, считаем, что группа одна, но без номера
        var match = Regex.Match(group.ToUpper(), @"\d+");
        return match.Success ? int.Parse(match.Value) : 0; // Если нашли номер группы, возвращаем его, иначе 0
    }

    // Приведение группы к общему формату ГРN
    private string GetGroupNumberFormatted(string group)
    {
        int groupNum = GetGroupNumber(group);
        return groupNum > 0 ? "ГР" + groupNum : ""; // Возвращаем ГР1, ГР2 и т.д. или пустую строку, если группа не указана
    }

    // Метод для создания пустой первой ячейки
    private void CreateEmptyFirstCell(OpenXmlElement element, string lessonMark)
    {
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessonMark));
        if (originalCell != null)
        {
            var newCell = (TableCell)originalCell.Clone();
            newCell.RemoveAllChildren();
            newCell.Append(new Paragraph(new Run(new Text("")))); // Пустая ячейка
            var parentRow = originalCell.Ancestors<TableRow>().FirstOrDefault();
            parentRow?.Append(newCell);
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
            var parentRow = originalCell.Ancestors<TableRow>().FirstOrDefault();
            parentRow?.Append(newCell);
        }
    }

    // Метод для создания дополнительных ячеек для последующих групп
    private void CreateAdditionalGroupCell(OpenXmlElement element, string lessonMark, string groupNum, LessonItem lesson, bool isLast)
    {
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessonMark));
        if (originalCell != null)
        {
            var newCell = (TableCell)originalCell.Clone();
            newCell.RemoveAllChildren();

            CreateLessonParagraph(newCell, lessonMark, lesson.Name.ToUpper(), lesson.Teacher, lesson.Room, groupNum);

            var parentRow = originalCell.Ancestors<TableRow>().FirstOrDefault();
            parentRow?.Append(newCell);

            // Устанавливаем границы для новой ячейки
            newCell.TableCellProperties.Append(new TableCellBorders(
                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = isLast ? new UInt32Value((uint)8) : 4 } // Последняя ячейка с границей 8
            ));
        }
    }

    // Метод для уменьшения ширины первой ячейки
    private void AdjustCellWidth(OpenXmlElement element, string lessonMark, int groupCount)
    {
        var originalCell = element.Descendants<TableCell>().FirstOrDefault(tc => tc.InnerText.Contains(lessonMark));
        if (originalCell != null)
        {
            double newWidth = 500 / groupCount; // Уменьшаем ширину ячейки, делим на количество групп
            originalCell.TableCellProperties = new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (newWidth * 20).ToString() });
        }
    }

    // Генерация суффикса для номера урока
    private string GetSuffixForLessonNumber(string number)
    {
        return number switch
        {
            "1" => "One",
            "2" => "Two",
            "3" => "Three",
            "4" => "Four",
            "5" => "Five",
            "6" => "Six",
            _ => "Unknown"
        };
    }

    // Дополнительный метод для создания параграфа для второй группы
    private void CreateLessonParagraphForGroup(OpenXmlElement element, string lessonMark, string lessonName, string teacherInitials, string room)
    {
        foreach (var textElement in element.Descendants<Text>())
        {
            if (textElement.Text.Contains(lessonMark))
            {
                // Создаем новый Run для lessonName (жирный и курсив)
                var lessonNameRun = new Run(
                    new RunProperties(
                        new Bold(),                    // Жирный текст
                        new Italic(),                  // Курсив
                        new FontSize() { Val = "18" }  // Устанавливаем шрифт 9
                    ),
                    new Text(lessonName)
                );

                // Проверяем, нужно ли добавлять префикс "КАБ"
                string roomText = room != null && IsRoomValid(room) ? $" КАБ {room}" : $" {room}";

                // Создаем новый Run для teacherInitials + комната (жирный и курсив)
                var teacherRun = new Run(
                    new RunProperties(
                        new Bold(),                    // Жирный текст
                        new Italic(),                  // Курсив
                        new FontSize() { Val = "18" }  // Шрифт 9
                    ),
                    new Break() { Type = BreakValues.TextWrapping },   // Переход на новую строку (Shift+Enter)
                    new Text($"{teacherInitials}{roomText}")
                );

                // Заменяем текст в элементе на два новых Run
                textElement.Parent.InsertBeforeSelf(lessonNameRun);
                textElement.Parent.InsertBeforeSelf(teacherRun);

                textElement.Remove(); // Удаляем старый текст
            }
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

    // Метод для удаления неиспользованных маркеров с суффиксами
    private void ClearUnusedMarkers(OpenXmlElement element, params string[] dayMarks)
    {
        // Суффиксы для каждого маркера от One до Six
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
                        // Очищаем текст, если маркер найден
                        textElement.Text = textElement.Text.Replace(fullMarker, string.Empty);
                    }
                }
            }
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

    // Метод для создания форматированного абзаца с переходом на новую строку (Shift+Enter) и жирностью + курсивом
    private void CreateLessonParagraph(OpenXmlElement element, string lessonMark, string lessonName, string teacherInitials, string room, string group)
    {
        foreach (var textElement in element.Descendants<Text>())
        {
            if (textElement.Text.Contains(lessonMark))
            {
                // Формируем текст для группы, добавляя метку группы (например, ГР1)
                string lessonText = $"{lessonName} {group}";

                // Создаем новый Run для lessonName (жирный и курсив)
                var lessonRun = new Run(
                    new RunProperties(new Bold(), new Italic(), new FontSize() { Val = "18" }),
                    new Text(lessonText)
                );

                // Создаем новый Run для teacherInitials + комната (жирный и курсив)
                string roomText = room != null && IsRoomValid(room) ? $" КАБ {room}" : $" {room}";
                var teacherRun = new Run(
                    new RunProperties(new Bold(), new Italic(), new FontSize() { Val = "18" }),
                    new Break() { Type = BreakValues.TextWrapping },
                    new Text($"{teacherInitials}{roomText}")
                );

                // Заменяем текст в элементе на два новых Run
                textElement.Parent.InsertBeforeSelf(lessonRun);
                textElement.Parent.InsertBeforeSelf(teacherRun);

                textElement.Remove(); // Удаляем старый текст
            }
        }
    }

    // Метод для проверки, является ли строка допустимым номером кабинета (состоит только из цифр или цифр с дефисом)
    private bool IsRoomValid(string room)
    {
        // Регулярное выражение для проверки: строка содержит только цифры или цифры с дефисом
        return Regex.IsMatch(room, @"^\d+(-\d+)?$");
    }
}
