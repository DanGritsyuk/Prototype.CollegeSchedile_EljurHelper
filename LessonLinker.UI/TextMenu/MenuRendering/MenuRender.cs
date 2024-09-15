using LessonLinker.Common.Entities.Extensions;

namespace LessonLinker.UI.TextMenu.MenuRendering
{
    public class MenuRender
    {
        private const int HEADER_LINE_COUNT = 2;

        public string headerText;
        public string footerText;

        private int rowCursorPosition;
        private int colCursorPosition;
        private int consoleLines;
        private int largestLine;
        private int charMaxCountInLine;
        private bool isEscActive;
        private bool showHelpControl;
        private string prefix;
        private string prefixMark;
        private IDictionary<string, List<string>> menuData;
        private ISet<PageData> pagesMap;

        private PageData currentPage;

        public MenuRender(IDictionary<string, List<string>> menuData, int consoleLines,
                bool isEscActive, bool showHelpControl, string headerText, string footerText, string prefix,
                string prefixMark) :
            this(menuData, consoleLines, 190, isEscActive, showHelpControl, headerText, footerText, prefix, prefixMark)
        { }

        public MenuRender(IDictionary<string, List<string>> menuData, int consoleLines, int charMaxCountInLine,
                bool isEscActive, bool showHelpControl, string headerText, string footerText, string prefix,
                string prefixMark)
        {
            this.menuData = menuData;
            this.consoleLines = consoleLines;
            this.charMaxCountInLine = charMaxCountInLine;
            this.isEscActive = isEscActive;
            this.showHelpControl = showHelpControl;

            this.rowCursorPosition = Console.CursorTop;
            this.colCursorPosition = Console.CursorLeft;

            this.headerText = headerText == null ? "" : headerText;
            this.footerText = footerText == null ? "" : footerText;

            this.prefix = prefix == null || prefix.IsEmpty() ? "> " : prefix;
            this.prefixMark = prefixMark == null ? "" : prefixMark;
            this.largestLine = GetLargestLineLength();
            this.pagesMap = SplitDataToPages(this.menuData, this.consoleLines, HEADER_LINE_COUNT);
            this.currentPage = GetPageByLineIndex(0);
        }

        public int StartRenderMenu(int index)
        {
            try
            {
                currentPage = GetPageByLineIndex(index);
            }
            catch (Exception e)
            {
                ClearConsoleText();
                Console.WriteLine(e.Source);
                return -1;
            }

            int largestKey = GetLargestKeyTasks();
            this.consoleLines = largestKey > this.consoleLines ? largestKey + 1 : this.consoleLines;

            int pageCount = this.pagesMap.Count;

            DrawMenu();
            while (true)
            {
                var frameChanged = false;
                ConsoleKey key = Console.ReadKey(true).Key;
                switch (key)
                {
                    case ConsoleKey.Enter:
                        ClearConsoleText();
                        return currentPage.CurrentLineIndex + currentPage.StartLineIndex + 1;
                    case ConsoleKey.UpArrow:
                        if (currentPage.CurrentLineIndex > 0)
                        {
                            currentPage.CurrentLineIndex--;
                            GoCursorToStartPosition();
                            frameChanged = true;
                        }
                        break;
                    case ConsoleKey.DownArrow:
                        if (currentPage.CurrentLineIndex < currentPage.LinesCount - 1)
                        {
                            currentPage.CurrentLineIndex++;
                            GoCursorToStartPosition();
                            frameChanged = true;
                        }
                        break;
                    case ConsoleKey.LeftArrow:
                        if (currentPage.PageId > 1)
                        {
                            currentPage = GetNextPage(-1);
                            ClearConsoleText();
                            frameChanged = true;
                        }
                        break;
                    case ConsoleKey.RightArrow:
                        if (currentPage.PageId < pageCount)
                        {
                            currentPage = GetNextPage(1);
                            ClearConsoleText();
                            frameChanged = true;
                        }
                        break;
                    case ConsoleKey.Escape:
                        if (this.isEscActive)
                        {
                            ClearConsoleText();
                            return 0;
                        }
                        break;
                }
                if (frameChanged)
                    DrawMenu();
            }
        }

        private void DrawMenu()
        {
            int blockIdCount = 0;
            Console.WriteLine(this.headerText);

            string helpText = this.showHelpControl ? GetHelpText() : string.Empty;

            foreach (var chapter in currentPage.Data)
            {
                string chapterTitle = chapter.Key;
                List<string> chapterLines = chapter.Value;

                if (blockIdCount > 0)
                {
                    Console.WriteLine();
                }
                Console.WriteLine(chapter.Key);

                for (int i = 0; i < chapterLines.Count; i++)
                {
                    clearLine();
                    bool isSelected = i + blockIdCount == currentPage.CurrentLineIndex;
                    string prToConsole = string.Empty;
                    string lineText = chapterLines[i].Replace("\n", " ");

                    if (this.prefixMark.IsEmpty())
                        prToConsole = (isSelected ? this.prefix : " ".Repeat(this.prefix.Length)) + lineText;
                    else
                        prToConsole = isSelected ? lineText.Replace(this.prefixMark, this.prefix)
                                : lineText.Replace(this.prefixMark, " ".Repeat(this.prefixMark.Length));

                    if (prToConsole.Length > charMaxCountInLine)
                        prToConsole = prToConsole.Substring(0, charMaxCountInLine - 3) + "...";
                    else if (prToConsole.Length < charMaxCountInLine)
                        prToConsole += " ".Repeat(largestLine - prToConsole.Length);

                    if (isSelected)
                    {
                        Console.BackgroundColor = Console.ForegroundColor;
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.WriteLine(prToConsole);
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine(prToConsole);
                    }

                }
                blockIdCount += chapterLines.Count;
            }

            Console.WriteLine("\n".Repeat(GetEmptyLinesCount()));

            if (!helpText.IsEmpty())
                Console.WriteLine(helpText);

            if (!footerText.IsEmpty())
                Console.WriteLine(this.footerText);

        }

        private string GetHelpText()
        {
            int pagesCount = this.pagesMap.Count;

            string strPageNumbers = string.Empty;
            if (pagesCount > 1)
            {
                strPageNumbers = "-".Repeat(pagesCount);
                string strPageId = currentPage.PageId.ToString();
                for (int i = 0; i < strPageId.Length; i++)
                {
                    strPageNumbers = strPageNumbers.Substring(0, currentPage.PageId - 1 + i) + strPageId[i]
                            + strPageNumbers.Substring(currentPage.PageId + i);
                }
            }

            string indent = " ".Repeat(this.charMaxCountInLine / 4);
            if (pagesCount > 2)
            {
                if (currentPage.PageId > 1)
                {
                    strPageNumbers = "< " + strPageNumbers;
                }
                else
                {
                    strPageNumbers = "  " + strPageNumbers;
                }
                if (currentPage.PageId < pagesCount)
                {
                    strPageNumbers += " >";
                }
                else
                {
                    strPageNumbers += "  ";
                }
            }
            string pagesSwitchInfo = pagesCount > 1 ? "Стрелки право/лево - переключать страницы. " : "";
            string helpInfo = "Стрелки вверх/вниз - перемещаться между строками. " + pagesSwitchInfo
                    + "Enter - выбрать задачу. Для выхода нажмите Esc.";
            this.largestLine = helpInfo.Length > this.largestLine ? helpInfo.Length : this.largestLine;
            string pagesInfo = indent + strPageNumbers + "\n" + "=".Repeat(this.largestLine);

            string padding = pagesInfo + "\n" + helpInfo;

            return padding + "\n\n";
        }

        private int GetLargestKeyTasks()
        {
            int largestKeyTasks = 0;
            foreach (var chapter in this.menuData)
            {
                int lengthKeyTasks = chapter.Value.Count;
                largestKeyTasks = lengthKeyTasks > largestKeyTasks ? lengthKeyTasks : largestKeyTasks;
            }
            return largestKeyTasks + HEADER_LINE_COUNT + 1;
        }

        private PageData GetPageByLineIndex(int index)
        {
            foreach (var page in this.pagesMap)
            {
                if (index >= page.StartLineIndex && index < page.StartLineIndex + page.LinesCount)
                {
                    page.CurrentLineIndex = index - page.StartLineIndex;
                    return page;
                }
            }
            throw new Exception("Page not found!");
        }

        private int GetLargestLineLength()
        {
            int largestLineLength = 0;
            foreach (var chapter in this.menuData)
            {
                //if (largestLineLength < chapter.Key.Length)
                //{
                //    largestLineLength = chapter.Key.Length;
                //}

                if (chapter.Value.Count == 0)
                    continue;

                int maxLength = chapter.Value.Max(str => str.Length);
                if (largestLineLength < maxLength)
                    largestLineLength = maxLength;
            }

            int prLength = this.prefix.Length;
            int prMarkLength = this.prefixMark.Length;
            largestLineLength += (prLength > prMarkLength ? prLength : prMarkLength);
            return largestLineLength > charMaxCountInLine ? charMaxCountInLine
                    : largestLineLength;
        }

        private static ISet<PageData> SplitDataToPages(IDictionary<string, List<string>> menuData, int linesPage,
                int headerLineCount)
        {
            int pageLineCount = 0;
            int pageFirstIndex = 0;
            int pageId = 1;
            var pageData = new Dictionary<string, List<string>>();
            var pagesMap = new HashSet<PageData>();
            int i = 0;

            foreach (var chapter in menuData)
            {
                int countLinesKeyTasks = chapter.Value.Count();
                if (countLinesKeyTasks + headerLineCount + pageLineCount >= linesPage)
                {
                    PageData page = new PageData(pageId, pageFirstIndex, pageData);
                    pagesMap.Add(page);
                    pageLineCount = 0;
                    pageId += 1;
                    pageFirstIndex += page.LinesCount;
                    pageData.Clear();
                }
                pageLineCount += countLinesKeyTasks + headerLineCount;
                pageData.Add(chapter.Key, chapter.Value);
                if (i == menuData.Keys.Count - 1)
                {
                    pagesMap.Add(new PageData(pageId, pageFirstIndex, pageData));
                }
                i++;
            }
            return pagesMap;
        }

        private PageData GetNextPage(int step)
        {
            int currentIndex = currentPage.CurrentLineIndex;
            int pageId = currentPage.PageId;
            var page = pagesMap.FirstOrDefault(p => p.PageId == pageId + step);
            currentPage = page != null ? page : throw new Exception();
            currentPage.CurrentLineIndex = currentIndex < currentPage.LinesCount ? currentIndex : currentPage.LinesCount - 1;
            return currentPage;
        }

        private void ClearConsoleText()
        {
            GoCursorToStartPosition();
            for (int i = 0; i <= this.consoleLines; i++)
            {
                clearLine();
                Console.WriteLine();
            }
            GoCursorToStartPosition();
        }

        private void clearLine()
        {
            int rowCurrent = Console.CursorTop;
            Console.Write(" ".Repeat(this.largestLine));
            Console.SetCursorPosition(this.colCursorPosition, rowCurrent);
        }

        private void GoCursorToStartPosition() =>
            Console.SetCursorPosition(this.colCursorPosition, this.rowCursorPosition);

        private int GetEmptyLinesCount() =>
            this.consoleLines - currentPage.LinesCount - currentPage.Data.Count * HEADER_LINE_COUNT;
    }
}
