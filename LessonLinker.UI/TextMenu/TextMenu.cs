using LessonLinker.UI.TextMenu.MenuRendering;

namespace LessonLinker.UI.TextMenu
{
    internal class TextMenu
    {
        private readonly int CONSOLE_LINES = 12;
        private bool runMenu;

        public TextMenu() =>
            OpenMenu();

        public int YesNoCancelDialog(string message)
        {
            var menuData = new Dictionary<string, List<string>>
            {
                { message, new List<string> { "Да", "Нет", "Отмена " } }
            };

            return DrawDialogMenu(menuData, 0);
        }

        public bool YesNoDialog(string message)
        {
            var menuData = new Dictionary<string, List<string>>
            {
                { message, new List<string> { "Да", "Нет " } }
            };

            int commandKey = DrawDialogMenu(menuData, 1);
            return commandKey == 1;
        }

        public int ContentMenu(ICollection<string> members, string headerText)
        {
            var menuData = new Dictionary<string, List<string>>();
            List<string> namesOnPage = new List<string>();
            int pageNumber = 0;

            foreach (var name in members)
            {
                namesOnPage.Add(name);
                if (namesOnPage.Count >= CONSOLE_LINES - 3 ||
                    namesOnPage.Count + pageNumber * 10 == members.Count)
                {
                    menuData.Add(headerText + " " + ++pageNumber + ":", namesOnPage);
                    namesOnPage = new List<string>();
                }
            }

            return DrawContentMenu(menuData, 1) - 1;
        }

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action fEnd) =>
            SelectionMenu(menuData, f1, null, null, null, null, null, null, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, null, null, null, null, null, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, null, null, null, null, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action f4, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, null, null, null, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action f4, Action f5,
                Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, f5, null, null, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action f4, Action f5,
                Action f6, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, f5, f6, null, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action f4, Action f5,
                Action f6, Action f7, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, f5, f6, f7, null, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action f4, Action f5,
                Action f6, Action f7, Action f8, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, f5, f6, f7, f8, null, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
                Action f1, Action f2, Action f3, Action f4, Action f5,
                Action f6, Action f7, Action f8, Action f9, Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, f5, f6, f7, f8, f9, null, fEnd);

        public void RunMenuWithActions(IDictionary<string, List<string>> menuData,
        Action? f1, Action? f2, Action? f3, Action? f4, Action? f5,
        Action? f6, Action? f7, Action? f8, Action? f9, Action? f10,
        Action fEnd) =>
            SelectionMenu(menuData, f1, f2, f3, f4, f5, f6, f7, f8, f9, f9, fEnd);

        private void SelectionMenu(IDictionary<string, List<string>> menuData,
                Action? f1, Action? f2, Action? f3, Action? f4, Action? f5,
                Action? f6, Action? f7, Action? f8, Action? f9, Action? f10,
                Action fEnd)
        {
            int commandKey = 1;

            Func<Action?, bool> runFunc = (Action? func) =>
            {
                if (func != null)
                    func();
                else
                    fEnd();
                return true;
            };

            OpenMenu();
            while (runMenu)
            {
                commandKey = DrawContentMenu(menuData, commandKey == 0 ? 1 : commandKey);
                switch (commandKey)
                {
                    case 1:
                        runFunc(f1);
                        break;
                    case 2:
                        runFunc(f2);
                        break;
                    case 3:
                        runFunc(f3);
                        break;
                    case 4:
                        runFunc(f4);
                        break;
                    case 5:
                        runFunc(f5);
                        break;
                    case 6:
                        runFunc(f6);
                        break;
                    case 7:
                        runFunc(f7);
                        break;
                    case 8:
                        runFunc(f8);
                        break;
                    case 9:
                        runFunc(f9);
                        break;
                    case 10:
                        runFunc(f10);
                        break;
                    default:
                        fEnd();
                        break;
                }
            }
        }

        public void OpenMenu()
        {
            this.runMenu = true;
        }

        public void CloseMenu()
        {
            this.runMenu = false;
        }

        private int DrawContentMenu(IDictionary<string, List<string>> menuData, int taskId)
        {
            var menu = new MenuRender(menuData, CONSOLE_LINES, true, true, "", "", "> ", "");
            return menu.StartRenderMenu(taskId - 1);
        }

        private int DrawDialogMenu(IDictionary<string, List<string>> menuData, int taskIndex)
        {
            var menu = new MenuRender(menuData, 0, false, false, "", "", "", "");
            return menu.StartRenderMenu(taskIndex);
        }
    }
}
