using LessonLinker.Common.Entities.Extensions;

namespace LessonLinker.UI.TextMenu
{
    public class TextMenuUI
    {
        private TextMenu menu;
        private IDictionary<int, string> data;
        private bool appIsRun;

        public TextMenuUI()
        {
            appIsRun = true;
            data = new Dictionary<int, string>();
            menu = new TextMenu();
        }


        public void Startup()
        {
            while (appIsRun)
            {
                ClsField();
                startupMenu();
            }
        }

        public void ShowMessage(string text)
        {
            Console.WriteLine(text);
            Console.WriteLine("      Нажмите любую клавишу для продолжения...");
            Console.ReadKey();
            ClsField();
        }

        private void ClsField()
        {
            string header = "ПРОГРАММА: Интерфейс консольного меню. Версия 0.000.00 pre-Alpha"
                    + " ".Repeat(32);
            Console.Clear();
            Console.WriteLine(header);
            Console.WriteLine("=".Repeat(header.Length));
        }

        private void startupMenu()
        {
            var menuData = new Dictionary<string, List<string>>
            {
                {
                    "ВЫБЕРИТЕ ДЕЙСТВИЕ",
                    new List<string> { "СОЗДАТЬ НОВЫЙ СПИСОК", "ВЫБРАТЬ СПИСОК", "УДАЛИТЬ СПИСОК" }
                },
                { "-----", new List<string> { "ВЫХОД" } }
            };
            menu.RunMenuWithActions(menuData, this.NewDataList, this.MenuDataList, this.DeleteDataList, this.CloseApp);
        }

        private void NewDataList()
        {
            this.data.Add(0, "Заметка 1");
            this.data.Add(1, "Заметка 2");

            ShowMessage("Создан новый лист!");
        }

        private void MenuDataList()
        {
            MenuList("Заметки");
        }

        private void DeleteDataList()
        {

        }

        public int SelectListMemberId(IDictionary<int, string> members)
        {
            return 0;
        }

        public int MenuListData()
        {
            var menuItems = data.Values.ToArray();
            int itemId = menu.ContentMenu(menuItems, "ЗАМЕТКИ");

            if (itemId < 0)
                return itemId;

            var selectedItem = menuItems[itemId];

            foreach (var item in data)
            {
                if (item.Value.Equals(selectedItem))
                {
                    return item.Key;
                }
            }

            throw new InvalidOperationException("Selected item not found.");
        }

        private void MenuList(string name)
        {
            var menuData = new Dictionary<string, List<string>>
            {
                {
                    "ДЕЙСТВИЯ СО СПИСКОМ: " + name,
                    new List<string> { "ДОБАВИТЬ СТРОКУ", "ПОКАЗАТЬ СТРОКИ", "УДАЛИТЬ СТРОКУ" }
                },
                { "-----", new List<string> { "ВЫХОД В ГЛАВНОЕ МЕНЮ" } }
            };

            menu.RunMenuWithActions(menuData, this.AddItemToList, this.MenuShowData,
                    this.DeleteData, this.CloseListMenu);
        }

        private void AddItemToList()
        {

        }

        private void MenuShowData()
        {

            if (data.Count() == 0)
            {
                if (menu.YesNoDialog("Нет данных. Добавить?"))
                {
                    AddItemToList();
                }
                return;
            }

            MenuListData();
        }

        private void DeleteData()
        {

        }

        private void CloseListMenu()
        {
            menu.CloseMenu();
        }

        private void CloseApp()
        {
            if (menu.YesNoDialog("Завершить работу приложения?"))
            {
                appIsRun = false;
                menu.CloseMenu();

                Console.Clear();
                Console.WriteLine("Приложение завершило работу. Для выхода нажмите любую клавишу...");
                Console.ReadKey();
                Console.Clear();
            }
        }
    }
}
