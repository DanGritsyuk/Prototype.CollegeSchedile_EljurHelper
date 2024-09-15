namespace LessonLinker.UI.TextMenu.MenuRendering
{
    class PageData
    {
        private int page_id;
        private int start_line_index;
        private int current_line_index;
        private Dictionary<string, List<string>> data;
        private int lines_count;

        public PageData(int page_id, int start_line_index, Dictionary<string, List<string>>? page_data = null, int current_line_index = 0)
        {
            this.page_id = page_id;
            this.start_line_index = start_line_index;
            this.current_line_index = current_line_index;
            this.data = DataInit(page_data);
            this.lines_count = TasksCount();
        }

        public int PageId
        {
            get { return page_id; }
            set { page_id = value; }
        }

        public int StartLineIndex
        {
            get { return start_line_index; }
            set { start_line_index = value; }
        }

        public int CurrentLineIndex
        {
            get { return current_line_index; }
            set { current_line_index = value; }
        }

        public Dictionary<string, List<string>> Data
        {
            get { return data; }
            set { data = DataInit(value); }
        }

        public int LinesCount
        {
            get { return lines_count; }
            set { lines_count = value; }
        }

        private int TasksCount()
        {
            int count = 0;
            foreach (var key_work in data)
            {
                count += key_work.Value.Count;
            }
            return count;
        }

        private Dictionary<string, List<string>> DataInit(Dictionary<string, List<string>>? page_data) =>
            page_data != null ? new Dictionary<string, List<string>>(page_data) : new Dictionary<string, List<string>>();
    }
}