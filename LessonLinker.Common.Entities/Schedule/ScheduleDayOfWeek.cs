namespace LessonLinker.Common.Entities.Schedule
{
    public class ScheduleDays
    {
        public static string Monday = "Понедельник";
        public static string Tuesday = "Вторник";
        public static string Wednesday = "Среда";
        public static string Thursday = "Четверг";
        public static string Friday = "Пятница";
        public static string Saturday = "Суббота";
        public static string Sunday = "Воскресенье";

        // Опционально: Можете добавить метод для получения всех дней недели
        public static IEnumerable<string> GetAllDays()
        {
            return new List<string>
        {
            Monday,
            Tuesday,
            Wednesday,
            Thursday,
            Friday,
            Saturday,
            Sunday
        };
        }
    }

}
