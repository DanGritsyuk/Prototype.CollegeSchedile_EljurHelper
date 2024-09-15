namespace LessonLinker.Common.Entities.Schedule
{
    public class Lesson
    {
        public string Subject { get; set; } // Название предмета
        public string Teacher { get; set; } // Фамилия И.О. преподавателя
        public string? Room { get; set; } // Кабинет
        public int Group { get; set; } // Группа (если есть деление)
    }
}
