namespace LessonLinker.Common.Entities.Schedule
{
    internal class AcademicDay
    {
        public Dictionary<int, IEnumerable<Lesson>> lessons { get; set; } // словарь: номер пары, урок
    }
}
