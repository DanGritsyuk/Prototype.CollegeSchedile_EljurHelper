using LessonLinker.Common.Entities.Schedule;

namespace LessonLinker.Common.Entities.Extensions
{
    public static class ScheduleDaysExtensions
    {
        public static string ToRussianString(this string day)
        {
            return day switch
            {
                "Monday" => "Понедельник",
                "Tuesday" => "Вторник",
                "Wednesday" => "Среда",
                "Thursday" => "Четверг",
                "Friday" => "Пятница",
                "Saturday" => "Суббота",
                "Sunday" => "Воскресенье",
                _ => throw new ArgumentOutOfRangeException(nameof(day), day, null)
            };
        }
    }
}
