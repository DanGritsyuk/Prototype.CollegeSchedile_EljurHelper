namespace LessonLinker.Common.Entities.Schedule
{
    internal class WeeklySchedule
    {
        public string GruopName { get; set; }
        public Dictionary<ScheduleDays, IEnumerable<AcademicDay>> StudyDays { get; set; }
    }
}