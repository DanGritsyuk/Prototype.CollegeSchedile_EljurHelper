namespace LessonLinker.Common.Entities.Extensions
{
    public static class StringExtensions
    {
        public static string Repeat(this string str, int count) =>
             new string(Enumerable.Repeat(str, count).SelectMany(s => s).ToArray());

        public static bool IsEmpty(this string str) =>
            string.IsNullOrEmpty(str);
    }
}
