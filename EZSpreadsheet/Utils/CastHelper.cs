using System.Reflection;

namespace EZSpreadsheet.Utils
{
    internal class CastHelper
    {
        internal static dynamic Cast(object src, Type type)
        {
            var nullableType = Nullable.GetUnderlyingType(type);
            if (nullableType != null)
                type = nullableType;

            var castMethod = typeof(CastHelper)
                .GetMethod("CastGeneric", BindingFlags.Static | BindingFlags.NonPublic)!
                .MakeGenericMethod(type);
            return castMethod.Invoke(null, new[] { src })!;
        }
        internal static T CastGeneric<T>(object src)
        {
            return (T)Convert.ChangeType(src, typeof(T));
        }
    }
}
