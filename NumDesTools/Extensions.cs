using System.Linq;

namespace NumDesTools;

public static class Extensions
{
    //string[][]×ª»»Îªstring[,]
    public static T[,] ToRectangularArray<T>(this T[][] source)
    {
        var rowCount = source.Length;
        var colCount = source.Max(x => x.Length);
        var result = new T[rowCount, colCount];
        for (var r = 0; r < rowCount; r++)
        {
            var row = source[r];
            for (var c = 0; c < colCount; c++)
            {
                if (c < row.Length)
                    result[r, c] = row[c];
                else
                    result[r, c] = default(T);
            }
        }
        return result;
    }
}