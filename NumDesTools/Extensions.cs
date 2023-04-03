using System.Linq;

namespace NumDesTools;

public static class Extensions
{
    //string[][]×ª»»Îªstring[,]
    public static T[,] ToRectangularArray<T>(this T[][] source)
    {
        int rowCount = source.Length;
        int colCount = source.Max(x => x.Length);
        T[,] result = new T[rowCount, colCount];
        for (int r = 0; r < rowCount; r++)
        {
            T[] row = source[r];
            for (int c = 0; c < colCount; c++)
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