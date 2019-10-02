using System;
using System.Collections.Generic;
using System.Globalization;

namespace LogPresence
{
    internal static class StringExtensions
    {
        public static string[] CsvSplit(this string line, char separator, char colEscapeChar)
        {
            var res = new List<string>();

            int currStart = 0;

            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == separator)
                {
                    res.Add(line.Substring(currStart, i - currStart));
                    currStart = i + 1;
                }
                else if (currStart == i && line[i] == colEscapeChar)
                {
                    currStart = i + 1;
                    i++;
                    bool done = false;
                    bool doReplacement = false;
                    while (!done)
                    {
                        while (i < line.Length && line[i] != colEscapeChar)
                        {
                            i++;
                        }

                        if (i == line.Length)
                        {
                            throw new InvalidOperationException($"No end to line: >" + line + "<");
                        }

                        if (i < line.Length - 1 && line[i + 1] == colEscapeChar)
                        {
                            i += 2;
                            done = false;
                            doReplacement = true;
                        }
                        else
                        {
                            if (i != line.Length - 1 && line[i + 1] != separator)
                            {
                                throw new InvalidOperationException("Escaped column not ending in separator: >" + line + "<");
                            }

                            var col = line.Substring(currStart, i - currStart);
                            col = doReplacement ? col.Replace(new string(colEscapeChar, 2), colEscapeChar.ToString(CultureInfo.InvariantCulture)) : col;
                            res.Add(col);
                            i++;
                            currStart = i + 1;
                            done = true;
                        }
                    }
                }
            }

            if (currStart <= line.Length)
            {
                res.Add(line.Substring(currStart, line.Length - currStart));
            }

            return res.ToArray();
        }
    }
}
