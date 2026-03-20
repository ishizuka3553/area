using System;
using System.Diagnostics;
using System.IO;
using NPOI.XWPF.UserModel;

namespace WordTableReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. 引数のチェック
            if (args.Length == 0)
            {
                Console.WriteLine("使用法: WordTableReader.exe <ファイルパス>");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"エラー: ファイルが見つかりません: {filePath}");
                return;
            }

            // 2. ストップウォッチの開始
            Stopwatch sw = new Stopwatch();
            sw.Start();

            try
            {
                Console.WriteLine($"{Path.GetFileName(filePath)} を読み取り中...");

                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Wordドキュメントの読み込み
                    XWPFDocument doc = new XWPFDocument(fs);

                    // 3. 表（テーブル）の読み取り
                    int tableCount = 1;
                    foreach (var table in doc.Tables)
                    {
                        Console.WriteLine($"--- 表 {tableCount++} ---");
                        foreach (var row in table.Rows)
                        {
                            // 各セルのテキストをタブ区切りで結合して表示
                            var cellTexts = row.GetTableCells().ConvertAll(c => c.GetText());
                            Console.WriteLine(string.Join(" | ", cellTexts));
                        }
                        Console.WriteLine();
                    }
                }

                // 4. 計測終了
                sw.Stop();
                Console.WriteLine("----------------------------------");
                Console.WriteLine($"処理完了。");
                Console.WriteLine($"実行時間: {sw.ElapsedMilliseconds} ms ({sw.Elapsed.TotalSeconds:F2} 秒)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
        }
    }
}