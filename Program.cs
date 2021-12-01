using System;
using System.Diagnostics;

namespace cs_con_framework_excel_new
{
    class Program
    {
        static void Main(string[] args)
        {
            // Excel アプリケーション
            dynamic excelApp =
                Activator
                    .CreateInstance(Type
                        .GetTypeFromProgID("Excel.Application"));

            // Excel のパス( カレントディレクトリに作成 )
            string path = Environment.CurrentDirectory + @"\sample.xlsx";

            // Excel を表示( 完成したらコメント化 )
            excelApp.Visible = true;

            // 警告を出さない
            excelApp.DisplayAlerts = false;

            // // ブックを追加( 新規 )
            dynamic workbooks = excelApp.Workbooks;
            dynamic book = workbooks.Add();

            dynamic sheet = book.Sheets(1);
            // ブックを取得( 一つしかないので、Count は 1 )
            book = excelApp.Workbooks(excelApp.Workbooks.Count);
            book.Sheets(1).Name = "最初のシート";
            // https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheets.add
            book.Sheets.Add(After: book.Sheets(1));
            book.Sheets(2).Name = "追加のシート";
            // 先頭シートをアクティブにする
            book.Sheets(1).Activate();
            // セルに値をセット
            book.Sheets(1).Cells(1, 1).Value = "社員名";
            book.Sheets(1).Cells(2, 1).Value = "山田　太郎甚左衛門";
            book.Sheets(1).Cells(3, 1).Value = "鈴木　一郎";
            book.Sheets(1).Cells(4, 1).Value = "佐藤　洋子";
            // 列幅自動調整
            book.Sheets(1).Columns("A:A").EntireColumn.AutoFit();
            // 保存
            book.SaveAs(path);

            // 閉じる
            book.Close();

            // Excel 終了
            excelApp.Quit();

            // 解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject (excelApp);

            // C# ではほぼ完全解放無理なので強制終了させる
            foreach (var p in Process.GetProcessesByName("EXCEL"))
            {
                if (p.MainWindowTitle == "")
                {
                    p.Kill();
                }
            }

            // ファイルの種類によってアプリケーションを起動する
            ProcessStartInfo processStartInfo = new ProcessStartInfo("RunDLL32.EXE", $"url.dll,FileProtocolHandler {path}" );
            Process.Start(processStartInfo);
        }
    }
}
