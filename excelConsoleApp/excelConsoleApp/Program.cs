using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace excelConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // 엑셀 어플리케이션 객체 생성
            Application excel = new Application();

            // 새 워크북 생성
            Workbook workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            // 첫번째 워크시트 선택
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            // 데이터 추가
            worksheet.Cells[1, 1] = "이름";
            worksheet.Cells[1, 2] = "나이";
            worksheet.Cells[2, 1] = "홍길동";
            worksheet.Cells[2, 2] = "30";

            // 파일 저장
            //string filePath = @"C:\example.xlsx";
            string filePath = @"D:\C#\Excel\example.xlsx";
            workbook.SaveAs(filePath);

            // 엑셀 어플리케이션 종료
            excel.Quit();
        }
    }
}
