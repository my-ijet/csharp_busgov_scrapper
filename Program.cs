using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using ClosedXML.Excel;

var today_str = DateTime.Today.ToString();

//read the xlsx file
var name_of_file = "список.xlsx";
using var workbook = new XLWorkbook(name_of_file);
var ws = workbook.Worksheet(1);

int gl_row = 3;
int gl_column = 4;

// get last column of reviews
while (true)
{
  var date_of_review = ws.Cell(1, gl_column).Value.ToString();
  if (string.IsNullOrEmpty(date_of_review)) break;
  if (date_of_review == today_str) { break; }
  gl_column += 2;
}
// set date of review
ws.Cell(1, gl_column).Value = DateTime.Today;
ws.Range(ws.Cell(1, gl_column), ws.Cell(1, gl_column + 1)).Merge();
ws.Cell(2, gl_column).Value = "Оценка";
ws.Cell(2, gl_column + 1).Value = "Кол-во отзывов";

while (true)
{
  var name_of_org = ws.Cell(gl_row, 3).Value.ToString();
  if (string.IsNullOrEmpty(name_of_org)) break;
  Console.WriteLine($"Обработка {name_of_org} ...");
  var org = new OrgInfo(name_of_org);

  var web_address = ws.Cell(gl_row, 2).Value.ToString();
  if (string.IsNullOrEmpty(web_address))
  {
    Console.WriteLine("\t отсутствует вэб адрес");
    next_org();
    continue;
  }
  org.web_address = web_address;

  try
  {
    get_org_info(ref org);
  }
  catch (System.Exception)
  {
    Console.WriteLine("\t ошибка получения информации со страницы");
    next_org();
    continue;
  }

  ws.Cell(gl_row, gl_column).Value = org.num_average;         // Общая средняя оценка
  ws.Cell(gl_row, gl_column + 1).Value = org.num_of_reviews;  // Количество отзывов
  ws.Cell(gl_row + 1, gl_column).Value = org.num5;            // 5
  ws.Cell(gl_row + 2, gl_column).Value = org.num4;            // 4
  ws.Cell(gl_row + 3, gl_column).Value = org.num3;            // 3
  ws.Cell(gl_row + 4, gl_column).Value = org.num2;            // 2
  ws.Cell(gl_row + 5, gl_column).Value = org.num1;            // 1
  ws.Cell(gl_row + 6, gl_column).Value = org.num_total;       // Итого оценок

  // Prettier
  ws.Cell(gl_row, gl_column).Style.NumberFormat.Format = "0.00";
  ws.Cell(gl_row, gl_column + 1).Style.NumberFormat.Format = "0";
  ws.Cell(gl_row + 1, gl_column).Style.NumberFormat.Format = "0";
  ws.Cell(gl_row + 2, gl_column).Style.NumberFormat.Format = "0";
  ws.Cell(gl_row + 3, gl_column).Style.NumberFormat.Format = "0";
  ws.Cell(gl_row + 4, gl_column).Style.NumberFormat.Format = "0";
  ws.Cell(gl_row + 5, gl_column).Style.NumberFormat.Format = "0";
  ws.Cell(gl_row + 6, gl_column).Style.NumberFormat.Format = "0";

  ws.Range(
    ws.Cell(gl_row + 1, gl_column + 1),
    ws.Cell(gl_row + 6, gl_column + 1)
   ).Merge();


  next_org();
}

Console.Write("Обработка завершена.");
Console.WriteLine("нажмите любую кнопку для завершения ...");
Console.ReadKey(true);
workbook.Save();

cross_platform_open_file(name_of_file);


void get_org_info(ref OrgInfo org_info)
{
}

[MethodImpl(MethodImplOptions.AggressiveInlining)]
void next_org()
{ gl_row += 7; }

void cross_platform_open_file(string filePath)
{
  if (OperatingSystem.IsMacOS())
  {
    Process.Start("open", filePath);
  }
  else if (OperatingSystem.IsWindows())
  {
    Process.Start(filePath);
  }
  else
  {
    Process.Start("xdg-open", filePath);
  }
}

record struct OrgInfo
{
  public string name;
  public string web_address = "";
  public int num_of_reviews = 0;
  public int num5 = 0, num4 = 0, num3 = 0, num2 = 0, num1 = 0;
  readonly public int num_total { get { return num5 + num4 + num3 + num2 + num1; } }
  readonly public float num_average
  {
    get
    {
      float number = (1 * num1 + 2 * num2 + 3 * num3 + 4 * num4 + 5 * num5) / (float)num_total;
      return MathF.Round(number, 2);
    }
  }

  public OrgInfo(string new_name)
  {
    name = new_name;
  }
}