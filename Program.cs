using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;

var chromeDriverService = ChromeDriverService.CreateDefaultService();
chromeDriverService.HideCommandPromptWindow = true;  // This line hides the command prompt window
chromeDriverService.EnableVerboseLogging = false;    // This line disables verbose logging
chromeDriverService.SuppressInitialDiagnosticInformation = true;  // This line suppresses initial diagnostic information

var driver_options = new ChromeOptions();
driver_options.AddArgument("headless");
IWebDriver driver = new ChromeDriver(chromeDriverService, driver_options);

AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionHandler;

//read the xlsx file
var name_of_file = "список.xlsx";
using var workbook = new XLWorkbook(name_of_file);
var ws = workbook.Worksheet(1);

var today_str = DateTime.Today.ToString();
var wait_seconds = 15;

int gl_row = 3;
int gl_column = 4;


Console.CursorVisible = false;
Console.Clear();

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
  Console.Write($"Обработка {name_of_org} ...");
  var org = new OrgInfo(name_of_org);

  // Prettier
  format_org_cells();

  var web_address = ws.Cell(gl_row, 2).Value.ToString();
  if (string.IsNullOrEmpty(web_address))
  {
    Console.WriteLine("\n\t отсутствует вэб адрес");
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
    Console.WriteLine("\n\t ошибка получения информации со страницы");
    next_org();
    continue;
  }
  Console.WriteLine("OK                                 ");

  ws.Cell(gl_row, gl_column).Value = org.num_average;         // Общая средняя оценка
  ws.Cell(gl_row, gl_column + 1).Value = org.num_of_reviews;  // Количество отзывов
  ws.Cell(gl_row + 1, gl_column).Value = org.num5;            // 5
  ws.Cell(gl_row + 2, gl_column).Value = org.num4;            // 4
  ws.Cell(gl_row + 3, gl_column).Value = org.num3;            // 3
  ws.Cell(gl_row + 4, gl_column).Value = org.num2;            // 2
  ws.Cell(gl_row + 5, gl_column).Value = org.num1;            // 1
  ws.Cell(gl_row + 6, gl_column).Value = org.num_total;       // Итого оценок


  next_org();
}
outline_cells();

Console.Write("Обработка завершена.");
Console.WriteLine("нажмите любую кнопку для завершения ...");
Console.ReadKey(true);
workbook.Save();

cross_platform_open_file(name_of_file);
driver.Quit();
// end

void outline_cells()
{
  var range = ws.Range(
    ws.Cell(1, gl_column),
    ws.Cell(gl_row - 1, gl_column + 1));
  range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
  range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
}

void format_org_cells()
{
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
}

void get_org_info(ref OrgInfo org)
{
  // loading the target web page 
  driver.Navigate().GoToUrl(org.web_address);

  var cursor_position = Console.GetCursorPosition();
  cursor_position.Left -= 3;
  Console.SetCursorPosition(cursor_position.Left, cursor_position.Top);
  for (int i = 0; i < wait_seconds; i++)
  {
    Console.Write($"({wait_seconds - i})                                 ");
    Console.SetCursorPosition(cursor_position.Left, cursor_position.Top);
    Thread.Sleep(1000); // sleep for 1000 milliseconds = 1 second
  }

  var div = driver.FindElement(By.ClassName("independent-rating-tab-feedback-title-text"));
  div = div.FindElement(By.TagName("span"));
  org.num_of_reviews = Int32.Parse(div.Text.TrimStart('(').TrimEnd(')'));

  div = driver.FindElement(By.ClassName("rating-by-values"));
  var rows = div.FindElements(By.ClassName("rating-by-values-row"));

  div = rows[0].FindElement(By.ClassName("count-of-votes"));
  org.num5 = Int32.Parse(div.Text);
  div = rows[1].FindElement(By.ClassName("count-of-votes"));
  org.num4 = Int32.Parse(div.Text);
  div = rows[2].FindElement(By.ClassName("count-of-votes"));
  org.num3 = Int32.Parse(div.Text);
  div = rows[3].FindElement(By.ClassName("count-of-votes"));
  org.num2 = Int32.Parse(div.Text);
  div = rows[4].FindElement(By.ClassName("count-of-votes"));
  org.num1 = Int32.Parse(div.Text);
}

[MethodImpl(MethodImplOptions.AggressiveInlining)]
void next_org()
{ gl_row += 7; }

void UnhandledExceptionHandler(object sender, UnhandledExceptionEventArgs e)
{
  Exception ex = (Exception)e.ExceptionObject;
  Console.WriteLine($"Произошла ошибка:\n{ex.Message}");
  // Handle the uncaught exception here
  driver.Quit();
}

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
      if (num_total == 0) { return 0; }
      float number = (1 * num1 + 2 * num2 + 3 * num3 + 4 * num4 + 5 * num5) / (float)num_total;
      return MathF.Round(number, 2);
    }
  }

  public OrgInfo(string new_name)
  {
    name = new_name;
  }
}