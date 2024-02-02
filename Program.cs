using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;


var wait_seconds = 30;
// parse command line
if (args.Length > 0)
{
  var first_arg = args[0];
  if (first_arg == "-h" || first_arg == "--help")
  {
    print_comand_line_help();
    return;
  }
  if (args.Length < 2 || args.Length > 2)
  {
    Console.WriteLine("Неверное количество аргументов.");
    print_comand_line_help();
    return;
  }

  var second_arg = args[1];
  if (first_arg == "-t")
  {
    try { wait_seconds = int.Parse(second_arg); }
    catch (System.Exception e)
    {
      Console.WriteLine($"Ошибка преобразования второго аргумента:\n{e.Message}");
      print_comand_line_help();
      return;
    }
  }
  else
  {
    Console.WriteLine("Нет таких аргументов.");
    print_comand_line_help();
    return;
  }
}

void print_comand_line_help()
{
  Console.WriteLine("\nИспользование: -t <время_ожидания_в_секундах>");
}
// parse command line



var chromeDriverService = ChromeDriverService.CreateDefaultService();
chromeDriverService.HideCommandPromptWindow = true;  // This line hides the command prompt window
chromeDriverService.EnableVerboseLogging = false;    // This line disables verbose logging
chromeDriverService.SuppressInitialDiagnosticInformation = true;  // This line suppresses initial diagnostic information

var driver_options = new ChromeOptions();
driver_options.AddArgument("headless");
IWebDriver driver = new ChromeDriver(chromeDriverService, driver_options);

AppDomain.CurrentDomain.UnhandledException += (object sender, UnhandledExceptionEventArgs e) =>
                                              {
                                                Exception ex = (Exception)e.ExceptionObject;
                                                Console.WriteLine($"Произошла ошибка:\n{ex.Message}");
                                                // Handle the uncaught exception here
                                                // quit_app();
                                              };

Console.CancelKeyPress += new ConsoleCancelEventHandler(
    (object? sender, ConsoleCancelEventArgs e) => { quit_app(); }
  );

//read the xlsx file
var name_of_file = "список.xlsx";
using var workbook = new XLWorkbook(name_of_file);
var ws = workbook.Worksheet(1);

var today_str = DateTime.Today.ToString();

int gl_row = 3;
int gl_column = 4;


Console.CursorVisible = false;
Console.Clear();

Console.WriteLine($"Ожидание загрузки страницы максимум {wait_seconds} сек.");
Console.WriteLine("Используйте аргумент -t для изменения.");

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
  Console.WriteLine(" ГОТОВО.                                 ");

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

workbook.Save();

Console.CursorVisible = true;
Console.WriteLine("Обработка завершена.");
Console.WriteLine("нажмите любую кнопку для завершения ...");
Console.ReadKey(true);

cross_platform_open_file(Path.GetFullPath(name_of_file));

quit_app();


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
  var cell_num_average = ws.Cell(gl_row, gl_column);          // Общая средняя оценка
  var cell_num_of_reviews = ws.Cell(gl_row, gl_column + 1);   // Количество отзывов
  var cell_num5 = ws.Cell(gl_row + 1, gl_column);             // 5
  var cell_num4 = ws.Cell(gl_row + 2, gl_column);             // 4
  var cell_num3 = ws.Cell(gl_row + 3, gl_column);             // 3
  var cell_num2 = ws.Cell(gl_row + 4, gl_column);             // 2
  var cell_num1 = ws.Cell(gl_row + 5, gl_column);             // 1
  var cell_num_total = ws.Cell(gl_row + 6, gl_column);        // Итого оценок

  List<IXLCell> list_of_cells = [cell_num_average, cell_num_of_reviews, cell_num5, cell_num4, cell_num3, cell_num2, cell_num1, cell_num_total];

  for (int i = 0; i < list_of_cells.Count; i++)
  {
    string new_format = "0";
    if (i == 0) new_format = "0.00";
    list_of_cells[i].Style.NumberFormat.Format = new_format;
  }

  ws.Range(
    ws.Cell(gl_row + 1, gl_column + 1),
    ws.Cell(gl_row + 6, gl_column + 1)
   ).Merge();

  if (gl_column == 4) return;
  // previous rating
  var prev_cell_num_average = ws.Cell(gl_row, gl_column - 2);          // Общая средняя оценка
  var prev_cell_num_of_reviews = ws.Cell(gl_row, gl_column - 1);       // Количество отзывов
  var prev_cell_num5 = ws.Cell(gl_row + 1, gl_column - 2);             // 5
  var prev_cell_num4 = ws.Cell(gl_row + 2, gl_column - 2);             // 4
  var prev_cell_num3 = ws.Cell(gl_row + 3, gl_column - 2);             // 3
  var prev_cell_num2 = ws.Cell(gl_row + 4, gl_column - 2);             // 2
  var prev_cell_num1 = ws.Cell(gl_row + 5, gl_column - 2);             // 1
  var prev_cell_num_total = ws.Cell(gl_row + 6, gl_column - 2);        // Итого оценок

  List<IXLCell> list_of_prev_cells = [prev_cell_num_average, prev_cell_num_of_reviews, prev_cell_num5, prev_cell_num4, prev_cell_num3, prev_cell_num2, prev_cell_num1, prev_cell_num_total];

  for (int i = 0; i < list_of_prev_cells.Count; i++)
  {
    var cell = list_of_cells[i];
    var prev_cell = list_of_prev_cells[i];
    string symbol1 = (i > 3 && i < 7) ? "<" : ">";
    string symbol2 = (i > 3 && i < 7) ? ">" : "<";

    cell.AddConditionalFormat()
      .WhenIsTrue($"{cell} {symbol1} {prev_cell}")
      .Fill.SetBackgroundColor(XLColor.PastelGreen);
    cell.AddConditionalFormat()
      .WhenIsTrue($"{cell} {symbol2} {prev_cell}")
      .Fill.SetBackgroundColor(XLColor.PastelRed);
  }

}

void get_org_info(ref OrgInfo org)
{
  // loading the target web page
  driver.Navigate().GoToUrl(org.web_address);
  IWebElement div;
  Thread.Sleep(1000); // sleep for 1000 milliseconds = 1 second

  var cursor_position = Console.GetCursorPosition();
  cursor_position.Left -= 3;
  Console.SetCursorPosition(cursor_position.Left, cursor_position.Top);
  for (int i = 0; i < wait_seconds; i++)
  {
    Console.Write($"({wait_seconds - i})                                 ");
    Console.SetCursorPosition(cursor_position.Left, cursor_position.Top);
    Thread.Sleep(1000); // sleep for 1000 milliseconds = 1 second

    div = driver.FindElement(By.ClassName("independent-rating-tab-feedback-title-text"));
    div = div.FindElement(By.TagName("span"));
    org.num_of_reviews = Int32.Parse(div.Text.TrimStart('(').TrimEnd(')'));
    if (org.num_of_reviews > 0) break;
  }

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

void quit_app()
{
  Console.CursorVisible = true;
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
    Process.Start(@$"{filePath}");
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
