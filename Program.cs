using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/* To work eith EPPlus library */
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

/* For I/O purpose */
using System.IO;

/* For Diagnostics */
using System.Diagnostics;
using System.Data;
using Itenso.TimePeriod;
using System.Drawing;

namespace AutoAllocatev2
{
    class Program
    {
        private static List<Color> COLORS = new List<Color> { Color.Aqua,Color.Green,Color.LightBlue,Color.Yellow,Color.Snow };
        static void Main(string[] args)
        {
            var File = new FileInfo(@"..\..\..\Resource Allocation Sheet.xlsx");
            using (ExcelPackage package = new ExcelPackage(File))
            {
                ExcelWorkbook workBook = package.Workbook;
                if (workBook != null)
                {
                    Console.WriteLine(workBook.Worksheets.Count);
                    
                    ExcelWorksheet workSheet = workBook.Worksheets["Resources"];
                    List<Resource> resourceList = ReadResource(workSheet);
                    List<Resource> availableResourceList = new List<Resource>();
                    availableResourceList.AddRange(resourceList);

                    foreach (Resource resource in resourceList)
                    {
                        Console.WriteLine(resource.ToString());
                    }

                    ExcelWorksheet reqSheet = workBook.Worksheets["Requirement"];
                    List<Requirement> requirementList = ReadRequirement(reqSheet, resourceList);

                    //Sort Requirement By Priority
                    requirementList = requirementList.OrderBy(req => req.Priority).ToList();

                    foreach (Requirement requirement in requirementList)
                    {
                        Console.WriteLine(requirement.ToString());
                    }

                    
                    DateTime StartDate, EndDate;

                    FindStartAndEndDate(resourceList, out StartDate, out EndDate);
                    Console.WriteLine("StartDate ={0},EndDate={1}", StartDate, EndDate);
                  
                    DurationProvider provider = new DurationProvider();
                    List<Allocation> allocationList = new List<Allocation>();
                    foreach (Requirement requirement in requirementList)
                    {
                        Allocation allocation = new Allocation();
                        allocation.AllocationStartDate = StartDate;
                        allocation.AllocationEndDate = EndDate;
                        allocation.requirement = requirement;

                        foreach (KeyValuePair<string, int> kvp in requirement.EffortByType)
                        {

                            int maximumResource = requirement.MaximimResourcesByType[kvp.Key];
                            int longPole = requirement.LongPoleByType[kvp.Key];
                            foreach (Resource resource in availableResourceList)
                            {
                                if (kvp.Key.Equals(resource.Type))
                                {

                                    if (BusinessDaysUntil(resource.AvailableStartDate1, resource.AvailableEndDate1) >= kvp.Value)
                                    {
                                        Resource availableResource = new Resource();
                                        availableResource.Name = resource.Name;
                                        availableResource.Type = kvp.Key;
                                        availableResource.AvailableStartDate1 = resource.AvailableStartDate1;
                                        availableResource.AvailableEndDate1 = AddWorkDays(resource.AvailableStartDate1,kvp.Value);
                                        resource.AvailableStartDate1 = availableResource.AvailableEndDate1;

                                        DateTime tempDate = availableResource.AvailableStartDate1;
                                        while (DateTime.Compare(tempDate, availableResource.AvailableEndDate1) != 0)
                                        {
                                            availableResource.AllocationMap.Add(tempDate, true);
                                          tempDate = tempDate.AddDays(1);
                                        }

                                        allocation.resourceList.Add(availableResource);
                                        break;
                                    }
                                    else if (BusinessDaysUntil(resource.AvailableStartDate2, resource.AvailableEndDate2) >= kvp.Value)
                                    {
                                        Resource availableResource = new Resource();
                                        availableResource.Name = resource.Name;
                                        availableResource.Type = kvp.Key;
                                        resource.AvailableStartDate2 = availableResource.AvailableStartDate2 = resource.AvailableStartDate2;
                                        resource.AvailableEndDate2 = availableResource.AvailableEndDate2 = AddWorkDays(resource.AvailableStartDate2,kvp.Value);
                                        DateTime tempDate = availableResource.AvailableStartDate2;
                                        while (DateTime.Compare(tempDate, availableResource.AvailableEndDate2) != 0)
                                        {
                                            availableResource.AllocationMap.Add(tempDate, true);
                                            tempDate = tempDate.AddDays(1);
                                        }


                                        allocation.resourceList.Add(availableResource);
                                        break;
                                    }
                                    else if (BusinessDaysUntil(resource.AvailableStartDate3, resource.AvailableEndDate3) >= kvp.Value)
                                    {
                                        Resource availableResource = new Resource();
                                        availableResource.Name = resource.Name;
                                        availableResource.Type = kvp.Key;
                                        resource.AvailableStartDate3 = availableResource.AvailableStartDate3 = resource.AvailableStartDate3;
                                        resource.AvailableEndDate3 = availableResource.AvailableEndDate3 = AddWorkDays(resource.AvailableStartDate3,kvp.Value);
                                        DateTime tempDate = availableResource.AvailableStartDate3;
                                        while (DateTime.Compare(tempDate, availableResource.AvailableEndDate3) != 0)
                                        {
                                            availableResource.AllocationMap.Add(tempDate, true);
                                            tempDate = tempDate.AddDays(1);
                                        }

                                        allocation.resourceList.Add(availableResource);
                                        break;
                                    }
                                    else if (BusinessDaysUntil(resource.AvailableStartDate4, resource.AvailableEndDate4) >= kvp.Value)
                                    {
                                        Resource availableResource = new Resource();
                                        availableResource.Name = resource.Name;
                                        availableResource.Type = kvp.Key;
                                        resource.AvailableStartDate4 = availableResource.AvailableStartDate4 = resource.AvailableStartDate4;
                                        resource.AvailableEndDate4 = availableResource.AvailableEndDate4 = AddWorkDays(resource.AvailableStartDate4,kvp.Value);
                                        DateTime tempDate = availableResource.AvailableStartDate4;
                                        while (DateTime.Compare(tempDate, availableResource.AvailableEndDate4) != 0)
                                        {
                                            availableResource.AllocationMap.Add(tempDate, true);
                                            tempDate = tempDate.AddDays(1);
                                        }

                                        allocation.resourceList.Add(availableResource);
                                        break;
                                    }
                                    else if (BusinessDaysUntil(resource.AvailableStartDate5, resource.AvailableEndDate5) >= kvp.Value)
                                    {
                                        Resource availableResource = new Resource();
                                        availableResource.Name = resource.Name;
                                        availableResource.Type = kvp.Key;
                                        resource.AvailableStartDate5 = availableResource.AvailableStartDate5 = resource.AvailableStartDate5;
                                        resource.AvailableEndDate5 = availableResource.AvailableEndDate5 = AddWorkDays(resource.AvailableStartDate5,kvp.Value);
                                        DateTime tempDate = availableResource.AvailableStartDate5;
                                        while (DateTime.Compare(tempDate, availableResource.AvailableEndDate5) != 0)
                                        {
                                            availableResource.AllocationMap.Add(tempDate, true);
                                            tempDate = tempDate.AddDays(1);
                                        }

                                        allocation.resourceList.Add(availableResource);
                                        break;
                                    }
                                }
                            }
                        }
                        allocationList.Add(allocation);
                        Console.WriteLine(allocation.ToString());
                    }


                    foreach (Resource resource in availableResourceList)
                    {
                        Console.WriteLine(resource.ToString());
                    }

                    ExcelWorksheet allocationSheet = package.Workbook.Worksheets.Add("Allocation");
                    allocationSheet.Column(1).Width = 12;
                    allocationSheet.Column(1).Style.Font.Size = 8;                    
                    allocationSheet.Column(2).Width = 12;
                    allocationSheet.Column(2).Style.Font.Size = 8;
                    allocationSheet.Cells[1,1].Value = "Requirement";
                    allocationSheet.Cells[1, 1, 2, 1].Merge = true;
                    
                    allocationSheet.Cells[1,2].Value = "Resource";
                    allocationSheet.Cells[1, 2, 2, 2].Merge = true;
                    int cellCount = 3;
                    while (DateTime.Compare(StartDate,EndDate)!=0)
                    {
                        if (StartDate.DayOfWeek != DayOfWeek.Saturday && StartDate.DayOfWeek != DayOfWeek.Sunday)
                        {
                            allocationSheet.Cells[1, cellCount].Value = StartDate.ToString("dd-MMM");
                            allocationSheet.Cells[1, cellCount].Style.Font.Size = 8;
                            allocationSheet.Cells[2, cellCount].Value = StartDate.ToString("ddd");
                            allocationSheet.Cells[2, cellCount].Style.Font.Size = 8;
                            allocationSheet.Cells[2, cellCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            allocationSheet.Column(cellCount).Width = 4.80;
                            cellCount++;
                                
                        }
                        StartDate = StartDate.AddDays(1);
                    }
                    allocationSheet.Row(2).Height = 9.80;
                    int rowCount = 3;
                    foreach (Allocation alcn in allocationList)
                    {                       

                        Random rand = new Random();
                        int max = byte.MaxValue + 1; // 256
                        int r = rand.Next(max);
                        int g = rand.Next(max);
                        int b = rand.Next(max);
                        Color color = Color.FromArgb(r, g, b);

                        int startRowCount = rowCount;
                        allocationSheet.Cells[startRowCount, 1].Value = alcn.requirement.Name;

                        foreach (Resource resource in alcn.resourceList)
                        {
                            allocationSheet.Cells[rowCount, 2].Value = resource.Name;

                            foreach (KeyValuePair<DateTime, bool> kvp in resource.AllocationMap)
                            {
                                var query3 = (from cell in allocationSheet.Cells["A:AZ"]
                                              where cell.Text.Equals(kvp.Key.ToString("dd-MMM"))                                                   
                                              select cell);
                                foreach (var cell in query3)    
                                {

                                    allocationSheet.Cells[rowCount, cell.Start.Column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    allocationSheet.Cells[rowCount, cell.Start.Column].Style.Fill.BackgroundColor.SetColor(color);
                                }
                            }

                            rowCount++;       
                        }

                        allocationSheet.Cells[startRowCount, 1, rowCount - 1, 1].Merge = true;
                        allocationSheet.Cells[startRowCount, 1, rowCount - 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                         
                    }

                    allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    package.Save();
                }
            }
        }

        private static DateTime AddWorkDays(DateTime date, int workingDays, params DateTime[] bankHolidays)
        {
            int direction = workingDays < 0 ? -1 : 1;
            DateTime newDate = date;
            // If a working day count of Zero is passed, return the date passed
            if (workingDays == 0)
            {

                newDate = date;
            }
            else
            {
                while (workingDays != -direction)
                {
                    if (newDate.DayOfWeek != DayOfWeek.Saturday &&
                        newDate.DayOfWeek != DayOfWeek.Sunday &&
                        Array.IndexOf(bankHolidays, newDate) < 0)
                    {
                        workingDays -= direction;
                    }
                    // if the original return date falls on a weekend or holiday, this will take it to the previous / next workday, but the "if" statement keeps it from going a day too far.

                    if (workingDays != -direction)
                    { newDate = newDate.AddDays(direction); }
                }
            }
            return newDate;
        }

        private static int BusinessDaysUntil(DateTime firstDay, DateTime lastDay, params DateTime[] bankHolidays)
        {
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            if (firstDay > lastDay)
                throw new ArgumentException("Incorrect last day " + lastDay);

            TimeSpan span = lastDay - firstDay;
            int businessDays = span.Days + 1;
            int fullWeekCount = businessDays / 7;
            // find out if there are weekends during the time exceedng the full weeks
            if (businessDays > fullWeekCount * 7)
            {
                // we are here to find out if there is a 1-day or 2-days weekend
                // in the time interval remaining after subtracting the complete weeks
                int firstDayOfWeek = (int)firstDay.DayOfWeek;
                int lastDayOfWeek = (int)lastDay.DayOfWeek;
                if (lastDayOfWeek < firstDayOfWeek)
                    lastDayOfWeek += 7;
                if (firstDayOfWeek <= 6)
                {
                    if (lastDayOfWeek >= 7)// Both Saturday and Sunday are in the remaining time interval
                        businessDays -= 2;
                    else if (lastDayOfWeek >= 6)// Only Saturday is in the remaining time interval
                        businessDays -= 1;
                }
                else if (firstDayOfWeek <= 7 && lastDayOfWeek >= 7)// Only Sunday is in the remaining time interval
                    businessDays -= 1;
            }

            // subtract the weekends during the full weeks in the interval
            businessDays -= fullWeekCount + fullWeekCount;

            // subtract the number of bank holidays during the time interval
            foreach (DateTime bankHoliday in bankHolidays)
            {
                DateTime bh = bankHoliday.Date;
                if (firstDay <= bh && bh <= lastDay)
                    --businessDays;
            }

            return businessDays;
        }

        private static void FindStartAndEndDate(List<Resource> resourceList, out DateTime StartDate, out DateTime EndDate)
        {
            List<DateTime> StartDateList = new List<DateTime>();
            StartDateList.Add((from d in resourceList select d.AvailableStartDate1).Min());
            StartDateList.Add((from d in resourceList select d.AvailableStartDate2).Min());
            StartDateList.Add((from d in resourceList select d.AvailableStartDate3).Min());
            StartDateList.Add((from d in resourceList select d.AvailableStartDate4).Min());
            StartDateList.Add((from d in resourceList select d.AvailableStartDate5).Min());

            StartDate = StartDateList.Where(q => q >= DateTime.Now).Min();

            List<DateTime> EndDateList = new List<DateTime>();
            EndDateList.Add((from d in resourceList select d.AvailableEndDate1).Min());
            EndDateList.Add((from d in resourceList select d.AvailableEndDate2).Min());
            EndDateList.Add((from d in resourceList select d.AvailableEndDate3).Min());
            EndDateList.Add((from d in resourceList select d.AvailableEndDate4).Min());
            EndDateList.Add((from d in resourceList select d.AvailableEndDate5).Min());

            EndDate = EndDateList.Where(q => q >= DateTime.Now).Min();
        }

        private static List<Requirement> ReadRequirement(ExcelWorksheet workSheet, List<Resource> resourceList)
        {
            List<Requirement> requirementList = new List<Requirement>();
            List<string> distinctResourceTypes = resourceList.Select(s => s.Type).Distinct().ToList<String>();
            int totalRows = workSheet.Dimension.End.Row;
            int totalCols = workSheet.Dimension.End.Column;
            for (int i = 1; i <= totalCols; i++)
            {
                string columnName = String.Format("{0}", workSheet.Cells[1, i].Text);
                if (distinctResourceTypes.Contains(columnName)) // Type Column
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        req.EffortByType.Add(columnName, Int32.Parse(workSheet.Cells[j, i].Text));

                    }

                }

                else if (columnName.Contains("Long Pole"))
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        req.LongPoleByType.Add(columnName.Replace("Long Pole", "").Trim(), Int32.Parse(workSheet.Cells[j, i].Text));

                    }

                }

                else if (columnName.Contains("Maximum"))
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        req.MaximimResourcesByType.Add(columnName.Replace("Maximum", "").Trim(), Int32.Parse(workSheet.Cells[j, i].Text));

                    }

                }

                else if (columnName.Contains("Priority"))
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        req.Priority = Int32.Parse(workSheet.Cells[j, i].Text);

                    }

                }

                else
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        req.Name = workSheet.Cells[j, i].Text;

                    }

                }
            }
            return requirementList;
        }

        private static List<Resource> ReadResource(ExcelWorksheet resourceSheet)
        {

            List<Resource> resourceList = new List<Resource>();
            int totalRows = resourceSheet.Dimension.End.Row;
            int totalCols = resourceSheet.Dimension.End.Column;
            for (int i = 2; i <= totalRows; i++)
            {
                Resource resource = new Resource();


                dynamic cellValue = String.Format("{0}", resourceSheet.Cells[i, 1].Text);
                resource.Name = cellValue;

                cellValue = String.Format("{0}", resourceSheet.Cells[i, 2].Text);
                resource.Type = cellValue;

                cellValue = DateTime.Parse(resourceSheet.Cells[i, 3].Text);
                resource.AvailableStartDate1 = cellValue;

                cellValue = DateTime.Parse(resourceSheet.Cells[i, 4].Text);
                resource.AvailableEndDate1 = cellValue;

                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 5].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 5].Text);
                    resource.AvailableStartDate2 = cellValue;
                }

                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 6].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 6].Text);
                    resource.AvailableEndDate2 = cellValue;
                }

                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 7].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 7].Text);
                    resource.AvailableStartDate3 = cellValue;
                }

                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 8].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 8].Text);
                    resource.AvailableEndDate3 = cellValue;
                }
                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 9].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 9].Text);
                    resource.AvailableStartDate4 = cellValue;
                }

                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 10].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 10].Text);
                    resource.AvailableEndDate4 = cellValue;
                }
                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 11].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 11].Text);
                    resource.AvailableStartDate5 = cellValue;
                }

                if (!string.IsNullOrEmpty(resourceSheet.Cells[i, 12].Text))
                {
                    cellValue = DateTime.Parse(resourceSheet.Cells[i, 12].Text);
                    resource.AvailableEndDate5 = cellValue;
                }
                resourceList.Add(resource);

            }
            return resourceList;
        }

        private static DataTable WorksheetToDataTable(ExcelWorksheet oSheet)
        {
            int totalRows = oSheet.Dimension.End.Row;
            int totalCols = oSheet.Dimension.End.Column;
            DataTable dt = new DataTable(oSheet.Name);
            DataRow dr = null;
            for (int i = 1; i <= totalRows; i++)
            {
                if (i > 1) dr = dt.Rows.Add();
                for (int j = 1; j <= totalCols; j++)
                {
                    if (i == 1)
                        dt.Columns.Add(oSheet.Cells[i, j].Value.ToString());
                    else
                    {
                        string myString = String.Format("{0}", oSheet.Cells[i, j].Text);
                        dr[j - 1] = myString;
                    }
                }
            }
            return dt;
        }

    }
}
