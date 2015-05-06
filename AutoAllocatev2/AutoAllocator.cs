namespace AutoAllocatev2
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    /* For Diagnostics */
    using System.Diagnostics;
    using System.Drawing;
    /* For I/O purpose */
    using System.IO;
    using System.Linq;
    using System.Runtime.Serialization.Formatters.Binary;

    /* To work eith EPPlus library */
    using OfficeOpenXml;

    public class AutoAllocator
    {
        #region Methods

        public static void Allocate(string args)
        {
            var File = new FileInfo(args);

            using (ExcelPackage package = new ExcelPackage(File))
            {
                ExcelWorkbook workBook = package.Workbook;
                if (workBook != null)
                {
                    ExcelWorksheet workSheet = workBook.Worksheets["Resources"];
                    List<Resource> resourceList = ReadResource(workSheet);

                    ExcelWorksheet reqSheet = workBook.Worksheets["Requirement"];
                    List<Requirement> requirementList = ReadRequirement(reqSheet, resourceList);

                    /* Sort Requirement By Priority*/
                    requirementList = requirementList.OrderBy(req => req.Priority).ToList();
                    requirementList = UpdateRequriment(requirementList);

                    /* Find Which Requirement has Longest Allocation and use it for first allocation Date */
                    int longestPole = (from x in requirementList select x.LongPoleByType.Aggregate((l, r) => l.Value > r.Value ? l : r)).Max(a => a.Value);

                    DateTime StartDate, EndDate;

                    FindStartAndEndDate(resourceList, out StartDate, out EndDate);
                    Console.WriteLine("The Allocation StartDate={0} and EndDate={1}", StartDate, EndDate);

                    List<Allocation> allocationList = new List<Allocation>();
                    Dictionary<DateTime, int> AllocationCountMap = new Dictionary<DateTime, int>();

                    /* Try to allocate within longPole and keep incrementing until you find and efficient allocation.*/
                    int daysUntil = BusinessDaysUntil(StartDate, EndDate);
                    DateTime newEndDate = AddWorkDays(StartDate, longestPole);

                    for (int i = longestPole; i <= daysUntil; i++)
                    {
                        /* Clone the Requirement List and Resource List */
                        List<Resource> originalResourceList = (List<Resource>)FromBinary(ToBinary(resourceList));
                        List<Requirement> originalRequirementList = (List<Requirement>)FromBinary(ToBinary(requirementList));

                        /* Update the Resource End Date to Match the New End Date Calculated. */
                        originalResourceList = UpdateResourceEndDate(originalResourceList, newEndDate, i);

                        /* Allocate Resource Now */
                        allocationList = AutoAllocateResource(originalResourceList, originalRequirementList, StartDate, newEndDate);
                        if (allocationList.Count(a => a.Allocated) == allocationList.Count)
                        {
                            /*All are allocated with least possible resource and date.*/
                            break;
                        }
                        else if (newEndDate.CompareTo(EndDate) <= 0)
                        {
                            /* May not be a efficient allocation ,so lets keep count of which date has what allocation count */
                            AllocationCountMap.Add(newEndDate, allocationList.Count(a => a.Allocated));
                            newEndDate = AddWorkDays(newEndDate, 1);
                        }
                        else if (newEndDate.CompareTo(EndDate) > 0)
                        {
                            /* Didnt find efficient allocation, so find a least date which has more allocation and use it as new date and return to the user */
                            DateTime newEndDate1 = AllocationCountMap.Where(a => a.Value == AllocationCountMap.Values.Max()).OrderBy(a => a.Key).First().Key;

                            /* Clone the Requirement List and Resource List Again.*/
                            originalResourceList = (List<Resource>)FromBinary(ToBinary(resourceList));
                            originalRequirementList = (List<Requirement>)FromBinary(ToBinary(requirementList));

                            originalResourceList = UpdateResourceEndDate(originalResourceList, newEndDate1, daysUntil);
                            allocationList = AutoAllocateResource(originalResourceList, originalRequirementList, StartDate, newEndDate);
                            break;
                        }
                        Console.WriteLine("Iteration {0} Allocated {1} requirement for Date {2}", (i - longestPole), AllocationCountMap.Values.Max(), newEndDate.ToShortDateString());
                    }
                    /* Print the Allocation Sheet */
                    foreach (Allocation alloc in allocationList)
                    {
                        Console.WriteLine(alloc.ToString());
                    }
                    /* Now we have the data, Create the Allocation sheet */
                    CreateAllocationWorkSheet(package, StartDate, EndDate, allocationList, resourceList);

                    package.Save();
                }
            }
        }

        private static DateTime AddWorkDays(DateTime date, int workingDays, params DateTime[] bankHolidays)
        {
            int direction = workingDays < 0 ? -1 : 1;
            DateTime newDate = date;
            /*If a working day count of Zero is passed, return the date passed */
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
                    /*f the original return date falls on a weekend or holiday, this will take it to the previous / next workday, but the "if" statement keeps it from going a day too far.*/

                    if (workingDays != -direction)
                    { newDate = newDate.AddDays(direction); }
                }
            }
            return newDate;
        }

        private static List<Allocation> AutoAllocateResource(List<Resource> availableResourceList, List<Requirement> requirementList, DateTime StartDate, DateTime EndDate)
        {
            /* This is the heart of the allocation program. It tries to allocate the resource against a requriement from the start state within the resource available end date efficiently. */
            List<Allocation> allocationList = new List<Allocation>();

            /* Loop till all the requirements are determined to be allocated or not allocated */
            do
            {
                /* Loop through each of the Requirement */
                foreach (Requirement requirement in requirementList)
                {
                    /* Find any dependent requirement for the current requirement. The expectation is requirements are incrementally dependent in the Requirement Sheet.*/
                    Allocation dependentRequirement = (from n in allocationList where n.requirement.Id == requirement.DependsOnId select n).FirstOrDefault();

                    /* If the current requirement doesnt depend on any Reqiurement or The dependent Requirement is already allocated then continue with the current requirement allocation */
                    if (dependentRequirement != null && dependentRequirement.Allocated || requirement.DependsOnId == requirement.Id)
                    {
                        /* If the current requiement is not allocated already then proceed*/
                        if (allocationList.Where(a => a.requirement.Name.Equals(requirement.Name)).Count() == 0)
                        {

                            Allocation allocation = new Allocation();
                            allocation.AllocationStartDate = StartDate;
                            allocation.AllocationEndDate = EndDate;
                            allocation.requirement = requirement;
                            allocation.AllocationByType = (Dictionary<string, int>)FromBinary(ToBinary(requirement.EffortByType));
                            List<Resource> originalResourceList = (List<Resource>)FromBinary(ToBinary(availableResourceList));

                            /* For the current requirement, loop through the effort by each of teh resource, we have to allocate one resource type at a time. */
                            foreach (KeyValuePair<string, int> kvp in requirement.EffortByType)
                            {
                                /* find maximum resources need for a type */
                                int maximumResource = requirement.MaximimResourcesByType[kvp.Key];
                                /* find the long pole for a type */
                                int longPole = requirement.LongPoleByType[kvp.Key];
                                /* Counter for retry logic, we are skipping resource if earlist one is available or if the resource is already allocated for same requirement due to efficiency reason. This retry will ensure we have tried fare amount of time */
                                int UnableToAllocate = 0;
                                /* Flag to skip the resource if the resource is already alloacted for the same requirement,till other resource are tried */
                                int AllocationTry = 1;

                                /* Till maximum resource for a allocation by type is reached and the requirement need allocation for the type and retry count didnt reach  N^2 count*/
                                while (allocation.resourceList.Where(a => a.Type.Equals(kvp.Key)).GroupBy(b => b.Name).Count() < maximumResource && allocation.AllocationByType[kvp.Key] > 0 && UnableToAllocate++ < (availableResourceList.Count * availableResourceList.Count))
                                {
                                    /* Loop through all the resource */
                                    foreach (Resource resource in availableResourceList)
                                    {
                                        /* if the resource type is same the required allocation type */
                                        if (kvp.Key.Equals(resource.Type))
                                        {
                                            /* Loop through the Available Dates*/
                                            for (int i = 0; i < resource.AvailableStartDate.Length; i++)
                                            {
                                                /* Compute new Start if we have a dependent which the current one depends on, this to ensure that this allocation start after it */
                                                DateTime tempStartDate = (dependentRequirement != null && IsBetween(dependentRequirement.AllocationEndDate, resource.AvailableStartDate[i], resource.AvailableEndDate[i])) ? dependentRequirement.AllocationEndDate : resource.AvailableStartDate[i];

                                                /* We are loopiong through several available ranges, if one of them is allocated, wwe use this flag to break out of the loop.*/
                                                bool isAllocatedInParticularDateRange = false;

                                                /* If we have free days for a resource. Intially longPole will be the longest day needed for the requriement. */
                                                if (longPole > 0 && BusinessDaysUntil(tempStartDate, resource.AvailableEndDate[i]) >= longPole)
                                                {
                                                    /* Skip the loop,of the resource is previously allocated for same requirement */
                                                    if (allocation.resourceList.Where(a => a.Type.Equals(kvp.Key) && a.Name.Equals(resource.Name) && a.AllocationTry == AllocationTry).Count() == 1)
                                                    {
                                                        continue;
                                                    }
                                                    /* This for evaulating resource dependency, if the current resource by type is depedent on same or other type and the other type is not within the dep.Days no of the days as mentioned in the requirement sheet, then allocation is not possible*/
                                                    if (allocation.resourceList.Count > 0)
                                                    {
                                                        Dependency dep = requirement.DependencyList.Where(a => a.Depender.Equals(allocation.resourceList.Last().Type) && a.Dependent.Equals(resource.Type)).FirstOrDefault();
                                                        if (dep != null && dep.Days != -1)
                                                        {
                                                            DateTime lastResourceStartDate = allocation.resourceList.Last().AllocationMap.First().Key;
                                                            DateTime dependentDate = AddWorkDays(lastResourceStartDate, dep.Days);
                                                            if (tempStartDate.CompareTo(dependentDate) > 0) continue;
                                                        }
                                                    }
                                                    /* Get all other resoure in after the current resource */
                                                    List<Resource> pendingList = availableResourceList.GetRange(availableResourceList.IndexOf(resource) + 1, availableResourceList.Count - availableResourceList.IndexOf(resource) - 1);

                                                    /* If there is no dependent requirement or if the  end of the requirement it depends on is between resource available date , then allocation is possible. */
                                                    if (dependentRequirement == null || IsBetween(dependentRequirement.AllocationEndDate, resource.AvailableStartDate[i], resource.AvailableEndDate[i]) || resource.AvailableStartDate[i].CompareTo(dependentRequirement.AllocationEndDate) >= 0)
                                                    {
                                                        /* Skip if other resource available with earlier available date */
                                                        if (pendingList.Where(a => a.AvailableStartDate[i].CompareTo(tempStartDate) < 0 && a.Type.Equals(kvp.Key)).Count() > 0)
                                                        {
                                                            continue;
                                                        }
                                                        /* Find if the same resource is used for allocation */
                                                        Resource availableResource = allocation.resourceList.Where(a => a.Name.Equals(resource.Name)).FirstOrDefault();

                                                        /* Create a New Resource */
                                                        if (availableResource == null)
                                                        {
                                                            availableResource = new Resource();
                                                            availableResource.AvailableStartDate = new DateTime[resource.AvailableStartDate.Length];
                                                            availableResource.AvailableEndDate = new DateTime[resource.AvailableStartDate.Length];
                                                            availableResource.Percentage = new float[resource.AvailableStartDate.Length];
                                                            allocation.resourceList.Add(availableResource);
                                                            availableResource.AllocationTry = AllocationTry;
                                                        }
                                                        availableResource.Name = resource.Name;
                                                        availableResource.Type = kvp.Key;
                                                        availableResource.AvailableStartDate[i] = tempStartDate;
                                                        availableResource.AvailableEndDate[i] = AddWorkDays(tempStartDate, longPole);

                                                        /* Reduce the allocation by type ,else is not possble at all,just in case*/
                                                        if (allocation.AllocationByType[kvp.Key] >= longPole)
                                                            allocation.AllocationByType[kvp.Key] = Convert.ToInt32(allocation.AllocationByType[kvp.Key] - (longPole * resource.Percentage[i]));
                                                        else allocation.AllocationByType[kvp.Key] = Convert.ToInt32((longPole * resource.Percentage[i]) - allocation.AllocationByType[kvp.Key]);

                                                        /* Reduce the longPole , which will be teh remainign allocation */
                                                        if (longPole > allocation.AllocationByType[kvp.Key]) longPole = allocation.AllocationByType[kvp.Key];

                                                        /* Set teh start date of the resource to the end date of current allocation */
                                                        resource.AvailableStartDate[i] = availableResource.AvailableEndDate[i];

                                                        /*Mark the allocation with a map, used to fill teh excel sheet */
                                                        DateTime tempDate = tempStartDate;
                                                        while (DateTime.Compare(tempDate, availableResource.AvailableEndDate[i]) != 0)
                                                        {
                                                            availableResource.AllocationMap.Add(tempDate, resource.Percentage[i]);
                                                            tempDate = tempDate.AddDays(1);
                                                        }

                                                        allocation.AllocationEndDate = tempDate;
                                                        /* Set the allocation flag to break out of the loop */
                                                        isAllocatedInParticularDateRange = true;
                                                    }
                                                }
                                                if (isAllocatedInParticularDateRange) break;
                                            }
                                        }
                                    }
                                    AllocationTry += 1;
                                }
                            }
                            /* After all tries, if we cannot allocate */
                            if (allocation.AllocationByType.Sum(x => x.Value) > 0) { availableResourceList = originalResourceList; allocation.Allocated = false; allocation.resourceList = new List<Resource>(); }
                            else { allocation.Allocated = true; }
                            allocationList.Add(allocation);

                        }
                    }
                    /* If Dependent Requirement is not allocated, dont allocate this requirement */
                    else
                    {
                        Allocation allocation = new Allocation();
                        allocation.AllocationStartDate = StartDate;
                        allocation.AllocationEndDate = EndDate;
                        allocation.requirement = requirement;
                        allocation.Allocated = false;
                        allocationList.Add(allocation);
                    }
                }
            } while (allocationList.Count != requirementList.Count);
            return allocationList;
        }

        private static int BusinessDaysUntil(DateTime firstDay, DateTime lastDay, params DateTime[] bankHolidays)
        {
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            if (firstDay.CompareTo(lastDay) >= 0)
                return 0;

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

        private static void CreateAllocationWorkSheet(ExcelPackage package, DateTime StartDate, DateTime EndDate, List<Allocation> allocationList, List<Resource> resourceList)
        {
            var allocationSheet = package.Workbook.Worksheets.Add("Allocation");

            allocationSheet.Column(1).Width = 12;
            allocationSheet.Column(1).Style.Font.Size = 8;
            allocationSheet.Column(2).Width = 12;
            allocationSheet.Column(2).Style.Font.Size = 8;
            allocationSheet.Column(3).Width = 12;
            allocationSheet.Column(3).Style.Font.Size = 8;
            allocationSheet.Column(4).Width = 8;
            allocationSheet.Column(4).Style.Font.Size = 8;
            allocationSheet.Cells[1, 1].Value = "Requirement";
            allocationSheet.Cells[1, 1, 2, 1].Merge = true;

            allocationSheet.Cells[1, 2].Value = "Resource";
            allocationSheet.Cells[1, 2, 2, 2].Merge = true;

            allocationSheet.Cells[1, 3].Value = "Type";
            allocationSheet.Cells[1, 3, 2, 3].Merge = true;

            allocationSheet.Cells[1, 4].Value = "Availability";
            allocationSheet.Cells[1, 4, 2, 4].Merge = true;

            int cellCount = 5;
            while (DateTime.Compare(StartDate, EndDate) <= 0)
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

            allocationSheet.Cells[1, cellCount].Value = "Allocated";
            allocationSheet.Cells[1, cellCount, 2, cellCount].Merge = true;
            allocationSheet.Column(cellCount).Width = 8;
            allocationSheet.Column(cellCount).Style.Font.Size = 8;

            allocationSheet.Cells[1, cellCount + 1].Value = "Inefficiency";
            allocationSheet.Cells[1, cellCount + 1, 2, cellCount + 1].Merge = true;
            allocationSheet.Column(cellCount + 1).Width = 8;
            allocationSheet.Column(cellCount + 1).Style.Font.Size = 8;

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
                    allocationSheet.Cells[rowCount, 3].Value = resource.Type;
                    Resource res = resourceList.Where(a => a.Name.Equals(resource.Name)).FirstOrDefault();
                    int totalAvailability = 0;
                    if (res != null)
                    {

                        for (int i = 0; i < res.AvailableStartDate.Length; i++)
                        {
                            totalAvailability += BusinessDaysUntil(res.AvailableStartDate[i], res.AvailableEndDate[i]);
                            DateTime actualAvailableStartDate = res.AvailableStartDate[i];
                            while (actualAvailableStartDate != DateTime.MinValue && actualAvailableStartDate <= resource.AvailableStartDate[i])
                            {
                                if (!resource.AllocationMap.ContainsKey(actualAvailableStartDate))
                                {
                                    var unAllocatedCells = (from cell in allocationSheet.Cells["A:IZ"]
                                                            where cell.Text.Equals(actualAvailableStartDate.ToString("dd-MMM"))
                                                            select cell);
                                    foreach (var cell in unAllocatedCells)
                                    {
                                        allocationSheet.Cells[rowCount, cell.Start.Column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.MediumGray;
                                        allocationSheet.Cells[rowCount, cell.Start.Column].Style.Fill.BackgroundColor.SetColor(color);

                                    }
                                }
                                actualAvailableStartDate = AddWorkDays(actualAvailableStartDate, 1);

                            }
                        }
                        allocationSheet.Cells[rowCount, 4].Value = totalAvailability;

                    }
                    float allocationDayCount = 0;
                    foreach (KeyValuePair<DateTime, float> kvp in resource.AllocationMap)
                    {
                        var allocatedCells = (from cell in allocationSheet.Cells["A:IZ"]
                                              where cell.Text.Equals(kvp.Key.ToString("dd-MMM"))
                                              select cell);
                        foreach (var cell in allocatedCells)
                        {
                            allocationSheet.Cells[rowCount, cell.Start.Column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            allocationSheet.Cells[rowCount, cell.Start.Column].Style.Fill.BackgroundColor.SetColor(color);
                            allocationSheet.Cells[rowCount, cell.Start.Column].Value = kvp.Value;
                            allocationSheet.Cells[rowCount, cell.Start.Column].Style.Font.Color.SetColor(color);
                            allocationDayCount += kvp.Value;
                        }
                    }

                    allocationSheet.Cells[rowCount, cellCount].Value = allocationDayCount;
                    allocationSheet.Cells[rowCount, cellCount + 1].Value = totalAvailability - allocationDayCount;

                    rowCount++;
                }
                if (rowCount > startRowCount)
                {
                    allocationSheet.Cells[startRowCount, 1, rowCount - 1, 1].Merge = true;
                    allocationSheet.Cells[startRowCount, 1, rowCount - 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }

            }

            cellCount++;
            allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            allocationSheet.Cells[1, 1, rowCount, cellCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            return;
        }

        private static void FindStartAndEndDate(List<Resource> resourceList, out DateTime StartDate, out DateTime EndDate)
        {
            List<DateTime> StartDateList = new List<DateTime>();

            StartDateList.AddRange(resourceList.SelectMany(a => a.AvailableStartDate).ToList());

            StartDate = StartDateList.Where(q => q >= DateTime.Now).Min();

            List<DateTime> EndDateList = new List<DateTime>();
            EndDateList.AddRange(resourceList.SelectMany(a => a.AvailableEndDate).ToList());

            EndDate = EndDateList.Where(q => q >= DateTime.Now).Max();
        }

        private static object FromBinary(Byte[] buffer)
        {
            MemoryStream ms = null;
            object deserializedObject = null;

            try
            {
                BinaryFormatter serializer = new BinaryFormatter();
                ms = new MemoryStream();
                ms.Write(buffer, 0, buffer.Length);
                ms.Position = 0;
                deserializedObject = serializer.Deserialize(ms);
            }
            finally
            {
                if (ms != null)
                    ms.Close();
            }
            return deserializedObject;
        }

        private static bool IsBetween<T>(T item, T start, T end)
        {
            return Comparer<T>.Default.Compare(item, start) >= 0
                && Comparer<T>.Default.Compare(item, end) <= 0;
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
                else if (columnName.Contains(":"))
                {
                    string[] depedency = columnName.Split(new string[] { ":" }, StringSplitOptions.None);
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        Dependency dependency = new Dependency();
                        dependency.Depender = depedency[0];
                        dependency.Dependent = depedency[1];
                        dependency.Days = Int32.Parse(workSheet.Cells[j, i].Text);
                        req.DependencyList.Add(dependency);

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

                else if (columnName.Contains("Dependency"))
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        if (workSheet.Cells[j, i] != null && !string.IsNullOrEmpty(workSheet.Cells[j, i].Text))
                            req.DependsOnId = Int32.Parse(workSheet.Cells[j, i].Text);
                        else req.DependsOnId = j - 1; // Dont depend on any task

                    }

                }

                else
                {
                    for (int j = 2; j <= totalRows; j++)
                    {
                        Requirement req = requirementList.ElementAtOrDefault(j - 2);
                        if (req == null) { req = new Requirement(); requirementList.Add(req); }
                        req.Name = workSheet.Cells[j, i].Text;
                        req.Id = j - 1;

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
                resource.AvailableStartDate = new DateTime[(totalCols - 2) / 3];
                resource.AvailableEndDate = new DateTime[(totalCols - 2) / 3];
                resource.Percentage = new float[(totalCols - 2) / 3];
                for (int j = 3; j <= totalCols; j += 3)
                {

                    if (!string.IsNullOrEmpty(resourceSheet.Cells[i, j].Text))
                    {
                        cellValue = DateTime.Parse(resourceSheet.Cells[i, j].Text);
                        resource.AvailableStartDate[j / 3 - 1] = cellValue;
                    }

                    if (!string.IsNullOrEmpty(resourceSheet.Cells[i, j + 1].Text))
                    {
                        cellValue = DateTime.Parse(resourceSheet.Cells[i, j + 1].Text);
                        resource.AvailableEndDate[j / 3 - 1] = cellValue;
                    }

                    if (!string.IsNullOrEmpty(resourceSheet.Cells[i, j + 2].Text))
                    {
                        cellValue = float.Parse(resourceSheet.Cells[i, j + 2].Text);
                        resource.Percentage[j / 3 - 1] = cellValue;
                    }
                }
                resourceList.Add(resource);

            }
            return resourceList;
        }

        private static Byte[] ToBinary(Object obj)
        {
            MemoryStream ms = null;
            Byte[] byteArray = null;
            try
            {
                BinaryFormatter serializer = new BinaryFormatter();
                ms = new MemoryStream();
                serializer.Serialize(ms, obj);
                byteArray = ms.ToArray();
            }
            catch (Exception unexpected)
            {
                Trace.Fail(unexpected.Message);
                throw;
            }
            finally
            {
                if (ms != null)
                    ms.Close();
            }
            return byteArray;
        }

        private static List<Requirement> UpdateRequriment(List<Requirement> requirementList)
        {
            foreach (Requirement re in requirementList)
            {

                List<string> typeArray = new List<string>();
                int index = 0;
                foreach (Dependency d in re.DependencyList)
                {
                    if (d.Days > 0 && !d.Depender.Equals(d.Dependent))
                    {
                        if (typeArray.Contains(d.Depender))
                        {
                            typeArray.Insert(typeArray.IndexOf(d.Depender) + 1, d.Dependent);
                        }
                        else if (typeArray.Contains(d.Dependent))
                        {
                            typeArray.Insert(typeArray.IndexOf(d.Depender) - 1, d.Depender);
                        }
                        else
                        {
                            typeArray.Insert(index, d.Depender);
                            typeArray.Insert(++index, d.Dependent);
                            index++;
                        }
                    }
                }
                Dictionary<string, int> newEffortByType = new Dictionary<string, int>();
                foreach (string type in typeArray)
                {
                    newEffortByType.Add(type, re.EffortByType[type]);
                    re.EffortByType.Remove(type);
                }

                foreach (KeyValuePair<string, int> type in re.EffortByType)
                {
                    newEffortByType.Add(type.Key, type.Value);
                }

                re.EffortByType = newEffortByType;
            }
            return requirementList;
        }

        private static List<Resource> UpdateResourceEndDate(List<Resource> originalResourceList, DateTime newEndDate, int longestPole)
        {
            foreach (Resource re in originalResourceList)
            {
                int arraySize = re.AvailableStartDate.Length;
                int businessDayUntil = 0;
                for (int i = 0; i < arraySize; i++)
                {
                    businessDayUntil += BusinessDaysUntil(re.AvailableStartDate[i], re.AvailableEndDate[i]);
                    if (businessDayUntil > longestPole)
                    {
                        re.AvailableEndDate[i] = AddWorkDays(re.AvailableStartDate[i], i == 0 ? businessDayUntil : (businessDayUntil - longestPole));
                        for (int j = i + 1; j < arraySize; j++)
                            re.AvailableStartDate[j] = re.AvailableEndDate[j] = DateTime.MinValue;
                        break;
                    }
                }
            }
            return originalResourceList;
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

        #endregion Methods
    }
}