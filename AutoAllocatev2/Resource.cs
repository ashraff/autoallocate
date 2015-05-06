using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAllocatev2
{
    [Serializable]
    public class Resource
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public DateTime[] AvailableStartDate { get; set; }
        public DateTime[] AvailableEndDate { get; set; }
        public float[] Percentage { get; set; }
        /*public DateTime AvailableStartDate2 { get; set; }
        public DateTime AvailableEndDate2 { get; set; }
        public float Percentage2 { get; set; }
        public DateTime AvailableStartDate3 { get; set; }
        public DateTime AvailableEndDate3 { get; set; }
        public float Percentage3 { get; set; }
        public DateTime AvailableStartDate4 { get; set; }
        public DateTime AvailableEndDate4 { get; set; }
        public float Percentage4 { get; set; }
        public DateTime AvailableStartDate5 { get; set; }
        public DateTime AvailableEndDate5 { get; set; }
        public float Percentage5 { get; set; }*/
        public Dictionary<DateTime, float> AllocationMap = new Dictionary<DateTime, float>();
        public int AllocationTry { get; set; }

        public override string ToString()
        {
            string value = string.Format("Name = {0}, Type ={1}, AvailableStartDate1 = {2}," +
            "AvailableEndDate1={3}", Name, Type, DateFormatter(AvailableStartDate[0]), DateFormatter(AvailableEndDate[0]));
            return value;

        }

        private string DateFormatter(DateTime dateTime)
        {
            if (DateTime.Compare(dateTime, DateTime.MinValue) != 0)
            {
                return dateTime.ToShortDateString();
            }
            return string.Empty;
        }
    }
}
