using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAllocatev2
{
    [Serializable]
    public class Allocation
    {
        public Requirement requirement { get; set; }
        public List<Resource> resourceList = new List<Resource>();
        public DateTime AllocationStartDate { get; set; }
        public DateTime AllocationEndDate { get; set; }
        public bool Allocated { get; set; }
        public Dictionary<string, int> AllocationByType = new Dictionary<string, int>();

        public override string ToString()
        {
            string value = string.Format("requirement = [{0}],resourceList=[{1}] AllocationStartDate =[{2}], AllocationEndDate = [{3}]," +
            "Allocated=[{4}]", requirement, string.Join(", ",
            resourceList.ConvertAll(m =>
                string.Format("'{0}'", m)).ToArray()), AllocationStartDate, AllocationEndDate, Allocated);
            return value;

        }


    }


}
