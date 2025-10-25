using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPoc.Contracts.DTO
{
    public class ExcelDto
    {
        public string? Name { get; set; }
        public string? Age { get; set; }
        public string? Gender { get; set; }
        public string? Email { get; set; }
        public List<string>? Preferences { get; set; }
        public string? Remarks { get; set; }
        public string? LastName { get; set; }
        public string? OrganizationName { get; set; }
        public string? BirthDate { get; set; }
        public string? Medicare { get; set; }
        public string? AssessmentCompletedBy { get; set; }
        public string? PrintArea { get; set; }
        public string? MedicationsTtlNum { get; set; }
    }
}
