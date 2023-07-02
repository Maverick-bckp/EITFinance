using System;

namespace EITFinance.Models.Common
{
    public abstract class BaseEntity
    {
        public int Id { get; set; }
        public DateTime cretedDate { get; set; }
        public int createdBy { get; set; }
        public DateTime modifyDate { get; set; }
        public int modifyBy { get; set; }
    }
}
