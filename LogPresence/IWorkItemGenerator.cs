using System;
using System.Collections.Generic;

namespace LogPresence
{
    public interface IWorkItemGenerator
    {
        IEnumerable<WorkItemOnDay> GetWorkItemsOnDay(DateTime date, decimal totalHours);
    }
}