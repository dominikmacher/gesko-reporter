using System;
using System.Collections.Generic;
using System.Text;

namespace GeskoReporter
{
    public class CallRecord
    {
        public string phoneId;
        public string phoneName;
        public string phoneNumber;
        public string date;
        public string time;
        public string duration;
        public string phoneUnits;
        public string cost;

        public CallRecord(string phoneId, string phoneName, string phoneNumber, string date, string time, string duration, string phoneUnits, string cost)
        {
            this.phoneId = phoneId;
            this.phoneName = phoneName;
            this.phoneNumber = phoneNumber;
            this.date = date;
            this.time = time;
            this.duration = duration;
            this.phoneUnits = phoneUnits;
            this.cost = cost;
        }
        public CallRecord()
        { }
    }
}
